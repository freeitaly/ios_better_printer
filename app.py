import hashlib
import logging
import os
import time
import threading
from flask import Flask, request
from pathlib import Path
from config import config
from converter import DocumentConverter
from wechat_api import WeChatAPI

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
converter = DocumentConverter()
wechat_api = WeChatAPI()

# 用于防止重复处理的消息缓存
processed_messages = {}
MESSAGE_CACHE_TTL = 60  # 缓存60秒

def verify_signature(signature: str, timestamp: str, nonce: str) -> bool:
    """验证微信服务器签名"""
    token = config.WECHAT_TOKEN
    tmp_list = [token, timestamp, nonce]
    tmp_list.sort()
    tmp_str = ''.join(tmp_list)
    tmp_hash = hashlib.sha1(tmp_str.encode('utf-8')).hexdigest()
    return tmp_hash == signature

def create_text_response(to_user: str, from_user: str, content: str) -> str:
    """创建文本消息回复XML"""
    return f"""<xml>
<ToUserName><![CDATA[{to_user}]]></ToUserName>
<FromUserName><![CDATA[{from_user}]]></FromUserName>
<CreateTime>{int(time.time())}</CreateTime>
<MsgType><![CDATA[text]]></MsgType>
<Content><![CDATA[{content}]]></Content>
</xml>"""

def cleanup_message_cache():
    """清理过期的消息缓存"""
    current_time = time.time()
    expired_keys = [
        k for k, v in processed_messages.items() 
        if current_time - v > MESSAGE_CACHE_TTL
    ]
    for key in expired_keys:
        del processed_messages[key]

def process_document_async(from_user: str, media_id: str, file_name: str):
    """
    异步处理文档转换
    
    由于微信要求5秒内回复，而转换可能需要更长时间，
    所以使用异步处理，通过客服消息接口发送结果
    """
    input_file = None
    output_pdf = None
    
    try:
        # 确定文件扩展名
        file_ext = Path(file_name).suffix
        if not file_ext:
            file_ext = '.docx'  # 默认扩展名
        
        # 下载文件
        timestamp = int(time.time() * 1000)
        input_file = os.path.join(config.TEMP_DIR, f"input_{timestamp}{file_ext}")
        wechat_api.download_media(media_id, input_file)
        
        # 转换为PDF
        output_pdf = converter.convert_to_pdf(input_file)
        
        # 上传PDF到微信
        pdf_media_id = wechat_api.upload_media(output_pdf, 'file')
        
        # 通过客服消息发送PDF（被动回复不支持file类型）
        success = wechat_api.send_file_message(from_user, pdf_media_id)
        
        if not success:
            # 如果客服消息发送失败，尝试发送文本提示
            wechat_api.send_text_message(
                from_user, 
                "⚠️ PDF生成成功但发送失败，请稍后重试。\n\n"
                "提示：如果反复失败，可能是公众号未开通客服消息权限。"
            )
        
    except Exception as e:
        logger.error(f"处理文档失败: {str(e)}")
        # 发送错误消息
        error_msg = f"❌ 转换失败: {str(e)}\n\n支持格式: Word(.doc/.docx), Excel(.xls/.xlsx), PPT(.ppt/.pptx)"
        wechat_api.send_text_message(from_user, error_msg)
    
    finally:
        # 清理临时文件
        if input_file:
            converter.cleanup_file(input_file)
        if output_pdf:
            converter.cleanup_file(output_pdf)

@app.route('/wechat', methods=['GET', 'POST'])
def wechat_handler():
    """微信消息处理器"""
    
    # GET请求：微信服务器验证
    if request.method == 'GET':
        signature = request.args.get('signature', '')
        timestamp = request.args.get('timestamp', '')
        nonce = request.args.get('nonce', '')
        echostr = request.args.get('echostr', '')
        
        if verify_signature(signature, timestamp, nonce):
            logger.info("微信服务器验证成功")
            return echostr
        else:
            logger.warning("微信服务器验证失败")
            return 'Invalid signature', 403
    
    # POST请求：处理消息
    try:
        xml_data = request.data.decode('utf-8')
        logger.info(f"收到消息: {xml_data[:200]}...")
        
        # 解析XML
        import xml.etree.ElementTree as ET
        root = ET.fromstring(xml_data)
        
        msg_type = root.find('MsgType').text
        from_user = root.find('FromUserName').text
        to_user = root.find('ToUserName').text
        msg_id = root.find('MsgId')
        msg_id = msg_id.text if msg_id is not None else str(time.time())
        
        # 清理过期缓存
        cleanup_message_cache()
        
        # 检查是否重复消息（微信可能重试）
        if msg_id in processed_messages:
            logger.info(f"跳过重复消息: {msg_id}")
            return 'success'
        
        # 标记消息已处理
        processed_messages[msg_id] = time.time()
        
        # 处理文件消息
        if msg_type == 'file':
            media_id = root.find('MediaId').text
            file_name_elem = root.find('FileName')
            file_name = file_name_elem.text if file_name_elem is not None else 'document.docx'
            
            logger.info(f"收到文件: {file_name}, MediaId: {media_id}")
            
            # 立即回复"处理中"消息，避免5秒超时
            # 然后异步处理文档，完成后通过客服消息发送
            processing_msg = "📄 正在转换您的文档，请稍候...\n预计需要5-15秒"
            
            # 启动异步处理线程
            thread = threading.Thread(
                target=process_document_async,
                args=(from_user, media_id, file_name)
            )
            thread.daemon = True
            thread.start()
            
            return create_text_response(from_user, to_user, processing_msg)
        
        # 处理文本消息
        elif msg_type == 'text':
            content = root.find('Content').text or ''
            
            if content.strip() in ['帮助', 'help', '?', '？', 'h']:
                help_text = """📄 作业排版助手使用说明

1️⃣ 在班级群里长按作业文件
2️⃣ 选择"转发" → 转发给本公众号
3️⃣ 等待5-15秒，自动收到PDF
4️⃣ 转发PDF给打印机小程序

✅ 支持格式: Word, Excel, PowerPoint
⏱️ 转换时间: 通常5-15秒
📱 完美还原Windows排版，告别打印错乱！"""
                return create_text_response(from_user, to_user, help_text)
            else:
                return create_text_response(
                    from_user, to_user, 
                    "请直接转发Word/Excel文件给我，我会帮您转换为PDF 📄\n\n发送「帮助」查看使用说明"
                )
        
        # 其他消息类型
        else:
            return create_text_response(
                from_user, to_user, 
                "请发送Word或Excel文件，我会帮您转换为PDF 📄"
            )
            
    except Exception as e:
        logger.error(f"处理消息异常: {str(e)}", exc_info=True)
        return 'success'  # 返回success避免微信重试

@app.route('/health', methods=['GET'])
def health_check():
    """健康检查接口"""
    return {'status': 'ok', 'service': 'wechat-doc-converter'}

@app.route('/', methods=['GET'])
def index():
    """根路径"""
    return {'message': 'WeChat Document Converter Service', 'health': '/health', 'wechat': '/wechat'}

if __name__ == '__main__':
    # 确保临时目录存在
    os.makedirs(config.TEMP_DIR, exist_ok=True)
    
    app.run(
        host=config.HOST,
        port=config.PORT,
        debug=config.DEBUG,
        threaded=True  # 启用多线程处理
    )
