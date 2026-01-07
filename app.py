import logging
import os
import time
import threading
import xml.etree.ElementTree as ET
from flask import Flask, request
from pathlib import Path
from config import config
from converter import DocumentConverter
from wecom_api import WeComAPI

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
converter = DocumentConverter()
wecom_api = WeComAPI()

# 用于防止重复处理的消息缓存
processed_messages = {}
MESSAGE_CACHE_TTL = 60  # 缓存60秒


def cleanup_message_cache():
    """清理过期的消息缓存"""
    current_time = time.time()
    expired_keys = [
        k for k, v in processed_messages.items() 
        if current_time - v > MESSAGE_CACHE_TTL
    ]
    for key in expired_keys:
        del processed_messages[key]


def create_text_response(to_user: str, from_user: str, content: str) -> str:
    """创建文本消息回复XML（明文，需要后续加密）"""
    return f"""<xml>
<ToUserName><![CDATA[{to_user}]]></ToUserName>
<FromUserName><![CDATA[{from_user}]]></FromUserName>
<CreateTime>{int(time.time())}</CreateTime>
<MsgType><![CDATA[text]]></MsgType>
<Content><![CDATA[{content}]]></Content>
</xml>"""


def process_document_async(from_user: str, media_id: str, file_name: str):
    """
    异步处理文档转换
    
    由于企业微信要求5秒内回复，而转换可能需要更长时间，
    所以使用异步处理，通过应用消息接口发送结果
    """
    input_file = None
    output_pdf = None
    
    try:
        # 确定文件扩展名
        file_ext = Path(file_name).suffix
        if not file_ext:
            file_ext = '.docx'  # 默认扩展名
        
        # 下载文件
        timestamp_ms = int(time.time() * 1000)
        input_file = os.path.join(config.TEMP_DIR, f"input_{timestamp_ms}{file_ext}")
        wecom_api.download_media(media_id, input_file)
        
        # 转换为PDF
        output_pdf = converter.convert_to_pdf(input_file)
        
        # 上传PDF到企业微信
        pdf_media_id = wecom_api.upload_media(output_pdf, 'file')
        
        # 发送PDF文件给用户
        success = wecom_api.send_file_message(from_user, pdf_media_id)
        
        if not success:
            wecom_api.send_text_message(
                from_user, 
                "⚠️ PDF生成成功但发送失败，请稍后重试。"
            )
        
    except Exception as e:
        logger.error(f"处理文档失败: {str(e)}")
        error_msg = f"❌ 转换失败: {str(e)}\n\n支持格式: Word(.doc/.docx), Excel(.xls/.xlsx), PPT(.ppt/.pptx)"
        wecom_api.send_text_message(from_user, error_msg)
    
    finally:
        # 清理临时文件
        if input_file:
            converter.cleanup_file(input_file)
        if output_pdf:
            converter.cleanup_file(output_pdf)


@app.route('/wecom', methods=['GET', 'POST'])
def wecom_handler():
    """企业微信消息处理器"""
    
    msg_signature = request.args.get('msg_signature', '')
    timestamp = request.args.get('timestamp', '')
    nonce = request.args.get('nonce', '')
    
    # GET请求：回调URL验证
    if request.method == 'GET':
        echostr = request.args.get('echostr', '')
        
        try:
            reply_echostr = wecom_api.crypto.verify_url(msg_signature, timestamp, nonce, echostr)
            logger.info("企业微信回调URL验证成功")
            return reply_echostr
        except Exception as e:
            logger.error(f"企业微信回调URL验证失败: {str(e)}")
            return 'Verification failed', 403
    
    # POST请求：处理消息
    try:
        xml_data = request.data.decode('utf-8')
        logger.info(f"收到加密消息: {xml_data[:200]}...")
        
        # 解析外层XML获取加密内容
        root = ET.fromstring(xml_data)
        encrypt_elem = root.find('Encrypt')
        if encrypt_elem is None:
            logger.error("消息中没有Encrypt字段")
            return 'success'
        
        encrypt_msg = encrypt_elem.text
        
        # 解密消息
        decrypted_xml = wecom_api.crypto.decrypt_message(msg_signature, timestamp, nonce, encrypt_msg)
        logger.info(f"解密后消息: {decrypted_xml[:200]}...")
        
        # 解析解密后的XML
        msg_root = ET.fromstring(decrypted_xml)
        
        msg_type = msg_root.find('MsgType').text
        from_user = msg_root.find('FromUserName').text  # 用户的userid
        to_user = msg_root.find('ToUserName').text      # 企业的corpid
        msg_id = msg_root.find('MsgId')
        msg_id = msg_id.text if msg_id is not None else str(time.time())
        
        # 清理过期缓存
        cleanup_message_cache()
        
        # 检查是否重复消息
        if msg_id in processed_messages:
            logger.info(f"跳过重复消息: {msg_id}")
            return 'success'
        
        # 标记消息已处理
        processed_messages[msg_id] = time.time()
        
        # 处理文件消息
        if msg_type == 'file':
            media_id = msg_root.find('MediaId').text
            file_name_elem = msg_root.find('FileName')
            file_name = file_name_elem.text if file_name_elem is not None else 'document.docx'
            
            logger.info(f"收到文件: {file_name}, MediaId: {media_id}")
            
            # 启动异步处理线程
            thread = threading.Thread(
                target=process_document_async,
                args=(from_user, media_id, file_name)
            )
            thread.daemon = True
            thread.start()
            
            # 创建回复消息
            reply_msg = create_text_response(from_user, to_user, "📄 正在转换您的文档，请稍候...\n预计需要5-15秒")
            # 加密回复
            encrypted_reply = wecom_api.crypto.encrypt_message(reply_msg, nonce, timestamp)
            return encrypted_reply
        
        # 处理文本消息
        elif msg_type == 'text':
            content = msg_root.find('Content').text or ''
            
            if content.strip() in ['帮助', 'help', '?', '？', 'h']:
                help_text = """📄 作业排版助手使用说明

1️⃣ 直接发送Word/Excel/PPT文件
2️⃣ 等待5-15秒，自动收到PDF
3️⃣ 转发PDF给打印机

✅ 支持格式: Word, Excel, PowerPoint
⏱️ 转换时间: 通常5-15秒
📱 完美还原Windows排版！"""
                reply_msg = create_text_response(from_user, to_user, help_text)
            else:
                reply_msg = create_text_response(
                    from_user, to_user, 
                    "请发送Word/Excel文件，我会帮您转换为PDF 📄\n\n发送「帮助」查看使用说明"
                )
            
            encrypted_reply = wecom_api.crypto.encrypt_message(reply_msg, nonce, timestamp)
            return encrypted_reply
        
        # 其他消息类型
        else:
            reply_msg = create_text_response(
                from_user, to_user, 
                "请发送Word或Excel文件，我会帮您转换为PDF 📄"
            )
            encrypted_reply = wecom_api.crypto.encrypt_message(reply_msg, nonce, timestamp)
            return encrypted_reply
            
    except Exception as e:
        logger.error(f"处理消息异常: {str(e)}", exc_info=True)
        return 'success'


@app.route('/health', methods=['GET'])
def health_check():
    """健康检查接口"""
    return {'status': 'ok', 'service': 'wecom-doc-converter'}


@app.route('/', methods=['GET'])
def index():
    """根路径"""
    return {
        'message': 'Enterprise WeChat Document Converter Service',
        'health': '/health',
        'wecom': '/wecom'
    }


if __name__ == '__main__':
    # 确保临时目录存在
    os.makedirs(config.TEMP_DIR, exist_ok=True)
    
    app.run(
        host=config.HOST,
        port=config.PORT,
        debug=config.DEBUG,
        threaded=True
    )
