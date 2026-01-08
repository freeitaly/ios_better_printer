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
    
    # ========== 第一步：记录原始请求（在任何处理之前）==========
    logger.info("=" * 60)
    logger.info(f"[DEBUG] 收到请求: Method={request.method}, URL={request.url}")
    logger.info(f"[DEBUG] Remote IP: {request.remote_addr}")
    logger.info(f"[DEBUG] Content-Type: {request.content_type}")
    logger.info(f"[DEBUG] Content-Length: {request.content_length}")
    
    # 对于POST请求，立即记录原始body（用于判断是否收到了请求）
    if request.method == 'POST':
        raw_body = request.data
        logger.info(f"[DEBUG] 原始Body长度: {len(raw_body)} 字节")
        logger.info(f"[DEBUG] 原始Body内容: {raw_body[:1000]}")  # 记录前1000字节
    
    msg_signature = request.args.get('msg_signature', '')
    timestamp = request.args.get('timestamp', '')
    nonce = request.args.get('nonce', '')
    
    # GET请求：回调URL验证
    if request.method == 'GET':
        echostr = request.args.get('echostr', '')
        
        try:
            reply_echostr = wecom_api.crypto.verify_url(msg_signature, timestamp, nonce, echostr)
            logger.info("[SUCCESS] 企业微信回调URL验证成功")
            return reply_echostr
        except Exception as e:
            logger.error(f"[ERROR] 企业微信回调URL验证失败: {str(e)}")
            return 'Verification failed', 403
    
    # POST请求：处理消息
    try:
        xml_data = request.data.decode('utf-8')
        
        # 解析外层XML获取加密内容
        root = ET.fromstring(xml_data)
        encrypt_elem = root.find('Encrypt')
        if encrypt_elem is None:
            logger.error("[ERROR] 消息中没有Encrypt字段")
            return 'success'
        
        encrypt_msg = encrypt_elem.text
        
        # 解密消息
        decrypted_xml = wecom_api.crypto.decrypt_message(msg_signature, timestamp, nonce, encrypt_msg)
        
        # ========== 关键：记录解密后的完整XML ==========
        logger.info(f"[DEBUG] 解密后完整XML:\n{decrypted_xml}")
        
        # 解析解密后的XML
        msg_root = ET.fromstring(decrypted_xml)
        
        # 获取消息类型
        msg_type_elem = msg_root.find('MsgType')
        msg_type = msg_type_elem.text if msg_type_elem is not None else 'unknown'
        
        logger.info(f"[DEBUG] >>>>>> 消息类型: {msg_type} <<<<<<")
        
        from_user = msg_root.find('FromUserName').text  # 用户的userid
        to_user = msg_root.find('ToUserName').text      # 企业的corpid
        msg_id = msg_root.find('MsgId')
        msg_id = msg_id.text if msg_id is not None else str(time.time())
        
        logger.info(f"[DEBUG] FromUser={from_user}, ToUser={to_user}, MsgId={msg_id}")
        
        # 清理过期缓存
        cleanup_message_cache()
        
        # 检查是否重复消息
        if msg_id in processed_messages:
            logger.info(f"[SKIP] 跳过重复消息: {msg_id}")
            return 'success'
        
        # 标记消息已处理
        processed_messages[msg_id] = time.time()
        
        # ========== 处理文件消息 ==========
        if msg_type == 'file':
            logger.info("[FILE] 检测到文件类型消息，开始处理...")
            
            media_id = msg_root.find('MediaId').text
            
            # 企业微信file消息可能用不同的字段名：FileName 或 Title
            file_name = None
            for field_name in ['FileName', 'Title', 'Name']:
                elem = msg_root.find(field_name)
                if elem is not None and elem.text:
                    file_name = elem.text
                    logger.info(f"[FILE] 从字段 {field_name} 获取文件名: {file_name}")
                    break
            
            if not file_name:
                file_name = 'document.docx'
                logger.info(f"[FILE] 未找到文件名字段，使用默认: {file_name}")
            
            logger.info(f"[FILE] 收到文件: {file_name}, MediaId: {media_id}")
            
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
            logger.info("[FILE] 已返回处理中提示，异步线程已启动")
            return encrypted_reply
        
        # ========== 处理文本消息 ==========
        elif msg_type == 'text':
            content = msg_root.find('Content').text or ''
            logger.info(f"[TEXT] 收到文本消息: {content}")
            
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
        
        # ========== 处理图片消息（添加日志） ==========
        elif msg_type == 'image':
            logger.info(f"[IMAGE] 收到图片消息，MediaId: {msg_root.find('MediaId').text if msg_root.find('MediaId') is not None else 'N/A'}")
            reply_msg = create_text_response(
                from_user, to_user, 
                "请发送Word或Excel文件，我会帮您转换为PDF 📄\n\n（暂不支持图片转换）"
            )
            encrypted_reply = wecom_api.crypto.encrypt_message(reply_msg, nonce, timestamp)
            return encrypted_reply
        
        # ========== 其他消息类型 ==========
        else:
            logger.info(f"[OTHER] 收到其他类型消息: {msg_type}")
            reply_msg = create_text_response(
                from_user, to_user, 
                "请发送Word或Excel文件，我会帮您转换为PDF 📄"
            )
            encrypted_reply = wecom_api.crypto.encrypt_message(reply_msg, nonce, timestamp)
            return encrypted_reply
            
    except Exception as e:
        logger.error(f"[ERROR] 处理消息异常: {str(e)}", exc_info=True)
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


@app.route('/debug/recent', methods=['GET'])
def debug_recent():
    """调试接口：查看最近处理的消息"""
    return {
        'processed_messages_count': len(processed_messages),
        'recent_messages': list(processed_messages.keys())[-10:],  # 最近10条
        'cache_ttl': MESSAGE_CACHE_TTL,
        'service_status': 'running'
    }


# ========== iOS Shortcuts API ==========

ALLOWED_EXTENSIONS = {'.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'}

@app.route('/api/convert', methods=['POST'])
def api_convert():
    """
    iOS Shortcuts 文档转换接口
    
    接收 multipart/form-data 格式的文件上传，返回 PDF 二进制流。
    
    请求:
        - Method: POST
        - Content-Type: multipart/form-data
        - Field: 'file' (required)
    
    响应:
        - 成功: PDF 文件 (Content-Type: application/pdf)
        - 失败: JSON 错误信息
    """
    from flask import send_file
    import io
    
    logger.info("=== iOS Shortcuts API 请求 ===")
    logger.info(f"Remote IP: {request.remote_addr}")
    
    # 检查文件字段
    if 'file' not in request.files:
        logger.error("请求中没有 'file' 字段")
        return {'error': '请上传文件', 'field': 'file'}, 400
    
    file = request.files['file']
    
    if file.filename == '':
        logger.error("文件名为空")
        return {'error': '文件名为空'}, 400
    
    # 检查文件扩展名
    file_ext = Path(file.filename).suffix.lower()
    if file_ext not in ALLOWED_EXTENSIONS:
        logger.error(f"不支持的文件类型: {file_ext}")
        return {
            'error': f'不支持的文件类型: {file_ext}',
            'allowed': list(ALLOWED_EXTENSIONS)
        }, 400
    
    input_file = None
    output_pdf = None
    
    try:
        # 保存上传的文件
        timestamp_ms = int(time.time() * 1000)
        input_file = os.path.join(config.TEMP_DIR, f"api_input_{timestamp_ms}{file_ext}")
        file.save(input_file)
        logger.info(f"文件已保存: {input_file}")
        
        # 转换为 PDF
        logger.info("开始转换...")
        output_pdf = converter.convert_to_pdf(input_file)
        logger.info(f"转换完成: {output_pdf}")
        
        # 读取 PDF 到内存
        with open(output_pdf, 'rb') as f:
            pdf_data = f.read()
        
        # 生成输出文件名
        output_filename = Path(file.filename).stem + '.pdf'
        
        logger.info(f"返回 PDF: {output_filename}, 大小: {len(pdf_data)} 字节")
        
        # 返回 PDF 流
        return send_file(
            io.BytesIO(pdf_data),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=output_filename
        )
        
    except Exception as e:
        logger.error(f"转换失败: {str(e)}", exc_info=True)
        return {'error': f'转换失败: {str(e)}'}, 500
        
    finally:
        # 清理临时文件
        if input_file:
            converter.cleanup_file(input_file)
        if output_pdf:
            converter.cleanup_file(output_pdf)


if __name__ == '__main__':
    # 确保临时目录存在
    os.makedirs(config.TEMP_DIR, exist_ok=True)
    
    app.run(
        host=config.HOST,
        port=config.PORT,
        debug=config.DEBUG,
        threaded=True
    )
