import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # 微信公众号配置（已弃用，保留兼容性）
    WECHAT_APP_ID = os.getenv('WECHAT_APP_ID', '')
    WECHAT_APP_SECRET = os.getenv('WECHAT_APP_SECRET', '')
    WECHAT_TOKEN = os.getenv('WECHAT_TOKEN', 'your_token_here')
    
    # 企业微信配置
    WECOM_CORP_ID = os.getenv('WECOM_CORP_ID', '')
    WECOM_AGENT_ID = os.getenv('WECOM_AGENT_ID', '')
    WECOM_SECRET = os.getenv('WECOM_SECRET', '')
    WECOM_TOKEN = os.getenv('WECOM_TOKEN', '')
    WECOM_ENCODING_AES_KEY = os.getenv('WECOM_ENCODING_AES_KEY', '')
    
    # 文件存储配置
    TEMP_DIR = os.getenv('TEMP_DIR', '/app/temp_files')
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
    
    # LibreOffice配置（备用转换引擎）
    LIBREOFFICE_PATH = os.getenv('LIBREOFFICE_PATH', '/usr/bin/soffice')
    CONVERSION_TIMEOUT = int(os.getenv('CONVERSION_TIMEOUT', '30'))  # 秒
    
    # Windows转换服务配置（主转换引擎）
    WINDOWS_CONVERTER_URL = os.getenv('WINDOWS_CONVERTER_URL', '')
    WINDOWS_CONVERTER_ENABLED = os.getenv('WINDOWS_CONVERTER_ENABLED', 'false').lower() == 'true'
    WINDOWS_CONVERTER_TIMEOUT = int(os.getenv('WINDOWS_CONVERTER_TIMEOUT', '60'))  # 秒
    
    # Flask配置
    DEBUG = os.getenv('DEBUG', 'False').lower() == 'true'
    HOST = os.getenv('HOST', '0.0.0.0')
    PORT = int(os.getenv('PORT', '5000'))

config = Config()
