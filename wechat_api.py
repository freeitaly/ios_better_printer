import requests
import time
import logging
import threading
from config import config

logger = logging.getLogger(__name__)

class WeChatAPI:
    """微信公众号API封装"""
    
    def __init__(self):
        self.app_id = config.WECHAT_APP_ID
        self.app_secret = config.WECHAT_APP_SECRET
        self.access_token = None
        self.token_expires_at = 0
        self._token_lock = threading.Lock()
    
    def get_access_token(self) -> str:
        """
        获取access_token，带缓存机制和线程安全
        
        Returns:
            str: access_token
        """
        # 先检查缓存（无锁快速路径）
        if self.access_token and time.time() < self.token_expires_at:
            return self.access_token
        
        # 需要刷新token时加锁
        with self._token_lock:
            # 双重检查，避免重复请求
            if self.access_token and time.time() < self.token_expires_at:
                return self.access_token
            
            # 请求新的access_token
            url = "https://api.weixin.qq.com/cgi-bin/token"
            params = {
                'grant_type': 'client_credential',
                'appid': self.app_id,
                'secret': self.app_secret
            }
            
            try:
                response = requests.get(url, params=params, timeout=10)
                response.raise_for_status()
                data = response.json()
                
                if 'access_token' not in data:
                    error_msg = data.get('errmsg', '未知错误')
                    raise Exception(f"获取access_token失败: {error_msg}")
                
                self.access_token = data['access_token']
                # 提前5分钟过期，避免边界问题
                self.token_expires_at = time.time() + data['expires_in'] - 300
                
                logger.info("成功获取access_token")
                return self.access_token
                
            except Exception as e:
                logger.error(f"获取access_token异常: {str(e)}")
                raise
    
    def download_media(self, media_id: str, save_path: str) -> str:
        """
        下载微信临时素材
        
        Args:
            media_id: 媒体文件ID
            save_path: 保存路径
            
        Returns:
            str: 保存的文件路径
        """
        access_token = self.get_access_token()
        url = "https://api.weixin.qq.com/cgi-bin/media/get"
        params = {
            'access_token': access_token,
            'media_id': media_id
        }
        
        try:
            response = requests.get(url, params=params, timeout=30, stream=True)
            response.raise_for_status()
            
            # 检查是否是错误响应
            content_type = response.headers.get('Content-Type', '')
            if 'application/json' in content_type:
                error_data = response.json()
                raise Exception(f"下载失败: {error_data.get('errmsg', '未知错误')}")
            
            # 保存文件
            with open(save_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            logger.info(f"文件下载成功: {save_path}")
            return save_path
            
        except Exception as e:
            logger.error(f"下载媒体文件异常: {str(e)}")
            raise
    
    def upload_media(self, file_path: str, media_type: str = 'file') -> str:
        """
        上传临时素材
        
        Args:
            file_path: 文件路径
            media_type: 媒体类型 (image/voice/video/file)
            
        Returns:
            str: media_id
        """
        access_token = self.get_access_token()
        url = "https://api.weixin.qq.com/cgi-bin/media/upload"
        params = {
            'access_token': access_token,
            'type': media_type
        }
        
        try:
            with open(file_path, 'rb') as f:
                files = {'media': f}
                response = requests.post(url, params=params, files=files, timeout=60)
                response.raise_for_status()
                data = response.json()
            
            if 'media_id' not in data:
                error_msg = data.get('errmsg', '未知错误')
                raise Exception(f"上传失败: {error_msg}")
            
            media_id = data['media_id']
            logger.info(f"文件上传成功: {media_id}")
            return media_id
            
        except Exception as e:
            logger.error(f"上传媒体文件异常: {str(e)}")
            raise
    
    def send_file_message(self, to_user: str, media_id: str) -> bool:
        """
        发送文件类型的客服消息
        
        注意：被动回复不支持file类型，必须使用客服消息接口
        
        Args:
            to_user: 接收用户的OpenID
            media_id: 文件的media_id
            
        Returns:
            bool: 是否发送成功
        """
        access_token = self.get_access_token()
        url = f"https://api.weixin.qq.com/cgi-bin/message/custom/send?access_token={access_token}"
        
        data = {
            "touser": to_user,
            "msgtype": "file",
            "file": {
                "media_id": media_id
            }
        }
        
        try:
            response = requests.post(url, json=data, timeout=10)
            response.raise_for_status()
            result = response.json()
            
            if result.get('errcode', 0) != 0:
                error_msg = result.get('errmsg', '未知错误')
                logger.error(f"发送客服消息失败: {error_msg}")
                return False
            
            logger.info(f"客服消息发送成功: to={to_user}")
            return True
            
        except Exception as e:
            logger.error(f"发送客服消息异常: {str(e)}")
            return False
    
    def send_text_message(self, to_user: str, content: str) -> bool:
        """
        发送文本类型的客服消息
        
        Args:
            to_user: 接收用户的OpenID
            content: 文本内容
            
        Returns:
            bool: 是否发送成功
        """
        access_token = self.get_access_token()
        url = f"https://api.weixin.qq.com/cgi-bin/message/custom/send?access_token={access_token}"
        
        data = {
            "touser": to_user,
            "msgtype": "text",
            "text": {
                "content": content
            }
        }
        
        try:
            response = requests.post(url, json=data, timeout=10)
            response.raise_for_status()
            result = response.json()
            
            if result.get('errcode', 0) != 0:
                error_msg = result.get('errmsg', '未知错误')
                logger.error(f"发送客服消息失败: {error_msg}")
                return False
            
            logger.info(f"客服文本消息发送成功: to={to_user}")
            return True
            
        except Exception as e:
            logger.error(f"发送客服消息异常: {str(e)}")
            return False
