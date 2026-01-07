"""
企业微信 API 封装

用于与企业微信自建应用交互，包括：
- 消息回调验证（AES加解密）
- 获取access_token
- 下载/上传媒体文件
- 发送应用消息
"""

import base64
import hashlib
import time
import struct
import socket
import logging
import threading
import requests
from Crypto.Cipher import AES
from config import config

logger = logging.getLogger(__name__)


class WXBizMsgCrypt:
    """企业微信消息加解密工具"""
    
    def __init__(self, token: str, encoding_aes_key: str, corp_id: str):
        self.token = token
        self.corp_id = corp_id
        # EncodingAESKey 是 Base64 编码的 AES 密钥
        self.aes_key = base64.b64decode(encoding_aes_key + '=')
    
    def _get_sha1_signature(self, token: str, timestamp: str, nonce: str, encrypt: str) -> str:
        """计算签名"""
        sort_list = [token, timestamp, nonce, encrypt]
        sort_list.sort()
        sha = hashlib.sha1(''.join(sort_list).encode('utf-8'))
        return sha.hexdigest()
    
    def _pkcs7_encode(self, data: bytes) -> bytes:
        """PKCS7 填充"""
        block_size = 32
        pad_count = block_size - (len(data) % block_size)
        return data + bytes([pad_count] * pad_count)
    
    def _pkcs7_decode(self, data: bytes) -> bytes:
        """PKCS7 去除填充"""
        pad_count = data[-1]
        return data[:-pad_count]
    
    def decrypt(self, encrypt_msg: str) -> str:
        """解密消息"""
        try:
            cipher = AES.new(self.aes_key, AES.MODE_CBC, self.aes_key[:16])
            decrypted = cipher.decrypt(base64.b64decode(encrypt_msg))
            decrypted = self._pkcs7_decode(decrypted)
            
            # 解析消息: random(16) + msg_len(4) + msg + corp_id
            msg_len = struct.unpack('>I', decrypted[16:20])[0]
            msg = decrypted[20:20 + msg_len].decode('utf-8')
            from_corp_id = decrypted[20 + msg_len:].decode('utf-8')
            
            if from_corp_id != self.corp_id:
                raise ValueError(f"CorpID不匹配: {from_corp_id} != {self.corp_id}")
            
            return msg
        except Exception as e:
            logger.error(f"解密失败: {str(e)}")
            raise
    
    def encrypt(self, reply_msg: str) -> str:
        """加密消息"""
        # random(16) + msg_len(4) + msg + corp_id
        random_bytes = get_random_bytes(16)
        msg_bytes = reply_msg.encode('utf-8')
        msg_len = struct.pack('>I', len(msg_bytes))
        corp_id_bytes = self.corp_id.encode('utf-8')
        
        plain = random_bytes + msg_len + msg_bytes + corp_id_bytes
        plain = self._pkcs7_encode(plain)
        
        cipher = AES.new(self.aes_key, AES.MODE_CBC, self.aes_key[:16])
        encrypted = cipher.encrypt(plain)
        return base64.b64encode(encrypted).decode('utf-8')
    
    def verify_url(self, msg_signature: str, timestamp: str, nonce: str, echostr: str) -> str:
        """验证回调URL，返回解密后的echostr"""
        signature = self._get_sha1_signature(self.token, timestamp, nonce, echostr)
        if signature != msg_signature:
            raise ValueError("签名验证失败")
        return self.decrypt(echostr)
    
    def decrypt_message(self, msg_signature: str, timestamp: str, nonce: str, encrypt_msg: str) -> str:
        """验证并解密消息"""
        signature = self._get_sha1_signature(self.token, timestamp, nonce, encrypt_msg)
        if signature != msg_signature:
            raise ValueError("消息签名验证失败")
        return self.decrypt(encrypt_msg)
    
    def encrypt_message(self, reply_msg: str, nonce: str, timestamp: str = None) -> str:
        """加密回复消息，返回XML格式"""
        timestamp = timestamp or str(int(time.time()))
        encrypt = self.encrypt(reply_msg)
        signature = self._get_sha1_signature(self.token, timestamp, nonce, encrypt)
        
        return f"""<xml>
<Encrypt><![CDATA[{encrypt}]]></Encrypt>
<MsgSignature><![CDATA[{signature}]]></MsgSignature>
<TimeStamp>{timestamp}</TimeStamp>
<Nonce><![CDATA[{nonce}]]></Nonce>
</xml>"""


def get_random_bytes(n: int) -> bytes:
    """生成随机字节"""
    import os
    return os.urandom(n)


class WeComAPI:
    """企业微信API封装"""
    
    def __init__(self):
        self.corp_id = config.WECOM_CORP_ID
        self.agent_id = config.WECOM_AGENT_ID
        self.secret = config.WECOM_SECRET
        self.access_token = None
        self.token_expires_at = 0
        self._token_lock = threading.Lock()
        
        # 消息加解密工具
        self.crypto = WXBizMsgCrypt(
            token=config.WECOM_TOKEN,
            encoding_aes_key=config.WECOM_ENCODING_AES_KEY,
            corp_id=self.corp_id
        )
    
    def get_access_token(self) -> str:
        """获取access_token，带缓存机制和线程安全"""
        if self.access_token and time.time() < self.token_expires_at:
            return self.access_token
        
        with self._token_lock:
            if self.access_token and time.time() < self.token_expires_at:
                return self.access_token
            
            url = "https://qyapi.weixin.qq.com/cgi-bin/gettoken"
            params = {
                'corpid': self.corp_id,
                'corpsecret': self.secret
            }
            
            try:
                response = requests.get(url, params=params, timeout=10)
                response.raise_for_status()
                data = response.json()
                
                if data.get('errcode', 0) != 0:
                    raise Exception(f"获取access_token失败: {data.get('errmsg')}")
                
                self.access_token = data['access_token']
                self.token_expires_at = time.time() + data['expires_in'] - 300
                
                logger.info("成功获取企业微信access_token")
                return self.access_token
                
            except Exception as e:
                logger.error(f"获取access_token异常: {str(e)}")
                raise
    
    def download_media(self, media_id: str, save_path: str) -> str:
        """下载媒体文件"""
        access_token = self.get_access_token()
        url = "https://qyapi.weixin.qq.com/cgi-bin/media/get"
        params = {
            'access_token': access_token,
            'media_id': media_id
        }
        
        try:
            response = requests.get(url, params=params, timeout=60, stream=True)
            
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
        """上传临时素材"""
        access_token = self.get_access_token()
        url = "https://qyapi.weixin.qq.com/cgi-bin/media/upload"
        params = {
            'access_token': access_token,
            'type': media_type
        }
        
        try:
            with open(file_path, 'rb') as f:
                files = {'media': f}
                response = requests.post(url, params=params, files=files, timeout=120)
                response.raise_for_status()
                data = response.json()
            
            if data.get('errcode', 0) != 0:
                raise Exception(f"上传失败: {data.get('errmsg', '未知错误')}")
            
            media_id = data['media_id']
            logger.info(f"文件上传成功: {media_id}")
            return media_id
            
        except Exception as e:
            logger.error(f"上传媒体文件异常: {str(e)}")
            raise
    
    def send_file_message(self, to_user: str, media_id: str) -> bool:
        """发送文件消息"""
        access_token = self.get_access_token()
        url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}"
        
        data = {
            "touser": to_user,
            "msgtype": "file",
            "agentid": int(self.agent_id),
            "file": {
                "media_id": media_id
            }
        }
        
        try:
            response = requests.post(url, json=data, timeout=10)
            response.raise_for_status()
            result = response.json()
            
            if result.get('errcode', 0) != 0:
                logger.error(f"发送消息失败: {result.get('errmsg')}")
                return False
            
            logger.info(f"文件消息发送成功: to={to_user}")
            return True
            
        except Exception as e:
            logger.error(f"发送消息异常: {str(e)}")
            return False
    
    def send_text_message(self, to_user: str, content: str) -> bool:
        """发送文本消息"""
        access_token = self.get_access_token()
        url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}"
        
        data = {
            "touser": to_user,
            "msgtype": "text",
            "agentid": int(self.agent_id),
            "text": {
                "content": content
            }
        }
        
        try:
            response = requests.post(url, json=data, timeout=10)
            response.raise_for_status()
            result = response.json()
            
            if result.get('errcode', 0) != 0:
                logger.error(f"发送文本消息失败: {result.get('errmsg')}")
                return False
            
            logger.info(f"文本消息发送成功: to={to_user}")
            return True
            
        except Exception as e:
            logger.error(f"发送文本消息异常: {str(e)}")
            return False
