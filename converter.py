import os
import subprocess
import logging
import requests
from pathlib import Path
from config import config

logger = logging.getLogger(__name__)

class DocumentConverter:
    """Office文档转PDF转换器 - 支持Windows和LibreOffice双引擎"""
    
    def __init__(self):
        self.temp_dir = Path(config.TEMP_DIR)
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.libreoffice_path = config.LIBREOFFICE_PATH
        self.timeout = config.CONVERSION_TIMEOUT
        
        # Windows转换服务配置
        self.windows_enabled = config.WINDOWS_CONVERTER_ENABLED
        self.windows_url = config.WINDOWS_CONVERTER_URL
        self.windows_timeout = config.WINDOWS_CONVERTER_TIMEOUT
        
        logger.info(f"转换引擎配置: Windows={self.windows_enabled}, URL={self.windows_url}")
    
    def convert_to_pdf(self, input_file_path: str) -> str:
        """
        将Office文档转换为PDF（智能选择转换引擎）
        
        Args:
            input_file_path: 输入文件路径
            
        Returns:
            str: 转换后的PDF文件路径
            
        Raises:
            Exception: 转换失败时抛出异常
        """
        input_path = Path(input_file_path)
        
        if not input_path.exists():
            raise FileNotFoundError(f"文件不存在: {input_file_path}")
        
        # 检查文件类型
        supported_extensions = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']
        if input_path.suffix.lower() not in supported_extensions:
            raise ValueError(f"不支持的文件类型: {input_path.suffix}")
        
        # 输出PDF文件名
        output_pdf = input_path.parent / f"{input_path.stem}.pdf"
        
        # 优先使用Windows转换服务
        if self.windows_enabled and self.windows_url:
            logger.info("尝试使用Windows转换服务...")
            try:
                result = self._convert_via_windows(input_path, output_pdf)
                if result:
                    logger.info("✅ Windows转换成功")
                    return result
                else:
                    logger.warning("Windows转换失败，降级到LibreOffice")
            except Exception as e:
                logger.error(f"Windows转换异常: {str(e)}，降级到LibreOffice")
        
        # 降级使用LibreOffice
        logger.info("使用LibreOffice转换...")
        return self._convert_via_libreoffice(input_path, output_pdf)
    
    def _convert_via_windows(self, input_path: Path, output_pdf: Path) -> str:
        """
        通过Windows服务转换文档
        
        Args:
            input_path: 输入文件路径
            output_pdf: 输出PDF路径
            
        Returns:
            str: PDF文件路径，失败返回None
        """
        try:
            # 发送文件到Windows服务
            with open(input_path, 'rb') as f:
                files = {'document': (input_path.name, f, 'application/octet-stream')}
                
                response = requests.post(
                    f"{self.windows_url}/convert",
                    files=files,
                    timeout=self.windows_timeout
                )
            
            # 检查响应状态
            if response.status_code != 200:
                logger.error(f"Windows服务返回错误: {response.status_code}")
                try:
                    error_data = response.json()
                    logger.error(f"错误详情: {error_data}")
                except:
                    pass
                return None
            
            # 保存返回的PDF
            with open(output_pdf, 'wb') as f:
                f.write(response.content)
            
            if not output_pdf.exists():
                logger.error("PDF文件保存失败")
                return None
            
            logger.info(f"Windows转换成功: {output_pdf.name}")
            return str(output_pdf)
            
        except requests.exceptions.Timeout:
            logger.error(f"Windows服务超时(>{self.windows_timeout}秒)")
            return None
        except requests.exceptions.ConnectionError:
            logger.error("无法连接到Windows服务")
            return None
        except Exception as e:
            logger.error(f"Windows转换异常: {str(e)}")
            return None
    
    def _convert_via_libreoffice(self, input_path: Path, output_pdf: Path) -> str:
        """
        通过LibreOffice转换文档（备用方案）
        
        Args:
            input_path: 输入文件路径
            output_pdf: 输出PDF路径
            
        Returns:
            str: PDF文件路径
            
        Raises:
            Exception: 转换失败时抛出异常
        """
        # 删除旧PDF
        if output_pdf.exists():
            os.remove(output_pdf)
        
        try:
            # 调用LibreOffice进行转换
            cmd = [
                self.libreoffice_path,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(input_path.parent),
                str(input_path)
            ]
            
            logger.info(f"开始LibreOffice转换: {input_path.name}")
            
            result = subprocess.run(
                cmd,
                timeout=self.timeout,
                capture_output=True,
                text=True
            )
            
            if result.returncode != 0:
                logger.error(f"LibreOffice转换失败: {result.stderr}")
                raise Exception(f"LibreOffice转换失败: {result.stderr}")
            
            if not output_pdf.exists():
                raise Exception("PDF文件未生成")
            
            logger.info(f"LibreOffice转换成功: {output_pdf.name}")
            return str(output_pdf)
            
        except subprocess.TimeoutExpired:
            logger.error(f"LibreOffice转换超时: {input_path}")
            raise Exception(f"转换超时(>{self.timeout}秒)")
        except Exception as e:
            logger.error(f"LibreOffice转换异常: {str(e)}")
            raise
    
    def cleanup_file(self, file_path: str):
        """清理临时文件"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.info(f"已清理文件: {file_path}")
        except Exception as e:
            logger.warning(f"清理文件失败: {str(e)}")
