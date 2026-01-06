"""
Windows文档转换服务 - 使用Microsoft Office或WPS Office COM自动化

部署在Windows虚拟机上，通过HTTP API提供文档转PDF转换服务

支持的Office套件（按优先级）:
1. Microsoft Office (Word, Excel, PowerPoint)
2. WPS Office (KWPS, KET, KWPP)

依赖:
- Windows 10/11 或 Windows Server
- Microsoft Office 或 WPS Office (2019及以上版本推荐)
- Python 3.11+
- pywin32 包

安装:
    pip install flask pywin32

运行:
    python windows_converter_service.py
"""

from flask import Flask, request, send_file, jsonify
import win32com.client
import pythoncom
import os
import logging
import tempfile
import time
from pathlib import Path
from werkzeug.utils import secure_filename

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB

# 临时文件目录
TEMP_DIR = Path(tempfile.gettempdir()) / 'office_converter'
TEMP_DIR.mkdir(exist_ok=True)

# 支持的文件类型
WORD_EXTENSIONS = ['.doc', '.docx']
EXCEL_EXTENSIONS = ['.xls', '.xlsx']
POWERPOINT_EXTENSIONS = ['.ppt', '.pptx']
SUPPORTED_EXTENSIONS = WORD_EXTENSIONS + EXCEL_EXTENSIONS + POWERPOINT_EXTENSIONS

def get_word_application():
    """获取Word或WPS文字应用程序实例"""
    # 优先尝试 MS Word，然后是 WPS 文字
    progids = ["Word.Application", "KWPS.Application"]
    
    for progid in progids:
        try:
            app = win32com.client.Dispatch(progid)
            logger.info(f"成功启动: {progid}")
            return app, progid
        except Exception as e:
            logger.debug(f"无法启动 {progid}: {str(e)}")
            continue
    
    raise Exception("未找到可用的Word/WPS文字处理应用")


def convert_word_to_pdf(input_path: str, output_path: str) -> bool:
    """使用Word或WPS文字COM转换为PDF"""
    word = None
    doc = None
    progid = None
    
    try:
        # 初始化COM
        pythoncom.CoInitialize()
        
        # 获取可用的Word应用
        word, progid = get_word_application()
        word.Visible = False
        word.DisplayAlerts = 0  # 禁用警告对话框
        
        # 打开文档
        logger.info(f"使用 {progid} 打开文档: {input_path}")
        doc = word.Documents.Open(input_path)
        
        # 导出为PDF (wdFormatPDF = 17, WPS也使用相同的值)
        logger.info(f"导出PDF: {output_path}")
        doc.SaveAs(output_path, FileFormat=17)
        
        logger.info(f"{progid} 转换成功")
        return True
        
    except Exception as e:
        logger.error(f"Word/WPS转换失败: {str(e)}")
        return False
        
    finally:
        # 清理资源
        try:
            if doc:
                doc.Close(SaveChanges=False)
            if word:
                word.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

def get_excel_application():
    """获取Excel或WPS表格应用程序实例"""
    # 优先尝试 MS Excel，然后是 WPS 表格
    progids = ["Excel.Application", "KET.Application"]
    
    for progid in progids:
        try:
            app = win32com.client.Dispatch(progid)
            logger.info(f"成功启动: {progid}")
            return app, progid
        except Exception as e:
            logger.debug(f"无法启动 {progid}: {str(e)}")
            continue
    
    raise Exception("未找到可用的Excel/WPS表格应用")


def convert_excel_to_pdf(input_path: str, output_path: str) -> bool:
    """使用Excel或WPS表格COM转换为PDF"""
    excel = None
    workbook = None
    progid = None
    
    try:
        pythoncom.CoInitialize()
        
        excel, progid = get_excel_application()
        excel.Visible = False
        excel.DisplayAlerts = False
        
        logger.info(f"使用 {progid} 打开文档: {input_path}")
        workbook = excel.Workbooks.Open(input_path)
        
        # 导出为PDF (xlTypePDF = 0, WPS也使用相同的值)
        logger.info(f"导出PDF: {output_path}")
        workbook.ExportAsFixedFormat(0, output_path)
        
        logger.info(f"{progid} 转换成功")
        return True
        
    except Exception as e:
        logger.error(f"Excel/WPS表格转换失败: {str(e)}")
        return False
        
    finally:
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

def get_powerpoint_application():
    """获取PowerPoint或WPS演示应用程序实例"""
    # 优先尝试 MS PowerPoint，然后是 WPS 演示
    progids = ["PowerPoint.Application", "KWPP.Application"]
    
    for progid in progids:
        try:
            app = win32com.client.Dispatch(progid)
            logger.info(f"成功启动: {progid}")
            return app, progid
        except Exception as e:
            logger.debug(f"无法启动 {progid}: {str(e)}")
            continue
    
    raise Exception("未找到可用的PowerPoint/WPS演示应用")


def convert_powerpoint_to_pdf(input_path: str, output_path: str) -> bool:
    """使用PowerPoint或WPS演示COM转换为PDF"""
    powerpoint = None
    presentation = None
    progid = None
    
    try:
        pythoncom.CoInitialize()
        
        powerpoint, progid = get_powerpoint_application()
        
        logger.info(f"使用 {progid} 打开文档: {input_path}")
        presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
        
        # 导出为PDF (ppSaveAsPDF = 32, WPS也使用相同的值)
        logger.info(f"导出PDF: {output_path}")
        presentation.SaveAs(output_path, 32)
        
        logger.info(f"{progid} 转换成功")
        return True
        
    except Exception as e:
        logger.error(f"PowerPoint/WPS演示转换失败: {str(e)}")
        return False
        
    finally:
        try:
            if presentation:
                presentation.Close()
            if powerpoint:
                powerpoint.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

def detect_available_apps():
    """检测可用的Office/WPS应用"""
    apps = {}
    
    # 检测文字处理应用
    for progid in ["Word.Application", "KWPS.Application"]:
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch(progid)
            app.Quit()
            apps['word'] = progid
            break
        except:
            pass
        finally:
            pythoncom.CoUninitialize()
    
    # 检测表格应用
    for progid in ["Excel.Application", "KET.Application"]:
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch(progid)
            app.Quit()
            apps['excel'] = progid
            break
        except:
            pass
        finally:
            pythoncom.CoUninitialize()
    
    # 检测演示应用
    for progid in ["PowerPoint.Application", "KWPP.Application"]:
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch(progid)
            app.Quit()
            apps['powerpoint'] = progid
            break
        except:
            pass
        finally:
            pythoncom.CoUninitialize()
    
    return apps


@app.route('/health', methods=['GET'])
def health_check():
    """健康检查"""
    # 检测可用的应用
    available_apps = detect_available_apps()
    
    return jsonify({
        'status': 'ok',
        'service': 'windows-office-converter',
        'version': '1.1.0',
        'available_apps': available_apps,
        'supported_extensions': SUPPORTED_EXTENSIONS
    })

@app.route('/convert', methods=['POST'])
def convert_document():
    """
    转换文档为PDF
    
    请求:
        - 文件作为 multipart/form-data 上传
        - 字段名: 'document'
    
    响应:
        - 成功: PDF文件 (application/pdf)
        - 失败: JSON错误信息
    """
    # 检查文件
    if 'document' not in request.files:
        return jsonify({'error': '没有文件上传'}), 400
    
    file = request.files['document']
    
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400
    
    # 验证文件扩展名
    filename = secure_filename(file.filename)
    file_ext = Path(filename).suffix.lower()
    
    if file_ext not in SUPPORTED_EXTENSIONS:
        return jsonify({
            'error': f'不支持的文件类型: {file_ext}',
            'supported': SUPPORTED_EXTENSIONS
        }), 400
    
    # 生成临时文件路径
    timestamp = int(time.time() * 1000)
    input_path = TEMP_DIR / f"input_{timestamp}{file_ext}"
    output_path = TEMP_DIR / f"output_{timestamp}.pdf"
    
    try:
        # 保存上传的文件
        logger.info(f"接收文件: {filename} ({file_ext})")
        file.save(str(input_path))
        
        # 根据文件类型选择转换方法
        success = False
        
        if file_ext in WORD_EXTENSIONS:
            success = convert_word_to_pdf(str(input_path), str(output_path))
        elif file_ext in EXCEL_EXTENSIONS:
            success = convert_excel_to_pdf(str(input_path), str(output_path))
        elif file_ext in POWERPOINT_EXTENSIONS:
            success = convert_powerpoint_to_pdf(str(input_path), str(output_path))
        
        if not success:
            return jsonify({'error': '转换失败'}), 500
        
        if not output_path.exists():
            return jsonify({'error': 'PDF文件未生成'}), 500
        
        # 返回PDF文件
        return send_file(
            str(output_path),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"{Path(filename).stem}.pdf"
        )
        
    except Exception as e:
        logger.error(f"转换异常: {str(e)}", exc_info=True)
        return jsonify({'error': f'服务器错误: {str(e)}'}), 500
        
    finally:
        # 清理临时文件（延迟删除，确保响应完成）
        def cleanup():
            time.sleep(5)
            try:
                if input_path.exists():
                    os.remove(input_path)
                if output_path.exists():
                    os.remove(output_path)
            except:
                pass
        
        import threading
        threading.Thread(target=cleanup, daemon=True).start()

if __name__ == '__main__':
    logger.info("=" * 60)
    logger.info("Windows Office 转换服务启动")
    logger.info("监听端口: 8080")
    logger.info("临时目录: " + str(TEMP_DIR))
    logger.info("=" * 60)
    
    app.run(
        host='0.0.0.0',
        port=8080,
        debug=False,
        threaded=True
    )
