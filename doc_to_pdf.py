#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Doc/Docx文件转换为PDF工具
支持单个文件转换和批量转换
支持Microsoft Word和LibreOffice两种转换引擎
"""

import os
import sys
import logging
import traceback
import subprocess
import time
from pathlib import Path

# 尝试导入win32com模块，用于Word转换
try:
    from win32com import client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

# 配置日志
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'conversion.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 全局变量，用于存储用户指定的LibreOffice路径
USER_SPECIFIED_LO_PATH = None

def set_libreoffice_path(path):
    """
    设置用户指定的LibreOffice路径
    
    Args:
        path: LibreOffice可执行文件路径
    """
    global USER_SPECIFIED_LO_PATH
    USER_SPECIFIED_LO_PATH = path
    logger.info(f"用户指定LibreOffice路径: {path}")

def find_libreoffice_path():
    """
    尝试找到LibreOffice的可执行文件路径
    
    Returns:
        str: LibreOffice可执行文件路径，如果未找到返回None
    """
    # 优先使用用户指定的路径
    if USER_SPECIFIED_LO_PATH:
        if os.path.exists(USER_SPECIFIED_LO_PATH):
            logger.info(f"使用用户指定的LibreOffice路径: {USER_SPECIFIED_LO_PATH}")
            return USER_SPECIFIED_LO_PATH
        else:
            logger.error(f"用户指定的LibreOffice路径不存在: {USER_SPECIFIED_LO_PATH}")
    
    # Windows上可能的LibreOffice路径
    paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\LibreOffice\program\soffice.exe"
    ]
    
    # 检查PATH环境变量
    for path in os.environ["PATH"].split(os.pathsep):
        possible_path = os.path.join(path, "soffice.exe")
        if os.path.exists(possible_path):
            paths.append(possible_path)
    
    # 返回第一个存在的路径
    for path in paths:
        if os.path.exists(path):
            logger.debug(f"找到LibreOffice: {path}")
            return path
    
    logger.warning("未找到LibreOffice安装路径")
    return None

def convert_with_word(input_file, output_file):
    """
    使用Microsoft Word将文件转换为PDF
    
    Args:
        input_file: 输入文件路径
        output_file: 输出PDF文件路径
    
    Returns:
        bool: 转换是否成功
    """
    if not WIN32COM_AVAILABLE:
        logger.error("win32com模块不可用，无法使用Word进行转换")
        return False
        
    word = None
    doc = None
    
    try:
        # 创建Word应用对象
        word = client.Dispatch('Word.Application')
        word.Visible = False  # 不显示Word界面
        word.DisplayAlerts = 0  # 不显示警告
        
        # 打开文档 - 使用绝对路径，避免路径问题
        doc = word.Documents.Open(input_file)
        
        # 保存为PDF
        doc.SaveAs(output_file, FileFormat=17)  # 17 表示PDF格式
        
        return True
        
    except Exception as e:
        logger.error(f"使用Word转换文件时出错: {input_file}")
        logger.error(f"错误详情: {str(e)}")
        logger.debug(traceback.format_exc())  # 记录详细的堆栈跟踪
        return False
    finally:
        # 确保资源被释放
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)  # 不保存更改
                logger.debug(f"Word文档已关闭: {input_file}")
        except Exception as e:
            logger.error(f"关闭Word文档时出错: {str(e)}")
        
        try:
            if word is not None:
                word.Quit()
                logger.debug("Word应用已退出")
        except Exception as e:
            logger.error(f"退出Word应用时出错: {str(e)}")
        
        # 强制垃圾回收，确保Word进程被释放
        import gc
        gc.collect()

def convert_with_libreoffice(input_file, output_file):
    """
    使用LibreOffice将文件转换为PDF
    
    Args:
        input_file: 输入文件路径
        output_file: 输出PDF文件路径
    
    Returns:
        bool: 转换是否成功
    """
    # 找到LibreOffice可执行文件
    libreoffice_path = find_libreoffice_path()
    if libreoffice_path is None:
        logger.error("无法找到LibreOffice，无法进行转换")
        return False
    
    try:
        # 构建命令
        output_dir = os.path.dirname(output_file)
        cmd = [
            libreoffice_path,
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            input_file,
            "--outdir",
            output_dir
        ]
        
        logger.debug(f"执行命令: {' '.join(cmd)}")
        
        # 执行命令
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=300  # 设置5分钟超时
        )
        
        # 检查是否成功
        if result.returncode == 0:
            logger.debug(f"LibreOffice输出: {result.stdout}")
            return True
        else:
            logger.error(f"LibreOffice转换失败，返回代码: {result.returncode}")
            logger.error(f"错误输出: {result.stderr}")
            return False
    
    except subprocess.TimeoutExpired:
        logger.error(f"LibreOffice转换超时: {input_file}")
        return False
    except Exception as e:
        logger.error(f"使用LibreOffice转换文件时出错: {input_file}")
        logger.error(f"错误详情: {str(e)}")
        logger.debug(traceback.format_exc())  # 记录详细的堆栈跟踪
        return False

def convert_single_file(input_file, output_file=None, engine=None):
    """
    将单个doc或docx文件转换为PDF
    
    Args:
        input_file: 输入文件路径
        output_file: 输出PDF文件路径，默认为输入文件同目录同名称的PDF文件
        engine: 使用的转换引擎，可选值: 'word', 'libreoffice', 默认为None(自动选择)
    
    Returns:
        bool: 转换是否成功
    """
    # 确保输入路径是绝对路径
    input_file = os.path.abspath(input_file)
    
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        logger.error(f"输入文件不存在: {input_file}")
        return False
    
    # 获取输入文件的扩展名
    file_ext = os.path.splitext(input_file)[1].lower()
    
    # 检查文件类型是否支持
    if file_ext not in ['.doc', '.docx']:
        logger.error(f"不支持的文件类型: {file_ext}，文件: {input_file}")
        return False
    
    # 如果未指定输出文件，则使用默认路径
    if output_file is None:
        output_file = os.path.splitext(input_file)[0] + '.pdf'
    else:
        output_file = os.path.abspath(output_file)
    
    # 确保输出目录存在
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            logger.info(f"创建输出目录: {output_dir}")
        except Exception as e:
            logger.error(f"创建输出目录失败: {str(e)}")
            return False
    
    start_time = time.time()
    logger.info(f"开始转换文件: {input_file} -> {output_file}")
    
    # 确定使用的引擎
    if engine == 'word':
        success = convert_with_word(input_file, output_file)
    elif engine == 'libreoffice':
        success = convert_with_libreoffice(input_file, output_file)
    else:  # 自动选择
        # 先尝试使用LibreOffice
        if convert_with_libreoffice(input_file, output_file):
            success = True
            logger.info("使用LibreOffice成功转换文件")
        else:
            # 如果LibreOffice失败，尝试使用Word
            logger.info("尝试使用Word进行转换")
            success = convert_with_word(input_file, output_file)
    
    end_time = time.time()
    
    if success:
        logger.info(f"文件转换成功: {output_file} (耗时: {end_time - start_time:.2f}秒)")
    else:
        logger.error(f"文件转换失败: {input_file}")
    
    return success

def convert_batch(directory, engine=None, output_to_root=False):
    """
    批量转换指定目录下的所有doc和docx文件为PDF
    
    Args:
        directory: 要处理的目录路径
        engine: 使用的转换引擎，可选值: 'word', 'libreoffice', 默认为None(自动选择)
        output_to_root: 是否将所有PDF保存到根目录，默认为False（保持原文件夹结构）
    
    Returns:
        dict: 包含转换统计信息的字典
    """
    # 确保目录路径是绝对路径
    directory = os.path.abspath(directory)
    
    if not os.path.exists(directory):
        logger.error(f"目录不存在: {directory}")
        return {"total": 0, "success": 0, "failed": 0, "failed_files": []}
    
    if not os.path.isdir(directory):
        logger.error(f"指定路径不是目录: {directory}")
        return {"total": 0, "success": 0, "failed": 0, "failed_files": []}
    
    start_time = time.time()
    logger.info(f"开始扫描目录: {directory}")
    
    # 获取目录下所有的doc和docx文件
    word_files = []
    try:
        for root, _, files in os.walk(directory):
            for file in files:
                file_ext = os.path.splitext(file)[1].lower()
                if file_ext in ['.doc', '.docx']:
                    word_files.append(os.path.join(root, file))
    except Exception as e:
        logger.error(f"扫描目录时出错: {str(e)}")
        return {"total": 0, "success": 0, "failed": 0, "failed_files": []}
    
    total = len(word_files)
    success = 0
    failed = 0
    failed_files = []
    
    logger.info(f"找到 {total} 个Word文件待转换")
    
    if total == 0:
        logger.warning(f"在目录 {directory} 中未找到任何.doc或.docx文件")
        return {"total": 0, "success": 0, "failed": 0, "failed_files": []}
    
    # 逐个转换文件
    for i, word_file in enumerate(word_files, 1):
        # 计算进度百分比
        progress = (i / total) * 100
        logger.info(f"处理文件 {i}/{total} ({progress:.1f}%): {word_file}")
        
        # 确定输出文件路径
        if output_to_root:
            # 将所有PDF保存到根目录
            file_name = os.path.basename(word_file)
            file_name_without_ext = os.path.splitext(file_name)[0]
            output_file = os.path.join(directory, f"{file_name_without_ext}.pdf")
        else:
            output_file = None  # 使用默认输出路径（与原文件同目录）
        
        if convert_single_file(word_file, output_file=output_file, engine=engine):
            success += 1
        else:
            failed += 1
            failed_files.append(word_file)
        
        # 每处理完5个文件，增加一个短暂延迟，避免Word进程占用过高
        if i % 5 == 0:
            logger.debug("短暂暂停以释放系统资源...")
            time.sleep(1)
    
    end_time = time.time()
    duration = end_time - start_time
    
    # 生成详细的转换报告
    logger.info("="*50)
    logger.info("批量转换完成")
    logger.info(f"总耗时: {duration:.2f}秒")
    logger.info(f"总文件数: {total}")
    logger.info(f"成功转换: {success}")
    logger.info(f"转换失败: {failed}")
    
    if failed > 0:
        logger.info("失败的文件列表:")
        for failed_file in failed_files:
            logger.info(f"  - {failed_file}")
    
    logger.info("="*50)
    
    return {
        "total": total,
        "success": success,
        "failed": failed,
        "failed_files": failed_files,
        "duration": duration
    }

if __name__ == "__main__":
    """
    主函数，支持命令行参数
    使用方法:
    1. 转换单个文件: python doc_to_pdf.py <文件路径>
    2. 批量转换: python doc_to_pdf.py <目录路径>
       或: python doc_to_pdf.py --batch <目录路径>
    3. 指定转换引擎: python doc_to_pdf.py <路径> --engine=libreoffice 或 --engine=word
    4. 指定LibreOffice路径: python doc_to_pdf.py <路径> --libreoffice-path="C:\Program Files\LibreOffice\program\soffice.exe"
    5. 将所有PDF保存到根目录: python doc_to_pdf.py <目录路径> --output-to-root
    """
    # 解析命令行参数
    input_path = None
    batch_mode = False
    engine = None
    output_to_root = False
    
    # 处理参数
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] == "--batch":
            batch_mode = True
            i += 1
        elif sys.argv[i].startswith("--engine="):
            engine_value = sys.argv[i].split("=")[1].lower()
            if engine_value in ['word', 'libreoffice']:
                engine = engine_value
            else:
                print(f"错误: 不支持的引擎类型 '{engine_value}'，请使用 'word' 或 'libreoffice'")
                sys.exit(1)
            i += 1
        elif sys.argv[i].startswith("--libreoffice-path="):
            lo_path = sys.argv[i].split("=", 1)[1]
            set_libreoffice_path(lo_path)
            i += 1
        elif sys.argv[i] == "--output-to-root":
            output_to_root = True
            i += 1
        else:
            if input_path is None:
                input_path = sys.argv[i]
            else:
                print("错误: 只能指定一个输入路径")
                sys.exit(1)
            i += 1
    
    # 检查是否提供了输入路径
    if input_path is None:
        print("使用方法:")
        print("1. 转换单个文件: python doc_to_pdf.py <文件路径>")
        print("2. 批量转换: python doc_to_pdf.py <目录路径>")
        print("   或: python doc_to_pdf.py --batch <目录路径>")
        print("3. 指定转换引擎: python doc_to_pdf.py <路径> --engine=libreoffice 或 --engine=word")
        print("4. 指定LibreOffice路径: python doc_to_pdf.py <路径> --libreoffice-path=\"C:\\Program Files\\LibreOffice\\program\\soffice.exe\"")
        print("5. 将所有PDF保存到根目录: python doc_to_pdf.py <目录路径> --output-to-root")
        print("\n如果LibreOffice安装在非标准位置，请使用--libreoffice-path参数指定路径")
        print("如果需要将嵌套文件夹中的所有文件转换后保存到根目录，请使用--output-to-root参数")
        sys.exit(1)
    
    # 根据路径类型和模式执行转换
    if os.path.isfile(input_path):
        # 单个文件转换
        if batch_mode:
            print("警告: --batch 参数与文件路径不兼容，将忽略 --batch 参数")
        if output_to_root:
            print("警告: --output-to-root 参数仅对目录批量转换有效，将忽略该参数")
        convert_single_file(input_path, engine=engine)
    elif os.path.isdir(input_path):
        # 批量转换
        convert_batch(input_path, engine=engine, output_to_root=output_to_root)
    else:
        print(f"错误: {input_path} 不存在")
        sys.exit(1)