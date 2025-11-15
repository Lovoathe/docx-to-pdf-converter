#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试脚本 - 用于验证doc_to_pdf.py的基本功能
"""

import os
import sys
import time

# 添加当前目录到系统路径以便导入模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from doc_to_pdf import convert_single_file, convert_batch, set_libreoffice_path


def print_separator():
    """打印分隔线"""
    print("=" * 50)


def test_single_file(engine=None):
    """
    测试单个文件转换功能
    """
    print_separator()
    print("测试单个文件转换功能...")
    print_separator()
    
    # 这里可以让用户输入文件路径，或者使用示例文件
    input_file = input("请输入要转换的Word文件路径 (直接按回车跳过此测试): ").strip()
    
    if not input_file:
        print("跳过单个文件转换测试")
        return False
    
    if not os.path.exists(input_file):
        print(f"错误: 文件 '{input_file}' 不存在")
        return False
    
    print(f"准备转换文件: {input_file}")
    if engine:
        print(f"使用引擎: {engine}")
    success = convert_single_file(input_file, engine=engine)
    
    if success:
        output_file = os.path.splitext(input_file)[0] + '.pdf'
        print(f"✅ 转换成功: {output_file}")
    else:
        print(f"❌ 转换失败，请查看日志文件获取详细信息")
    
    return success


def test_batch_conversion(engine=None):
    """
    测试批量转换功能
    """
    print_separator()
    print("测试批量转换功能...")
    print_separator()
    
    # 让用户输入目录路径
    directory = input("请输入要批量转换的目录路径 (直接按回车跳过此测试): ").strip()
    
    if not directory:
        print("跳过批量转换测试")
        return False
    
    if not os.path.exists(directory):
        print(f"错误: 目录 '{directory}' 不存在")
        return False
    
    if not os.path.isdir(directory):
        print(f"错误: '{directory}' 不是一个有效的目录")
        return False
    
    # 询问是否将所有PDF保存到根目录
    print("\nPDF输出位置设置:")
    print("1. 保持与源文件相同的目录结构")
    print("2. 所有PDF都保存到根目录")
    
    output_choice = input("请输入选择 (1-2, 默认1): ").strip()
    output_to_root = output_choice == "2"
    
    print(f"准备批量转换目录: {directory}")
    if engine:
        print(f"使用引擎: {engine}")
    if output_to_root:
        print("所有PDF将保存到根目录")
    print("开始转换，请耐心等待...")
    
    start_time = time.time()
    result = convert_batch(directory, engine=engine, output_to_root=output_to_root)
    end_time = time.time()
    
    print_separator()
    print("批量转换测试结果:")
    print(f"总文件数: {result['total']}")
    print(f"成功转换: {result['success']}")
    print(f"转换失败: {result['failed']}")
    print(f"实际耗时: {end_time - start_time:.2f}秒")
    
    if result['failed'] > 0:
        print("\n失败的文件列表:")
        for failed_file in result['failed_files']:
            print(f"  - {failed_file}")
    
    return result['success'] > 0 or result['total'] == 0


def check_dependencies():
    """
    检查程序依赖
    现在支持Microsoft Word和LibreOffice两种转换引擎，不需要强制安装pywin32
    """
    print_separator()
    print("检查程序依赖...")
    print_separator()
    
    # 检查pywin32（用于Word转换）
    try:
        import win32com.client
        print("✅ pywin32 库已安装（用于Microsoft Word转换）")
    except ImportError:
        print("⚠️  未找到 pywin32 库（仅影响Microsoft Word转换功能）")
        print("如果您计划使用Word进行转换，请运行:")
        print("  pip install pywin32")
    
    # 检查subprocess（Python标准库，用于LibreOffice转换）
    try:
        import subprocess
        print("✅ subprocess 模块可用（用于LibreOffice转换）")
    except ImportError:
        print("❌ 未找到 subprocess 模块（这应该是Python标准库的一部分）")
    
    # 由于现在支持多种转换引擎，不再强制要求特定依赖
    print("\n注意：程序现在支持Microsoft Word和LibreOffice两种转换引擎")
    print("您可以根据需要选择使用哪种引擎")
    
    return True  # 即使缺少pywin32也继续运行，因为可以使用LibreOffice

def select_engine():
    """
    让用户选择转换引擎
    """
    print_separator()
    print("请选择转换引擎:")
    print("1. 自动选择（先尝试LibreOffice，失败则尝试Word）")
    print("2. Microsoft Word")
    print("3. LibreOffice")
    
    choice = input("请输入选择 (1-3, 默认1): ").strip()
    
    if choice == "2":
        return "word"
    elif choice == "3":
        return "libreoffice"
    else:
        return None  # 自动选择


def main():
    """
    主测试函数
    """
    print("欢迎使用 Doc/Docx 转 PDF 工具测试程序")
    print("此程序将帮助您验证转换功能是否正常工作")
    print("\n注意: 请确保您的计算机已安装至少一种转换引擎:")
    print("  - Microsoft Word (商业软件)")
    print("  - LibreOffice (免费开源软件)")
    
    # 检查依赖
    check_dependencies()
    
    # 选择转换引擎
    engine = select_engine()
    
    # 如果是LibreOffice，询问是否需要指定路径
    if engine == "libreoffice" or engine is None:
        print("\nLibreOffice路径设置:")
        print("1. 自动检测（默认）")
        print("2. 手动指定LibreOffice路径")
        
        choice = input("请输入选择 (1-2, 默认1): ").strip()
        
        if choice == "2":
            lo_path = input("请输入LibreOffice可执行文件路径 (soffice.exe): ").strip()
            if lo_path:
                set_libreoffice_path(lo_path)
                print(f"已设置LibreOffice路径: {lo_path}")
            else:
                print("未提供路径，将使用自动检测")
    
    print("\n开始功能测试...")
    
    # 运行测试
    single_test_result = test_single_file(engine)
    print()  # 空行
    batch_test_result = test_batch_conversion(engine)
    
    print_separator()
    print("测试完成总结:")
    
    if single_test_result or batch_test_result:
        print("✅ 测试通过! 程序可以正常工作")
        print("\n使用方法:")
        print("  1. 转换单个文件: python doc_to_pdf.py 文件路径")
        print("  2. 批量转换: python doc_to_pdf.py 目录路径")
        print("  3. 指定转换引擎: python doc_to_pdf.py 路径 --engine=libreoffice 或 --engine=word")
        print("  4. 批量转换并将所有PDF保存到根目录: python doc_to_pdf.py 目录路径 --output-to-root")
        print("\n详细使用说明请查看 README.md 文件")
        print("\n如果您需要安装LibreOffice，请查看 libreoffice_install_guide.md 文件")
    else:
        print("⚠️  所有测试都被跳过或失败")
        print("请检查您的环境配置，确保已安装至少一种转换引擎")
        print("  - Microsoft Word")
        print("  - 或 LibreOffice (免费开源替代方案)")
        print("\n如有错误，请查看 conversion.log 文件获取详细信息")
        print("\nLibreOffice安装指南: libreoffice_install_guide.md")
    
    print_separator()
    input("按回车键退出...")
    return 0


if __name__ == "__main__":
    sys.exit(main())