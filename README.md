# docx-to-pdf-converter
# Doc/Docx 转 PDF 批量转换工具

这是一个使用Python编写的文档转换工具，可以将Word文档(.doc/.docx)精准转换为PDF格式，支持单个文件转换和批量转换指定目录下的所有文档。本工具支持Microsoft Word和LibreOffice两种转换引擎，您可以根据自己的需求选择使用。

## 功能特点

- ✅ 支持.doc和.docx两种格式文件转换为PDF
- ✅ 支持单个文件转换
- ✅ 支持批量转换指定目录下的所有Word文档，包括嵌套文件夹
- ✅ 支持Microsoft Word和LibreOffice两种转换引擎
- ✅ 自动选择可用的转换引擎
- ✅ 支持手动指定使用哪种转换引擎
- ✅ 支持手动指定LibreOffice安装路径（解决默认路径无法找到的问题）
- ✅ 支持将所有转换后的PDF统一保存到根目录（方便整理）
- ✅ 详细的日志记录，方便追踪转换过程
- ✅ 健壮的错误处理，确保程序稳定运行
- ✅ 自动创建输出目录（如有需要）
- ✅ 转换完成后生成统计报告

## 环境要求

- Windows 操作系统
- Python 3.6 或更高版本
- 转换引擎（**至少安装其中一个**）：
  - Microsoft Word（可选）：提供更精确的转换质量
  - LibreOffice（可选，免费开源）：作为Word的替代选择

## 安装依赖

1. 确保已安装Python
   - 可以从 [Python官网](https://www.python.org/downloads/) 下载安装

2. 安装所需的Python库
   ```bash
   pip install pywin32
   ```
   
   *注意：pywin32仅在使用Microsoft Word作为转换引擎时需要，使用LibreOffice不需要此依赖。*

## 使用方法

### 1. 转换单个文件

```bash
python doc_to_pdf.py 文件路径.docx
```

例如：
```bash
python doc_to_pdf.py "C:\Users\用户名\Documents\示例文档.docx"
```

### 2. 批量转换目录中的文件

```bash
python doc_to_pdf.py 目录路径
```
或者
```bash
python doc_to_pdf.py --batch 目录路径
```

例如：
```bash
python doc_to_pdf.py "C:\Users\用户名\Documents\需要转换的文件夹"
```

### 3. 指定转换引擎

您可以使用`--engine`参数明确指定使用哪种转换引擎：

```bash
# 使用LibreOffice引擎
python doc_to_pdf.py 文件路径.docx --engine=libreoffice

# 使用Word引擎
python doc_to_pdf.py 文件路径.docx --engine=word

# 批量转换时指定引擎
python doc_to_pdf.py 目录路径 --engine=libreoffice
```

*注意：如果不指定引擎，程序会自动尝试先使用LibreOffice，如果不可用则尝试使用Word。*

### 4. 指定LibreOffice路径

如果程序无法自动找到LibreOffice，可以手动指定其安装路径：

```bash
# 指定LibreOffice路径
python doc_to_pdf.py 文件路径.docx --libreoffice-path="D:\LibreOffice\program\soffice.exe"

# 批量转换时指定LibreOffice路径
python doc_to_pdf.py 目录路径 --libreoffice-path="D:\LibreOffice\program\soffice.exe"
```

### 5. 将所有PDF保存到根目录

批量转换时，可以选择将所有转换后的PDF文件统一保存到根目录：

```bash
# 批量转换并将所有PDF保存到根目录
python doc_to_pdf.py 目录路径 --output-to-root

# 结合其他参数使用
python doc_to_pdf.py 目录路径 --engine=libreoffice --output-to-root
```

## 输出说明

- **转换后的PDF文件**：将保存在原文件相同目录下，使用相同的文件名但扩展名为.pdf
- **日志文件**：`conversion.log` 文件将保存在脚本所在目录，记录详细的转换过程和可能的错误

## 常见问题

### 1. 程序提示找不到win32com模块

请确保已正确安装pywin32库：
```bash
pip install pywin32
```

### 2. 转换失败或出现错误

- 确保已安装至少一种转换引擎（Microsoft Word或LibreOffice）
- 如果指定了特定引擎，请确保该引擎已正确安装
- 检查文件是否被其他程序锁定或打开
- 查看 `conversion.log` 文件获取详细错误信息

### 3. 处理大量文件时速度较慢

- 程序每处理5个文件会自动暂停1秒以释放系统资源
- 转换大型文档需要一定时间
- 使用Word转换时，建议不要在转换过程中手动操作Word程序
- 使用LibreOffice转换时，确保没有其他LibreOffice实例正在运行

## 注意事项

- 程序会根据您的选择或自动检测使用可用的转换引擎
- 使用Microsoft Word进行转换时，转换过程中请尽量避免关闭程序或操作Word
- 使用LibreOffice进行转换时，请确保没有其他LibreOffice实例正在运行
- 对于受保护的文档，可能需要先解除保护才能转换
- 两种转换引擎在处理复杂格式时可能会有细微差异，Word通常提供更精确的转换结果

## 错误排查

如果遇到问题，请先查看日志文件 `conversion.log` 获取详细信息。日志文件包含了完整的错误堆栈，可以帮助识别问题所在。

### 常见错误及解决方法

1. **"无法找到LibreOffice"**
   - 确保已安装LibreOffice
   - 检查LibreOffice是否安装在标准位置
   - 使用`--libreoffice-path`参数手动指定LibreOffice的安装路径：
     ```bash
     python doc_to_pdf.py 文件路径 --libreoffice-path="LibreOffice安装路径\program\soffice.exe"
     ```

2. **"win32com模块不可用"**
   - 确保已安装pywin32库：`pip install pywin32`
   - 或者选择使用LibreOffice作为转换引擎：`--engine=libreoffice`

3. **转换过程中程序卡住**
   - 检查是否有其他Office或LibreOffice实例正在运行
   - 对于大型文档，可能需要更长的处理时间
   - 尝试重新启动程序或计算机

4. **PDF格式与原始文档有差异**
   - 尝试使用另一种转换引擎
   - 某些复杂格式可能在转换过程中有所变化
   - 对于重要文档，建议使用Word引擎以获得更精确的结果

## 许可证

[MIT License](LICENSE)

---

祝您使用愉快！如有任何问题，请参考日志文件或重新检查您的环境配置。
