# LibreOffice 下载和安装指南

LibreOffice 是一款免费开源的办公软件套件，可以作为 Microsoft Word 的替代选择，用于将 Word 文档转换为 PDF 格式。本指南将帮助您下载、安装并配置 LibreOffice。

## 为什么选择 LibreOffice？

- **免费开源**：完全免费使用，无许可证费用
- **功能强大**：支持大多数常见的文档格式转换
- **轻量级**：相比 Microsoft Office，资源占用更少
- **多平台支持**：可在 Windows、macOS 和 Linux 上运行
- **与我们的转换工具完美配合**：可作为 Word 的替代转换引擎

## 系统要求

- **操作系统**：Windows 7 或更高版本
- **内存**：至少 2GB RAM（推荐 4GB 或更多）
- **磁盘空间**：至少 3.5GB 可用空间
- **处理器**：Intel Pentium 4 或更高版本（或同等性能的处理器）

## 下载 LibreOffice

### 方法一：从官方网站下载（推荐）

1. 打开浏览器，访问 LibreOffice 官方网站：[https://zh-cn.libreoffice.org/download/libreoffice/](https://zh-cn.libreoffice.org/download/libreoffice/)

2. 选择您需要的版本：
   - 推荐下载最新的稳定版本（通常在页面顶部）
   - 对于普通用户，选择「LibreOffice Community」版本

3. 选择您的操作系统：
   - 点击「Windows」选项

4. 选择语言：
   - 选择「简体中文」

5. 点击下载按钮开始下载安装程序

### 方法二：从我们提供的链接下载

如果您无法访问官方网站，可以尝试以下下载链接（根据您的系统选择）：

- **Windows 64位系统**：[https://download.documentfoundation.org/libreoffice/stable/latest/win/x86_64/LibreOffice_最新版本_Win_x86-64.msi](https://download.documentfoundation.org/libreoffice/stable/latest/win/x86_64/LibreOffice_最新版本_Win_x86-64.msi)
- **Windows 32位系统**：[https://download.documentfoundation.org/libreoffice/stable/latest/win/x86/LibreOffice_最新版本_Win_x86.msi](https://download.documentfoundation.org/libreoffice/stable/latest/win/x86/LibreOffice_最新版本_Win_x86.msi)

> 注意：请将「最新版本」替换为当前的版本号，如「7.6.4」

## 安装 LibreOffice

1. **找到下载的安装文件**：
   - 通常在浏览器的「下载」文件夹中
   - 文件名为类似 `LibreOffice_7.6.4_Win_x86-64.msi` 的格式

2. **运行安装程序**：
   - 双击下载的安装文件
   - 如果出现用户账户控制提示，请点击「是」允许程序进行更改

3. **开始安装**：
   - 在欢迎页面，点击「下一步」

4. **选择安装类型**：
   - 对于大多数用户，推荐选择「典型」安装
   - 如果您想自定义安装选项，可以选择「自定义」

5. **选择组件**：
   - 在自定义安装中，可以选择需要安装的组件
   - 对于我们的转换工具，至少需要安装「Writer」组件（用于处理文字文档）
   - 点击「下一步」

6. **选择安装位置**：
   - 推荐使用默认安装位置
   - 点击「安装」开始安装过程

7. **等待安装完成**：
   - 安装过程可能需要几分钟时间
   - 请勿在安装过程中关闭安装窗口

8. **完成安装**：
   - 安装完成后，点击「完成」按钮

## 验证安装

安装完成后，您可以通过以下方式验证 LibreOffice 是否正确安装：

1. 点击「开始」菜单
2. 在程序列表中查找「LibreOffice」文件夹
3. 点击打开「LibreOffice Writer」
4. 如果程序正常启动，则安装成功

## 配置 LibreOffice（可选）

### 首次运行设置

1. 首次启动 LibreOffice 时，可能会提示您选择用户界面语言和其他初始设置
2. 选择「简体中文」作为界面语言
3. 完成其他设置后，点击「确定」

### 确保 LibreOffice 可以在命令行中使用（推荐）

为了确保我们的转换工具能够正确找到 LibreOffice，请按照以下步骤操作：

1. **检查安装路径**：
   - 默认情况下，LibreOffice 通常安装在以下位置之一：
     - `C:\Program Files\LibreOffice\program\`（64位系统）
     - `C:\Program Files (x86)\LibreOffice\program\`（32位系统）

2. **验证 soffice.exe 文件**：
   - 导航到安装目录
   - 确认 `soffice.exe` 文件存在

3. **将 LibreOffice 添加到系统 PATH（可选）**：
   - 如果您的系统上 LibreOffice 安装在非标准位置，可能需要将其添加到系统 PATH 环境变量
   - 对于大多数用户，这一步不是必需的，我们的脚本会自动搜索常见的安装位置

## 测试与我们的转换工具配合使用

安装完成后，您可以使用以下命令测试 LibreOffice 是否能与我们的转换工具正常工作：

```bash
python doc_to_pdf.py 您的文档路径.docx --engine=libreoffice
```

或者直接运行：

```bash
python doc_to_pdf.py 您的文档路径.docx
```

我们的脚本会自动尝试先使用 LibreOffice 进行转换。

## 常见问题解答

### 1. 安装过程中出现错误

- 确保您以管理员身份运行安装程序
- 检查您的计算机是否满足系统要求
- 确保有足够的磁盘空间
- 尝试下载安装文件并重新安装

### 2. 我们的转换工具无法找到 LibreOffice

- 确保 LibreOffice 已正确安装
- 检查 LibreOffice 是否安装在标准位置
- 查看 `conversion.log` 文件获取详细错误信息
- 尝试重新安装 LibreOffice

### 3. LibreOffice 转换的 PDF 格式与原始文档有差异

- 某些复杂格式在转换过程中可能会有所变化
- 对于重要文档，如果需要精确的格式保留，建议使用 Microsoft Word
- 尝试在 LibreOffice 中打开原始文档，调整格式后再转换

### 4. 转换过程中程序卡住或崩溃

- 关闭所有正在运行的 LibreOffice 实例
- 重新启动您的计算机
- 对于特别大的文档，可能需要增加转换超时时间

## 卸载 LibreOffice（如果需要）

如果您想卸载 LibreOffice，可以按照以下步骤操作：

1. 点击「开始」菜单 →「设置」→「应用」
2. 在应用列表中找到「LibreOffice」
3. 点击「卸载」并按照提示完成卸载过程

## 附加资源

- **官方文档**：[https://documentation.libreoffice.org/](https://documentation.libreoffice.org/)
- **官方支持论坛**：[https://ask.libreoffice.org/](https://ask.libreoffice.org/)
- **中文社区**：[https://zh-cn.libreoffice.org/community/](https://zh-cn.libreoffice.org/community/)

---

希望本指南能帮助您成功安装和使用 LibreOffice。如有任何问题，请参考我们的主文档中的错误排查部分，或查看 LibreOffice 官方文档寻求帮助。