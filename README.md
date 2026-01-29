# 多功能文档转换工具 v4.2

> 一款基于 Python + Tkinter 的**本地离线多功能文档批量转换工具**，支持图片、Word、Excel、PDF、网页等多种格式之间互转，尽量保留原始样式。

---

## ✨ 功能特性概览

- **图片批量转换**
  - 图片 → PDF（多张图片合并为 1 个 PDF）
  - 图片 → PPT（每张图片 1 页，自动等比居中铺满）
- **文档批量转换**
  - Word / WPS / LibreOffice 文档 → PDF
  - 支持多文件批量选择输出
- **表格批量转换**
  - Excel / WPS / LibreOffice 表格 → PDF
  - 支持多文件批量选择输出
- **网页批量转换**
  - URL → PDF（使用 Edge / Chrome + WebDriver 渲染）
  - 支持多行 URL 批量转换
- **PDF 批量处理**
  - PDF → PPT（每页转为图片嵌入 PPT）
  - PDF → 图片（按页导出为 PNG / JPG）
  - PDF → 提取原始内嵌图片
- **图片压缩与格式转换**
  - 多张图片 / 整个文件夹 → WebP
  - 支持设置输出质量与缩放比例
- **任务控制与日志**
  - 支持任务：暂停 / 继续 / 终止
  - 实时进度条与数量 / 百分比显示
  - 可滚动操作日志窗口，支持一键清空
- **浏览器驱动管理**
  - 自动在本机常见路径中查找 WebDriver
  - 提供详细的浏览器版本与驱动下载说明
  - 支持在线自动下载安装（需 webdriver-manager）

---

## 🧩 支持的转换类型

- 图片相关：
  - 输入：`*.png, *.jpg, *.jpeg, *.bmp, *.gif, *.tiff, *.webp`
  - 输出：PDF、PPT、WebP

- 文档相关：
  - 输入：`*.docx, *.doc, *.wps, *.rtf`
  - 输出：PDF

- 表格相关：
  - 输入：`*.xlsx, *.xls, *.csv`
  - 输出：PDF

- PDF 相关：
  - 输入：`*.pdf`
  - 输出：
    - PPT（`*.pptx`）
    - 图片（`*.png` / `*.jpg`）
    - 提取的原始图片（保持原始格式 `png/jpg/...`）

- 网页相关：
  - 输入：URL（每行一个）
  - 输出：PDF

---

## 💻 运行环境要求

- 操作系统：**Windows**（代码中大量使用 `win32com` / `winreg` / `*.exe`）
- Python 版本：**Python 3.7+**（推荐 3.8+）
- 必需第三方依赖（启动前必须安装）：
  - Pillow（图像处理）
  - python-pptx（PPT 读写）
  - PyMuPDF（fitz，PDF 渲染与图像提取）
- 可选依赖（对应功能才需要）：
  - pywin32（win32com，调用 Office / WPS）
  - selenium（网页渲染）
  - webdriver-manager（自动下载浏览器驱动，可选）

---

## 📦 依赖安装

在命令行（CMD / PowerShell）中执行：

1. 安装核心依赖：

    pip install Pillow python-pptx PyMuPDF

2. 如需使用 Office / WPS 进行文档 & 表格 → PDF（推荐安装）：

    pip install pywin32

3. 如需使用“网页 → PDF”功能：

    pip install selenium webdriver-manager

> 提示：程序启动时会自动检查 `Pillow / python-pptx / PyMuPDF`，缺失会直接退出并提示安装命令。

---

## 🚀 启动方式

1. 将本脚本保存为：
   - 例如：`converter_gui.py` 或任意 `.py` 文件名

2. 在脚本所在目录打开命令行，执行：

    python converter_gui.py

3. 启动后将出现带有多个分页（Tab）的图形界面，标题为：

    📄 多功能文档转换工具 v4.2 - 支持批量多文件

---

## 🖼 图片转换 Tab（“🖼️ 图片转换”）

### 1. 选择图片

- 点击：
  - “➕ 添加文件”：批量选择多张图片
  - “🗑️ 清除全部”：清空列表
  - “❌ 删除选中”：删除列表中选中的图片行
- 支持拖动排序：
  - “⬆️ 上移”：将选中项上移
  - “⬇️ 下移”：将选中项下移
- 下方会显示：`已选择: X 个文件`

### 2. 质量设置

- “⚙️ 质量设置”中可选：
  - 高（300 DPI）
  - 中（150 DPI）
  - 低（72 DPI）
- 内部对应不同的分辨率与压缩质量。

### 3. 图片 → PDF

- 按钮：`📄 转为 PDF`
- 流程：
  1. 选择输出 `.pdf` 文件路径
  2. 程序将按照列表顺序，将图片合并为一个 PDF
  3. 支持透明 PNG（自动转白底）、模式转换为 RGB

### 4. 图片 → PPT

- 按钮：`📊 转为 PPT`
- 流程：
  1. 选择输出 `.pptx` 文件路径
  2. 每张图片生成 1 张幻灯片
  3. 自动等比缩放，最大化居中铺满 16:9 画布（13.333 x 7.5 英寸）
  4. 可选择压缩质量（中 / 低时会缩小分辨率）

---

## 📝 文档转换 Tab（“📝 文档转换”）

### 支持格式

- `*.docx, *.doc, *.wps, *.rtf`

### 使用说明

1. 添加文档：
   - “➕ 添加文件”：多选 Word / WPS 文档
   - “🗑️ 清除全部”、“❌ 删除选中”：管理列表
2. 输出路径：
   - 点击“📄 批量转为 PDF”后，选择一个输出文件夹
3. 转换逻辑：
   - 优先使用：
     - Microsoft Word（win32com: `Word.Application`）
   - 其次尝试：
     - WPS（`KWPS.Application` 或 `KET.Application`）
   - 再其次：
     - LibreOffice（命令行 `soffice --headless --convert-to pdf`）
4. 输出文件名：
   - 与原文件名保持一致，仅扩展名变为 `.pdf`

---

## 📊 表格转换 Tab（“📊 表格转换”）

### 支持格式

- `*.xlsx, *.xls, *.csv`

### 使用说明

1. 添加表格文件：
   - “➕ 添加文件”多选 Excel / CSV 文件
2. 点击“📄 批量转为 PDF”
3. 转换逻辑优先级：
   - Microsoft Excel（`Excel.Application` → `ExportAsFixedFormat`）
   - WPS 表格（`KET.Application` / `ET.Application`）
   - LibreOffice（`soffice --convert-to pdf`）
4. 输出文件夹中生成与源表格同名的 `.pdf`

---

## 🌐 网页转换 Tab（“🌐 网页转换”）

### 输入 URL

- 文本框中：
  - 每行一个 URL
  - 可不写协议，程序会自动补全为 `https://`
  - 示例默认值：`https://example.com`
- 常见格式示例：
  - `https://www.example.com`
  - `www.example.com`
  - `example.com/path?a=1`

### 转换为 PDF

- 按钮：
  - “🌐 批量网页 → PDF”
  - “🗑️ 清除”：清空 URL 列表
- 当 URL 数量：
  - 为 1 个：弹出“保存 PDF”对话框
  - 大于 1 个：弹出“选择输出文件夹”对话框，多文件批量输出

### 渲染与转换方式

- 使用 Selenium + Edge/Chrome 无头浏览器进行渲染
- 优先顺序：
  1. 已找到的 Edge WebDriver（`msedgedriver.exe`）
  2. 已找到的 Chrome WebDriver（`chromedriver.exe`）
  3. 在线自动下载 WebDriver（需要 `webdriver-manager`）
- 截取方式：
  - 首选：通过 CDP 命令 `Page.printToPDF` 直接生成 PDF（保留文字和样式）
  - 失败时：退回为整页截图 → 转为 PDF（内容为图片）

### 文件命名规则

- 根据 URL 自动生成（截断为 <= 50 长度）：
  - 主机名与路径组成，非字母数字字符会被替换为下划线
  - 例如：
    - `https://www.example.com/a/b` →
      `www_example_com_a_b.pdf`

---

## 📄 PDF 转换 Tab（“📄 PDF转换”）

### 选择 PDF 文件

- 支持多选：
  - “➕ 添加文件”：多选 PDF
  - “🗑️ 清除全部”、“❌ 删除选中”：管理列表
- 下方显示当前选择文件数量

### 设置区域

- DPI：
  - `72 ~ 600` 可调，默认 `150`
  - 影响 PDF → 图片 / PDF → PPT 时的清晰度与文件大小
- 输出图片格式：
  - PNG（无损、清晰度高）
  - JPG（有损压缩、体积更小）

### 批量 PDF → PPT

- 按钮：`📊 批量PDF→PPT`
- 当仅选择 1 个 PDF：
  - 弹出“保存 PPT”对话框
- 当选择多个 PDF：
  - 选择“输出文件夹”，每个 PDF 转为一个独立的 `.pptx`
- 内部流程：
  1. 使用 PyMuPDF（fitz）逐页渲染为图片（指定 DPI）
  2. 每页 1 张图片，按比例缩放并居中到 PPT（16:9）
  3. 按原始页序生成 PPT

### 批量 PDF → 图片

- 按钮：`🖼️ 批量PDF→图片`
- 输出说明：
  - 若仅 1 个 PDF：
    - 直接输出到所选文件夹中（内部可能按 PDF 名创子目录，由逻辑决定）
  - 多个 PDF：
    - 以每个 PDF 文件名创建子文件夹（如 `xxx_page_001.png`）
- 文件命名：
  - `原文件名_page_001.png/jpg`

### 批量提取 PDF 内嵌图片

- 按钮：`📤 批量提取图片`
- 对每个 PDF：
  - 遍历所有页面，提取所有内嵌图片资源
  - 保留原始图片格式（通过 `ext` 字段获取后缀）
  - 文件命名：
    - `image_page{页号}_{序号}.{ext}`

---

## 🖼 WebP Tab（“🖼️ WebP”）

### 输入模式选择

- “选择图片文件”模式：
  - 使用“浏览”按钮选择多张图片文件
  - 文本框中显示类似“已选择 X 个文件”
- “选择文件夹”模式：
  - 使用“浏览”按钮选择一个文件夹
  - 将统计该文件夹中图片数量（按照常见图片后缀）

### 设置参数

- 质量（Quality）：
  - 通过滑块设置 `1 ~ 100`，默认 `85`
  - 右侧实时显示当前质量值
- 尺寸缩放（Resize）：
  - 100%（不缩放）
  - 75%（按 0.75 倍缩放）
  - 50%（按 0.5 倍缩放）

### 开始转换

- 按钮：`🔄 转换为 WebP`
- 流程：
  1. 选择输出文件夹
  2. 若是“文件模式”：将所选所有文件输出为同名 `.webp`
  3. 若是“文件夹模式”：扫描文件夹中所有图片进行转换
  4. 支持暂停 / 继续 / 终止控制

---

## ⏯ 任务控制与日志

### 任务控制按钮（底部）

- `⏸️ 暂停`
  - 暂停当前长任务（如：多页 PDF、批量转换）
  - 内部通过 `threading.Event` 控制循环等待
- `▶️ 继续`
  - 恢复暂停的任务
- `⏹️ 终止`
  - 触发停止标记，当前任务尽快安全结束

### 日志区域

- 标题：“📋 操作日志”
- 支持：
  - 实时输出每一步操作信息
  - 自动滚动到最新内容
  - 按钮“🗑️ 清除日志”：一键清空

---

## 🔧 浏览器驱动与帮助 Tab（“❓ 帮助”）

### 浏览器检测

启动时会自动检测：

- Microsoft Edge：
  - 执行路径
  - 版本号（通过 `winreg` 读取 `BLBeacon`）
- Google Chrome：
  - 执行路径
  - 版本号（通过 `winreg` 读取 `BLBeacon`）
- 本地 WebDriver：
  - `msedgedriver.exe`
  - `chromedriver.exe`
- 常见搜索路径：
  - 程序所在目录
  - `C:\WebDriver\`
  - `C:\Drivers\`
  - `C:\Program Files\WebDriver\`
  - `C:\Program Files (x86)\WebDriver\`
  - 用户目录、桌面、下载目录等
  - 系统 PATH 中所有路径

### 帮助 Tab 内容

- 展示：
  - 当前检测到的浏览器版本
  - Edge / Chrome 官方驱动下载地址
  - 将驱动放置到何处可被程序自动识别
- 按钮：
  - “🔄 重新检测驱动”：重新扫描本地 WebDriver
  - “📂 打开程序目录”：打开当前脚本所在文件夹

---

## ⚠️ 注意事项与常见问题

1. Office / WPS 未安装怎么办？
   - 文档 / 表格 → PDF 仍可通过 LibreOffice 实现（需自行安装 LibreOffice）
   - 若系统完全无 Office / WPS / LibreOffice，将无法进行文档 / 表格 → PDF 转换

2. 启动时报缺少模块：
   - 控制台会显示需要安装的包
   - 按提示执行 `pip install 对应包名` 后重新运行

3. 网页 → PDF 无法使用：
   - 需确保：
     - 已安装 `selenium`
     - 已安装支持的浏览器（Edge 或 Chrome）
     - 已在常见路径放置对应的 WebDriver，或安装 `webdriver-manager` 允许自动下载
   - 可在“❓ 帮助”页查看驱动说明和安装位置建议

4. 暂停 / 终止不立即生效？
   - 任务控制是在循环中定期检查标记，某些长耗时步骤（如单页渲染）仍需要等待该步骤结束才会响应

5. DPI 与清晰度的关系：
   - PDF → PPT / 图片 时：
     - DPI 越高，图像越清晰，文件也越大
     - 一般 150~200 DPI 即可满足大部分显示与打印需求

---

## 📁 代码结构概览

- `TaskController`：
  - 统一任务控制（暂停 / 继续 / 停止）状态管理
- `BrowserDriverManager`：
  - 搜索本地 WebDriver
  - 读取浏览器版本
  - 提供驱动下载说明文本
- `DocumentConverter`：
  - 核心转换逻辑封装：
    - 文档 / 表格 / 网页 / 图片 / PDF 等所有转换函数
  - 所有耗时操作均支持进度回调与任务控制
- `ConverterGUI`（Tkinter GUI）：
  - 多标签页 UI
  - 文件列表、参数设置、按钮事件
  - 与 `DocumentConverter` 解耦，通过回调连接日志与进度

---

## 📝 版本信息（v4.2）

- 新增：
  - **支持批量多文件选择**（文档、表格、PDF 等）
  - **任务控制功能**：暂停、继续、终止
- 改进：
  - **浏览器驱动管理**：修复驱动问题，支持离线查找
  - 较完整地**保留原始样式**（依赖对应引擎：Word / Excel / 渲染浏览器 / PDF 引擎）
- 其他：
  - 更多日志输出与错误提示
  - 更友好的 GUI 布局与中文提示

---
