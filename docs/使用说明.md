# offline-office-docs — Windows 内网离线办公文档处理插件

在**无外网、无 npm、无 Python** 的 Windows 内网机器上，用 ZCode 桌面端直接处理办公文档。

纯 **Windows PowerShell 5.1 + .NET Framework**（系统自带，零安装），无需管理员权限。

## 能力一览

| 格式 | 读取 | 编辑 | 创建 | 转换 |
|---|---|---|---|---|
| Word .docx | ✅ 结构/全文/段落列表 | ✅ 全文替换/追加/查找替换 | ✅ 从模板+文本 | ✅ 转 PDF（需 Office/WPS） |
| Excel .xlsx | ✅ 工作表/单元格/CSV 导出 | ✅ 单元格读写/CSV 导入/加行/增删工作表 | ✅ 新工作表 | ✅ 转 PDF（需 Office/WPS） |
| PPT .pptx | ✅ 幻灯片列表/文本提取 | ✅ 全局替换/整页改写/加页 | ✅ 克隆加页 | ✅ 转 PDF（需 Office/WPS） |
| PDF | ✅ 页数/文本提取（**支持中文**） | ✅ 合并/拆分/提取页/旋转 | ✅ 从文本生成（中文 OK） | ✅ 转 Word（需 Office 2013+/WPS） |

**纯 PowerShell PDF 引擎**（无需任何 Office）：
- 文本提取：解析 ToUnicode CMap（按字体区分，中文不乱码）、按视觉行组装
- 结构操作：合并/拆分/提取/旋转——保留内容流、资源、注释、继承的 MediaBox/Rotate，重算 xref（源文件 xref 损坏也能读）
- 生成：UTF-8 文本 → 分页 PDF；中文使用非嵌入 STSong-Light CID 字体

**装有 MS Office 或 WPS 时**，通过 COM 自动化做高保真互转（`Convert-OfficePdf.ps1`，自动探测引擎）。

## 安装（ZCode 桌面端）

1. 打开 **插件市场 / Plugin Marketplace**
2. **添加 / Add → 添加插件市场 / Add Plugin Marketplace**，粘贴本插件所在的市场根目录：
   ```
   <解压目录>/plugins
   ```
   （该目录下有 `marketplace.json`）
3. 打开 **个人 / Personal** → 找到市场 **dev-offline-office-local** → 插件 **离线办公文档处理 (Offline Office Docs)**
4. 点击 **安装 / Install**
5. 之后可在 **设置 → 插件** 管理启用状态

## 使用

安装后新建任务，在输入框选择该插件的技能（offline-office-docs），直接用自然语言下指令，例如：

- 「读取 D:\docs\报告.docx 的结构和前 10 段」
- 「把这个 Excel 第 1 个工作表导出成 CSV，改完再导回去」
- 「提取 report.pdf 第 2-4 页的文字」
- 「把这 3 个 PDF 合并成一个」
- 「新建一份 Word，内容如下：……」

也可以直接让 ZCode 调用脚本（代理会自动定位已安装路径）：

```powershell
# 查看任意文档结构
powershell -NoProfile -ExecutionPolicy Bypass -File <scripts>\Read-DocInfo.ps1 -Path "D:\docs\报表.xlsx"

# Excel 写单元格（自动识别数字/文本/布尔/公式）
powershell -NoProfile -ExecutionPolicy Bypass -File <scripts>\Edit-Xlsx.ps1 -Action setcells -Path "报表.xlsx" -Cells "A1=产品","B1=100","C1==B1*1.13"

# PDF 中文文本提取（带 BOM 的 UTF-8 输出）
powershell -NoProfile -ExecutionPolicy Bypass -File <scripts>\Edit-Pdf.ps1 -Action text -Path "报告.pdf" -OutputPath out.txt

# Office 转 PDF（自动选 MS Office / WPS）
powershell -NoProfile -ExecutionPolicy Bypass -File <scripts>\Convert-OfficePdf.ps1 -Action to-pdf -Path "合同.docx"
```

两个斜杠命令：`/inspect-doc <文件>` 快速查看文档结构；`/convert-pdf <文件> [输出目录]` 走 COM 转换。

## 修改安全性

- 所有修改类操作自动生成 `文件名.bak-时间戳` 备份（保留最近 5 份）
- 全程 UTF-8 处理；CSV 读取带 BOM 检测 + GB18030 回退

## 已知限制

- 老二进制格式（.doc/.xls/.ppt）不能直接编辑，先用 `Convert-OfficePdf.ps1` 转成 OOXML
- 加密 PDF / 加密文档无法处理（会明确报错）
- 扫描件（图片型）PDF 无文字层，提取不到文本（OCR 超出离线范围）
- `setcells` 写入的公式由 Excel/WPS 打开时计算，文件内无缓存值
- COM 转换需要本机装有 MS Office 或 WPS；`-Action detect` 可探测

## 目录结构

```
offline-office-docs/
├── .zcode-plugin/plugin.json   插件清单
├── skills/offline-office-docs/ 技能（SKILL.md + references/）
├── commands/                    斜杠命令（inspect-doc, convert-pdf）
├── scripts/                     全部 PowerShell 脚本
│   ├── _Common.ps1              ZIP/XML/编码共享库
│   ├── _PdfLib.Read.ps1         PDF 解析引擎
│   ├── _PdfLib.Write.ps1        PDF 写入引擎
│   ├── Read-DocInfo.ps1         文档结构总览
│   ├── Edit-Docx.ps1            Word 编辑
│   ├── Edit-Xlsx.ps1            Excel 编辑
│   ├── Edit-Pptx.ps1            PPT 编辑
│   ├── Edit-Pdf.ps1             PDF 操作
│   └── Convert-OfficePdf.ps1    COM 互转（Office/WPS）
└── templates/                   docx/xlsx/pptx 空白模板 + 空工作表 XML
```
