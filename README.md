# offline-office-docs

面向 **Windows 内网环境**（无外网、无 npm、无 Python）的 ZCode 桌面端离线办公文档处理插件。

纯系统自带 **PowerShell 5.1 + .NET Framework** 实现，**零依赖、免安装、无需管理员权限**，解压即用。

## 功能

| 类别 | 能力 |
|---|---|
| **Word / Excel / PPT（OOXML）** | 结构总览、文本提取、单元格读写（数字/布尔/公式自动识别）、CSV 导入导出、增删工作表、查找替换、整页改写、克隆加页，全部带自动备份 |
| **PDF（纯 PowerShell 引擎）** | 中文文本提取（按字体区分 ToUnicode CMap）、合并 / 拆分 / 提取页 / 旋转（保留资源与注释、重算 xref）、从文本生成中文 PDF |
| **Office ↔ PDF 互转** | 检测到 MS Office / WPS 时自动启用 COM 高保真转换 |

配套提供：技能（SKILL.md + 2 份参考文档）、2 个斜杠命令、空白模板（docx/pptx/xlsx）、插件市场清单。

## 仓库结构

```
offline-office-docs/
├── README.md            # 本文件
├── docs/                # 文档
│   ├── 安装说明.md       # 内网安装步骤
│   └── 使用说明.md       # 详细使用说明
├── plugin/              # 插件完整源码（可直接作为插件市场目录使用）
│   ├── marketplace.json
│   └── offline-office-docs/
│       ├── .zcode-plugin/plugin.json
│       ├── commands/    # 斜杠命令
│       ├── scripts/     # PowerShell 脚本（9 个）
│       ├── skills/      # SKILL.md + 参考文档
│       └── templates/   # 空白模板
└── dist/                # 打包产物
    └── offline-office-docs-plugin-0.1.0.zip
```

## 安装（目标 Windows 内网机器）

1. 将 `dist/offline-office-docs-plugin-0.1.0.zip` 解压到本机任意目录（如 `D:\zcode-ext\`），得到 `plugins/` 目录。
2. 打开 ZCode → **插件市场** → **添加 / Add** → **添加插件市场 / Add Plugin Marketplace**。
3. 选择解压出来的 `plugins` 目录（其下应有 `marketplace.json`）。
4. 添加成功后进入 **个人 / Personal**，找到市场 `dev-offline-office-local` 下的 **离线办公文档处理 (Offline Office Docs)**，点击 **安装 / Install**。
5. 新建任务，在技能/插件选择器中选中 `offline-office-docs`，发送文档处理指令即可使用。

## 快速验收示例

> 读取 D:\docs\报告.docx 的结构，然后提取 D:\docs\报告.pdf 的全部文字

预期：输出 docx 的段落/表格统计和 PDF 的中文文本内容，全程无需联网。

## 环境要求

- Windows 10/11 或 Server 2016+（自带 Windows PowerShell 5.1 与 .NET Framework 4.x）
- 无需外网、无需 npm、无需 Python、无需管理员权限
- 可选：装有 MS Office 或 WPS Office 时，额外解锁 Office↔PDF 高保真互转

## 版本

- v0.1.0：四个文档链路（docx/xlsx/pptx/pdf）全部通过 pwsh 真实运行测试，并经 LibreOffice 打开校验（中文内容、公式计算、页面顺序、引用完整性均正确）。
