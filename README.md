# zcode-toolbox

ZCode 桌面端工具集：**离线办公文档处理插件** + **内网适配修复包**，内网/外网环境均可用。

## 适用场景速览

| 场景 | offline-office-docs 插件 | zcode-intranet-fixes 适配包 |
|---|---|---|
| **外网环境**（可访问 z.ai / 官方云端） | ✅ 可正常安装使用（功能完整；外网下属"可选增强"，官方在线插件也能做到） | ⚠️ **一般无需使用**（CUA、动态工作流等默认可用，本包用于排查/主动加固时参考） |
| **内网环境**（无外网 / 被隔离） | ✅ **必需**（零依赖纯 PowerShell，是内网唯一的文档处理通道） | ✅ **必需**（解锁被内网阻断的 CUA、动态工作流等能力） |

> 判断要点：**外网 = 一切正常，工具集只是增强；内网 = 功能被禁用/不可达，工具集是解锁方案。**

---

## 模块清单

| 模块 | 目录 | 适用场景 | 说明 |
|---|---|---|---|
| **离线办公文档处理插件** | `plugin/` + `docs/` | 内网必需 / 外网可选 | 纯 PowerShell 5.1 + .NET 实现 docx/xlsx/pptx/pdf 处理，零依赖、解压即用 |
| **内网适配修复包** | `intranet-fixes/` | 专为内网设计（外网排查时参考） | CUA 电脑操控灰显修复 + 动态工作流解锁方案矩阵 + 内网能力全面检查报告 |

---

## 1. 离线办公文档处理插件（offline-office-docs）

**适用场景：内网必需 / 外网可选。** 面向 ZCode 桌面端的离线办公文档处理插件，在**无外网、无 npm、无 Python 的内网环境**下解决文档处理问题；外网环境安装同样可用（此时作为官方在线能力之外的本地增强）。

纯系统自带 **PowerShell 5.1 + .NET Framework** 实现，**零依赖、免安装、无需管理员权限**，解压即用。

### 功能

| 类别 | 能力 |
|---|---|
| **Word / Excel / PPT（OOXML）** | 结构总览、文本提取、单元格读写（数字/布尔/公式自动识别）、CSV 导入导出、增删工作表、查找替换、整页改写、克隆加页，全部带自动备份 |
| **PDF（纯 PowerShell 引擎）** | 中文文本提取（按字体区分 ToUnicode CMap）、合并 / 拆分 / 提取页 / 旋转（保留资源与注释、重算 xref）、从文本生成中文 PDF |
| **Office ↔ PDF 互转** | 检测到 MS Office / WPS 时自动启用 COM 高保真转换 |

配套提供：技能（SKILL.md + 2 份参考文档）、2 个斜杠命令、空白模板（docx/pptx/xlsx）、插件市场清单。

### 安装（目标 Windows 内网机器）

1. 将 `dist/offline-office-docs-plugin-0.1.0.zip` 解压到本机任意目录（如 `D:\zcode-ext\`），得到 `plugins/` 目录。
2. 打开 ZCode → **插件市场** → **添加 / Add** → **添加插件市场 / Add Plugin Marketplace**。
3. 选择解压出来的 `plugins` 目录（其下应有 `marketplace.json`）。
4. 添加成功后进入 **个人 / Personal**，找到市场 `dev-offline-office-local` 下的 **离线办公文档处理 (Offline Office Docs)**，点击 **安装 / Install**。
5. 新建任务，在技能/插件选择器中选中 `offline-office-docs`，发送文档处理指令即可使用。

### 快速验收示例

> 读取 D:\docs\报告.docx 的结构，然后提取 D:\docs\报告.pdf 的全部文字

预期：输出 docx 的段落/表格统计和 PDF 的中文文本内容，全程无需联网。

### 环境要求

- Windows 10/11 或 Server 2016+（自带 Windows PowerShell 5.1 与 .NET Framework 4.x）
- 无需外网、无需 npm、无需 Python、无需管理员权限
- 可选：装有 MS Office 或 WPS Office 时，额外解锁 Office↔PDF 高保真互转

### 版本

- v0.1.0：四个文档链路（docx/xlsx/pptx/pdf）全部通过 pwsh 真实运行测试，并经 LibreOffice 打开校验（中文内容、公式计算、页面顺序、引用完整性均正确）。

---

## 2. 内网适配修复包（zcode-intranet-fixes）

**适用场景：专为内网设计。** 外网环境下 CUA 电脑操控、动态工作流等功能默认可用，**一般无需使用本包**；仅在需要排查/主动加固时参考。本包解决的是内网（云端不可达）导致的以下三类问题：

- **CUA 电脑操控灰显修复**：Windows 实为首类支持平台，灰显三因（入口隐藏 `computerUseComposerEntryHidden`、内核开关 `ZCODE_CUA_PRODUCT_HELPER` 未放行、内置插件 computer-use 未启用）+ 分步开启清单- **动态工作流解锁方案矩阵**（10 条方案，已实测）：
  - 方案 A ★：CLI `--workflow-mode onDemand|alwaysOn`（零修改）
  - 方案 B ★★：`ZCODE_BASE_URL` + 企业 CA + 内网网关伪造灰度响应（**纯配置解锁桌面 Composer**）
  - 方案 C ★：官方 remote 运行时下发 remote.mode
  - 方案 D：asar 偏移补丁（host 侧 wa() 兜底）——用户内网实测有效
  - 方案 E：散装 zcode.cjs `Ufr` 补丁（现场实测确认有效）
  - 另有：UI 文件补丁（实测组合之一）、官方 Hooks 不可行性论证等
- **内网能力全面检查**：受限项清单 + 官方域名放行清单 + 代理/CA 注入 env 清单

### 使用

```powershell
# 1. 环境检查（探测 CUA 开关 / 动态工作流锚点 / 域名可达性）
powershell -ExecutionPolicy Bypass -File scripts/check-intranet.ps1

# 2. 启用动态工作流（首选散装 zcode.cjs，备选 asar，自动建计划任务自愈）
powershell -ExecutionPolicy Bypass -File scripts/enable-dynamic-workflow.ps1

# 3. 字节补丁工具（apply / rollback / verify）
python patches/apply-patch.py --file <目标文件> --offset <偏移> --new onDemand
```

---

## 仓库结构

```
zcode-toolbox/
├── README.md                # 本文件（仓库总览）
├── docs/                    # offline-office-docs 插件文档
│   ├── 安装说明.md
│   └── 使用说明.md
├── plugin/                  # offline-office-docs 插件完整源码
│   ├── marketplace.json
│   └── offline-office-docs/（脚本/技能/模板/命令）
├── intranet-fixes/          # zcode-intranet-fixes 内网适配修复包
│   ├── REPORT.md            # 完整取证报告（CUA + 动态工作流 + 全面检查）
│   ├── README.md
│   ├── config/              # 配置说明（代理/CA/env）
│   ├── scripts/             # check-intranet.ps1 / enable-dynamic-workflow.ps1
│   ├── patches/             # apply-patch.py / disabled-to-ondemand.patch
│   └── logs-backup/         # 取证日志备份
└── dist/                    # 打包产物
    ├── offline-office-docs-plugin-0.1.0.zip
    └── zcode-intranet-fixes-plugin-0.1.0.zip
```
