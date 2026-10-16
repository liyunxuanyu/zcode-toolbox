# ZCode 内网环境适配包 v0.1.0

本包解决 ZCode 桌面端在**内网（无外网）**环境下的三类问题：
**CUA（电脑操控）灰显**、**动态工作流不可用**、**其它云端依赖功能受限**。
与 offline-office-docs 插件**完全解耦**，不修改其任何文件。

## 快速使用

### 1. 先做体检（只读，安全）
```powershell
powershell -ExecutionPolicy Bypass -File scripts\check-intranet.ps1
```
输出各项功能的内网状态与建议。

### 2. 开启 CUA 入口（安全可逆）
按 `config\README.md` 第 1 节，将 `~/.zcode/v2/setting.json` 的
`computerUseComposerEntryHidden` 改为 `false`，重启 ZCode。

### 3. 启用动态工作流（字节补丁方案，带备份/回滚/计划任务）
```powershell
powershell -ExecutionPolicy Bypass -File scripts\enable-dynamic-workflow.ps1
```
- 自动备份到 `logs-backup\`
- 偏移 280109971 处 `disabled` → `onDemand`（8 字节等长替换）
- 注册计划任务自愈，防自动更新覆盖
- 回滚：`... -Rollback`

> 若目标机有 Python 3，也可用 `patches\apply-patch.py`（apply/rollback/verify）。

### 4. 内网其它功能离线方案
详见 `REPORT.md` 第三节：插件市场/更新/遥测/图片搜索等受限项的离线处理方式。

## 目录结构
```
zcode-intranet-fixes/
├── REPORT.md              # 调查报告（判定链 + 证据 + 方案）
├── README.md              # 本文件
├── config/                # 配置建议（CUA 入口、供应商、自动更新）
├── scripts/
│   ├── check-intranet.ps1          # 内网功能体检
│   └── enable-dynamic-workflow.ps1 # 动态工作流一键启用（备份/补丁/回滚/计划任务）
├── patches/
│   ├── disabled-to-ondemand.patch  # 字节补丁说明
│   └── apply-patch.py              # Python 补丁工具（apply/rollback/verify）
└── logs-backup/           # 补丁前备份（脚本自动写入）
```

## 证据可信度
| 结论 | 来源 |
|---|---|
| CUA 平台判定 / 入口默认隐藏 / 插件依赖 | ZCode 3.15.3 包内源码核实 |
| 动态工作流 env 被删、isPreview 烙死、云端灰度 disabled、偏移 280109971 | 用户内网实测截图 + ZCode 同源代码取证 |

## 注意
- 字节补丁是对本地安装程序的开关调整；版本升级后偏移可能变化，先跑体检重新定位
- 本包所有文件均为新增独立产物，未触碰 offline-office-docs 目录
