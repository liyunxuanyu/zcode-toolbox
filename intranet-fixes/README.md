# ZCode 内网环境适配包 v0.1.1

**适用场景：专为内网（无外网）设计。** 外网环境下 CUA 电脑操控、动态工作流等功能默认可用，一般无需使用本包；本包解决 ZCode 桌面端在**内网（无外网）**环境下的三类问题：
**CUA（电脑操控）灰显**、**动态工作流不可用**、**其它云端依赖功能受限**。
与 offline-office-docs 插件**完全解耦**，不修改其任何文件。

## 快速使用

> ✅ **2026-10-16 晚第三轮复核：脚本 P0/P1 全部关闭**（REPORT.md 6.5）。此前"不要用 -UseAsarOnly"的规避提示已作废——asar 分支现为"先校验偏移、后备份、不符即拒"，全流程无盲替换。
>
> 1. **桌面 Composer 生效位在 host 侧（app.asar）**；内核散装 `zcode.cjs` 的 `Ufr` 档位只对 CLI 直连兜底有效（第三轮取证：桌面会话由 host 显式下发 disabled，覆盖内核默认档，见 REPORT.md 7.0/7.3 方案E）。
> 2. 偏移 `280109971` 是实测那台机器的**安装特定值**（同代 Linux 包同偏移处为其他内容）——换机/升级后必须先定位再打补丁：用唯一特征串 `).mode:"disabled"` 定位 host 侧 wa() 兜底（标准流程见 REPORT.md 7.2）。
> 3. **纯配置替代（推荐先试，不改二进制）**：桌面 Composer 可用 `ZCODE_BASE_URL` + 企业 CA + 内网网关伪造灰度响应解锁（REPORT.md 7.3 方案B）；CLI/脚本会话直接 `zcode --workflow-mode alwaysOn` 零修改（方案A）。
> 4. 官方 Hooks 不能解锁动态工作流门禁（REPORT.md 7.1）；官方 Hooks/工作流/门禁的关系与完整方案矩阵见 REPORT.md 第七节。

### 1. 先做体检（只读，安全）
```powershell
powershell -ExecutionPolicy Bypass -File scripts\check-intranet.ps1
```
输出各项功能的内网状态与建议（含 asar 偏移 8 字节精确探测、灰度接口指向检查、检出汇总表）。建议用 64 位 PowerShell。

### 2. 开启 CUA 入口（安全可逆）
按 `config\README.md` 第 1 节，将 `~/.zcode/v2/setting.json` 的
`computerUseComposerEntryHidden` 改为 `false`，重启 ZCode。

### 3. 启用动态工作流（两路补丁，先校验后备份，不符即拒）
```powershell
powershell -ExecutionPolicy Bypass -File scripts\enable-dynamic-workflow.ps1
```
- 首选 `resources\glm\zcode.cjs` 锚点 `Ufr="disabled"` → `Ufr="onDemand"`（仅当全文件唯一 count=1 才替换；幂等跳过；主要惠及 CLI 场景）
- 备选 `app.asar` 偏移 280109971 处 8 字节（host 侧 wa() 兜底，桌面 Composer 生效位）：seek 校验 disabled → 备份 → 定点写 → 复验；已是 onDemand 幂等跳过；其它内容报错退出
- 自动备份到 `logs-backup\`（按文件种类分别回滚：`... -Rollback`）
- ⚠️ 跨机/跨版本先跑体检或按 REPORT.md 7.2 重新定位偏移，再改脚本顶部 `$asarOffset`
- 纯配置方案（方案A/B）见 REPORT.md 第七节——不需要打补丁

> 若目标机有 Python 3，也可用 `patches\apply-patch.py`（apply/rollback/verify，定点失败报错退出）。

### 4. 内网其它功能离线方案
详见 `REPORT.md` 第三节：插件市场/更新/遥测/图片搜索等受限项的离线处理方式。

## 目录结构
```
zcode-intranet-fixes/
├── REPORT.md              # 调查报告（判定链 + 证据 + 方案矩阵）
├── README.md              # 本文件
├── config/                # 配置建议（CUA 入口、供应商、自动更新）
├── scripts/
│   ├── check-intranet.ps1          # 内网功能体检（只读）
│   └── enable-dynamic-workflow.ps1 # 动态工作流一键启用（校验/备份/定点补丁/回滚/计划任务）
├── patches/
│   ├── disabled-to-ondemand.patch  # 字节补丁说明
│   └── apply-patch.py              # Python 补丁工具（apply/rollback/verify）
└── logs-backup/           # 补丁前备份（脚本自动写入）
```

## 证据可信度（第三轮复核后更新）
| 结论 | 来源 |
|---|---|
| CUA 入口默认隐藏 / 插件依赖 / **Windows 产品 helper 通道（非平台硬禁）** | ZCode 3.15.3 包内源码核实（REPORT.md 1.4） |
| 动态工作流 env 被 delete、isPreview 烙死、内核默认档位 disabled | ✅ 同代包逐字核实 |
| **灰度接口 `GET {ZCODE_BASE_URL\|默认 zcode.z.ai}/api/v1/client/configs`，消费 `data.configs.dynamicWorkflow`** | ✅ out/host/index.js 逐字核实（REPORT.md 7.0） |
| **偏移 280109971 = host 侧 wa() 兜底 disabled**（定位特征串 `).mode:"disabled"` 全文件唯一） | ✅ host bundle 逐字核实 + 用户实测交叉印证（REPORT.md 7.2） |
| **ZCODE_BASE_URL 用户级可重定向灰度接口、ZCODE_DYNAMIC_WORKFLOW_MODE 被删** | ✅ out/main/index.js buildHostProcessEnv 逐字核实 |
| **桌面会话由 host 显式下发 disabled（覆盖内核默认档）** | ✅ 内核 Lai/envelope 链逐字核实（REPORT.md 7.0 推论2） |
| TLS 标准校验无 pinning；`ZCODE_AGENT_CA_CERT`/`NODE_EXTRA_CA_CERTS` 可注入企业 CA | ✅ 内核/host 逐字核实 |
| 官方 Hooks 不能影响动态工作流门禁 | ✅ 官方文档能力清单 + 源码注册期排除（REPORT.md 7.1） |
| 工作流编排为本地能力（.workflow.js 落盘、原语 agent/parallel/pipeline/phase） | ✅ 内核逐字核实（REPORT.md 7.4） |

## 注意
- 字节补丁是对本地安装程序的开关调整；版本升级后偏移/锚点可能变化，先跑体检重新定位
- 已打补丁的机器保持自动更新关闭（计划任务自愈兜底）；`... -Rollback` 按文件种类回滚
- CUA 在 Windows 并非平台硬禁：按 REPORT.md 1.4 步骤开启入口 + 确保 helper/插件开关
- 本包所有文件均为新增独立产物，未触碰 offline-office-docs 目录
