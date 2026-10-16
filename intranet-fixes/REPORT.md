# ZCode 桌面端内网环境适配调查报告

- 版本：v0.1.1（第三轮复核与取证更新，2026-10-16）
- 日期：2026-10-16
- 适用环境：Windows 内网（无外网）ZCode 桌面端
- 调查方式：源码取证（app.asar 解析 / 主进程与 host bundle 逐字 grep / 配置与日志分析）+ 用户内网实测截图交叉验证 + 官方文档与第三方资料核实
- 与 offline-office-docs 插件完全解耦（本目录独立，未修改其任何文件）

---

## 一、Computer Use（电脑操控 / CUA）灰显问题

### 1.1 判定链（已在 ZCode 3.15.3 官方版安装包中定位并核实）

```text
cuaEnabled 判定（主进程 bundle，out/main/…）：
  enabled = process.platform === "darwin" && serviceAuthorityMode === "desktop-local"
  → 仅 macOS + 本地桌面授权模式为 true；Linux/Windows 一律 false

辅助开关（设置层）：
  setting.json 中 computerUseComposerEntryHidden = true（默认值即 true，入口默认隐藏）

插件依赖（内置插件缓存）：
  zcode-plugins-official/computer-use/0.6.4 —— CUA 依赖该内置插件

运行时日志佐证：
  cua-pip-session wiring resolved {"enabled":false,"platform":"linux","serviceAuthorityMode":"desktop-local"}
  事件 skipReason:"service-disabled"
```

### 1.2 内网 Windows 灰显的原因（结论）

1. **平台门禁**：官方版 CUA 的启用条件是 `platform === "darwin"`，Windows 端即便网络正常也非首类支持对象，判定链天然走到 disabled；
2. **服务门禁**：`service-disabled` 事件说明 CUA 需要云端服务授权，内网无法访问该服务 → 会话服务不可用；
3. **入口隐藏**：`computerUseComposerEntryHidden=true`（默认）导致 Composer 入口不显示，界面呈现"灰显/无入口"；
4. **插件依赖**：若内置插件 computer-use 0.6.4 未随包安装到目标机（内网无法在线拉取插件），入口与功能均不可用。

### 1.3 修复方案（内网可落地）

| 措施 | 操作 | 生效方式 |
|---|---|---|
| 1. 开启 Composer 入口 | 在偏好设置/配置文件将 `computerUseComposerEntryHidden` 设为 `false` | 设置层，安全可逆 |
| 2. 确保内置插件在位 | 检查 `~/.zcode/cli/plugins/cache/zcode-plugins-official/computer-use/0.6.4/`，缺失则从外网环境整目录拷贝（插件包很小，随安装包分发） | 文件层，离线可行 |
| 3. 平台判定不可改（官方策略） | 在 Windows 上即使开启入口，`enabled` 判定仍为 false——需要本地补丁（见 1.4）或接受"入口可见但服务不可用"的降级 | 产品层 |
| 4. 替代方案 | 内网如需桌面自动化，可改用系统自带方案：Windows 上的 PowerShell UI Automation / pywinauto 类本地脚本（不依赖 ZCode 云端服务） | 独立于 ZCode |

> 注：CUA 的 `enabled:false` 属官方平台/服务策略，非纯配置问题；"灰显"最直接可改善点是入口隐藏项与插件完整性。

### 1.4 复核修订：Windows 实际存在产品 helper 通道（1.1/1.2/1.3 第 3 行需按此修正）

本机同代包取证显示，1.1 的 `darwin && desktop-local` 判定只是 **macOS 专属支路**（standalone helper / pip 会话接线，app.asar 中 `Gn=process.platform==="darwin"&&e?.serviceAuthorityMode==="desktop-local"`），**并非 Windows 的总门禁**：

```text
shouldEnableDefaultCuaProductHelper：platform === "darwin" || platform === "win32"
  → Windows 与 macOS 同为 CUA 产品 helper 的首类支持平台
helper 准入门禁（平台无关）：hasBuiltInCuaPlugin && serviceAuthorityMode === "desktop-local"
  && !hasRemoteWorkspaceIdentity && !hasInjectedResolver
Windows 权限代理（permission broker）：命名管道 \\.\pipe\zcode-cua-helper-default
  （可用 env ZCODE_CUA_PERMISSION_BROKER_SOCKET 覆盖）；桌面端有 launchWindowsCuaHelperForStatus 拉起逻辑
内核侧开关：ZCODE_CUA_PRODUCT_HELPER（不设为 0/false/off 即默认放行；显式禁用时启用插件直接报错
  "computer-use built-in plugin requires ZCODE_CUA_PRODUCT_HELPER to be enabled"）；ZCODE_CUA_DEV_MODE 为研发旁路
SDK：@zcode/zcode-cua-plugin 0.6.4，win32 → target "windows"（Windows 为受支持目标）
```

对 1.2 的两点修订：
- "平台门禁：Windows 即便网络正常也非首类支持对象、判定链天然 disabled" → **结论过严**。Windows 走 helper 通道，灰显主因是 ①入口默认隐藏（`computerUseComposerEntryHidden=true`）②`ZCODE_CUA_PRODUCT_HELPER` 被显式禁用或 helper 未随桌面进程拉起 ③内置插件被压制/未启用，而非平台硬禁；
- "服务门禁 = 需要云端服务授权" → **术语修正**：`serviceAuthorityMode==="desktop-local"` 指本地桌面授权模式（相对 `desktop-attached-remote` / `standalone-server`），不依赖云端；日志 `skipReason:"service-disabled"` 是 `enabled:false` 的结果，且该日志采集于 Linux（Linux 无 helper，故 false），**不能外推到 Windows**。

1.3 第 3 行（"平台判定不可改"）按以下 Windows 内网落地步骤替代：
1) settings 将 `computerUseComposerEntryHidden` 设为 `false`；
2) 确保用户级 env `ZCODE_CUA_PRODUCT_HELPER` 未设为 0/false/off（无需强制设 1，默认即放行）；
3) 插件列表启用内置 computer-use 插件（缓存目录见 1.3 第 2 行）；
4) 保持 desktop-local 模式（勿将 `ZCODE_SERVICE_AUTHORITY_MODE` 设为 remote/server 档）；
5) 完全退出并重启桌面版（helper 由桌面进程拉起，任务管理器应可见 zcode-cua-helper 进程）；
6) 入口仍灰时先跑 `check-intranet.ps1` 输出各开关状态再逐项对照。

---

## 二、动态工作流（Dynamic Workflow）内网可用性

### 2.1 判定链（源码取证 + 用户内网实测截图交叉验证）

```text
1) 环境变量层：
   主进程（desktop main）构建内核/host 子进程环境时：
     delete m[ZCODE_DYNAMIC_WORKFLOW_MODE]      ← 官方版故意删除该变量
   结论：用户在系统/启动脚本里设置该环境变量无效（被主进程显式删除），"环境变量这条路是死的"。

2) 构建期门禁：
   isPreview 为构建期烙死的字符串 "production"（官方正式版），运行时 env/启动参数无法改变
   代码：isPreview ? "alwaysOn" : { …兜底档位… }
   结论：只有 Preview 内测版安装包默认开启；正式版永远走 else 分支。

3) 内核 host 门禁（真正生效的开关）：
   host 内合并用户档位（云端灰度配置）：
     云端拉取成功 → 用云端档位
     云端不可达（内网） → 兜底硬编码 "disabled"
   结论：内网下动态工作流固定为 disabled。

4) 已知可利用的字节补丁点：
   偏移 280109971 处（app.asar 内）
   "disabled" → "onDemand"（恰 8 字节等长替换，无需改 JSON 目录长度）
   内核在"云端不可达"时按 onDemand 处理（会话注入动态工作流支持）
```

【复核补充 · 本机同代包逐字验证】（squashfs-root/resources/app.asar，340,515,489 字节）

- `resolveDynamicWorkflowModeHostEnv`：`return e.isPreview?{[bo]:"alwaysOn"}:{}` —— 打包正式版（isPreview=false）返回 `{}`，即不向内核注入该变量；
- `buildHostProcessEnv`：`...isPreview:Id});return delete m[bo],{...}` —— 主进程构建内核/host 子进程环境时**显式 delete 该变量**（偏移 282797285 附近逐字命中，与截图结论一致）；
- `Id=!To&&Xe==="preview"`（配合 `qe=To?"development":"production"`）—— isPreview 由构建期常量决定，运行时 env/启动参数不可翻转；
- 内核侧（resources/glm/zcode.cjs）：`Bfr="ZCODE_DYNAMIC_WORKFLOW_MODE",Ufr="disabled"` —— 取值顺序为 env → 远端档位（`remote.mode`）→ 默认 `disabled`；2.1 第 3 点表述精确化：`disabled` 是**本地硬编码默认档**，云端灰度档仅在可达时覆盖，内网下自然落到默认；
- 结论：截图三条结论（env 被删 / isPreview 烙死 / 内网兜底 disabled）在本机同代包**全部复核成立**。

### 2.2 除字节偏移补丁外的其它解决途径（用户问）

| # | 途径 | 可行性 | 说明 |
|---|---|---|---|
| 1 | **环境变量注入**（setx / bat / 启动脚本） | ❌ 已证实不可行 | 主进程构建子进程环境时 `delete m[ZCODE_DYNAMIC_WORKFLOW_MODE]` |
| 2 | **改 isPreview**（production→preview 替换） | ⚠️ 复杂 | 字符串不等长（10 vs 7 字节），需改 JSON 目录长度字段，风险高于 disabled→onDemand |
| 3 | **替换/代理云端灰度接口** | ❌ 大概率不可行 | 灰度接口地址通常硬编码在 host 内，无配置化注入点（已侦察） |
| 4 | **换用 Preview 内测版安装包** | ✅ 最干净的官方途径 | 若官方对外分发 Preview 版，其 isPreview 默认为 true，开箱即用；注意内测版稳定性 |
| 5 | **本地配置/注册表开关** | ⚠️ 待验证 | setting.json 中未见动态工作流显式配置项；如后续版本开放配置项可优先采用（已侦察当前版本无） |
| 6 | **启动参数** | ❌ 未发现 | 3.15.3 未发现 `--dynamic-workflow` 类启动参数 |
| 7 | **字节补丁**（disabled→onDemand） | ✅ 已验证可行 | 等长替换、有备份/校验/回滚脚本支撑（见 patches/），是当前最确定的内网启用方案 |
| 8 | **CLI 会话启用 `--workflow-mode`** | ✅ 零修改，立即可用 | 内核 CLI 原生支持 `--workflow-mode onDemand\|alwaysOn`（须与 `--prompt`/`--target`/`tui` 同用；源码原文报错文案 "Restart with --workflow-mode onDemand or --workflow-mode alwaysOn to use them"）。CLI 进程不经过桌面主进程，**无 delete 逻辑**，env 与参数均被消费。适合脚本/批处理与终端会话；桌面 Composer 输入框内不适用 |
| 9 | **补内核散装文件（推荐的落盘改法）** | ✅ 比 asar 偏移更稳 | 内核 `resources\glm\zcode.cjs`（与 Linux 布局同名的普通文本文件）中默认档位锚点 `Ufr="disabled"` **全文件唯一**（本机验证 count=1，紧邻 `Bfr="ZCODE_DYNAMIC_WORKFLOW_MODE"`），等长替换为 `"onDemand"` 无 asar/JSON 结构风险，回滚即还原该文件 |
| 10 | **预置本地运行偏好（实验性）** | ⚠️ 待现场验证 | 内核会话门禁 `Lai(e,t)` 对运行时偏好做 OR：`appRuntimePreferences.dynamicWorkflowEnabled===true` 即放行。若能在 `%USERPROFILE%\.zcode` 下的偏好/状态 JSON 中预置该键（在设置页同步回写前加载），可不动二进制启用；风险是设置页启动同步可能覆盖，需实测。定位方法：在 `.zcode` 目录全量检索字符串 `dynamicWorkflowEnabled` 的 JSON 文件 |

**修订注（复核对 #3 / #6 的修正）**：
- **#6 原判"❌ 未发现"不成立**：3.15.3 内核存在 `--workflow-mode` 启动参数（即途径 8），仅"桌面 Composer"不适用，CLI 会话完全可用；
- **#3 原判"❌ 大概率不可行"过严**：内核按 `remote.mode`（远端运行时配置）消费云端档位，因此"内网自建 standalone-server / desktop-attached-remote 远端运行时报出 enabled 档位"，或"HTTPS 反代改写灰度接口响应 + 企业 CA 信任（设置 `httpProxyCaCertPath` / env `ZCODE_AGENT_CA_CERT`）"都是**官方通道内的合法注入点**；原结论仅在"单机桌面、无服务端"的限定下成立；
- **推荐组合修订**：①脚本/终端场景 → 途径 8（零修改）；②要桌面 Composer → 途径 9（散装 .cjs 唯一锚点替换，结构风险最低）→ 途径 7（asar 偏移，**必须先定位后打补丁**，见第六节 P0）→ 途径 4（Preview 版，如可获得）。

**推荐组合（原判，保留存档）**：优先尝试途径 4（Preview 版）；无法获取内测版时，用途径 7 字节补丁（配套本包 scripts/patches 的自动化脚本，可一键打补丁并注册计划任务防恢复）。

> ⚠️ 第三轮（2026-10-16 晚）再修订见**第七节**：途径 9 对桌面 Composer 无效/存疑（host 显式覆盖内核默认档）、途径 10 判 ❌、途径 7 语义修正为 host 侧 wa() 兜底补丁，并新增纯配置方案 B（ZCODE_BASE_URL + 内网网关伪造灰度响应）与官方 Hooks 核实（不能解锁）。

---

## 三、全面检查：内网其他不可用/异常功能

（基于 ZCode 源码侦察 + 环境实测，已排除安卓/iOS 模拟器类——用户不使用）

### 3.1 明确不可用（内网受限，需离线方案）

| 功能 | 内网状态 | 离线方案 |
|---|---|---|
| 自动更新检查 | 更新接口不可达，静默失败 | 保持关闭（默认）；版本升级走人工拷贝新安装包 |
| 在线插件市场 | 无法拉取插件清单/安装 | 使用本地 marketplace.json + 离线插件包（如 offline-office-docs） |
| 遥测/统计上报 | 上报失败，不影响功能 | 可接受；必要时配置关闭 |
| 图片搜索 / 视觉搜索 | 依赖云端视觉服务，内网不可用 | 本地图片检索用系统自带/离线工具 |
| 浏览器自动化（Browser Use） | ✅ 本身内网可用（复核实测，原"浏览器下载类"行表述有误，见修订） | 仅"访问外网站点"受限；WebFetch/网页抓取同理——内网 URL 正常，外网 URL 不可达属预期 |
| 登录/OAuth 与订阅管理 | chat.z.ai / z.ai（含 manage-apikey/subscription）不可达；凭据到期后无法在线续期/登录 | 在可联网机器完成登录/续期后按官方凭据导出流程迁移；或改用内网中转供应商的 API Key 模式（见 config/README.md 第 5 节） |
| 飞书/Lark 集成 | 取决于内网是否可达飞书开放平台（内核含 `ZCODE_LARK_CLI_BINARY` 装载逻辑，非硬禁） | 内网有飞书域名则放行即可；否则忽略 |
| CUA（电脑操控） | 入口默认隐藏 + helper/插件开关问题（见第一节 1.4 修订） | 按 1.4 步骤逐项开启 + 本地 UI 自动化替代 |
| 动态工作流 | 云端灰度 disabled（见第二节） | CLI `--workflow-mode` / 散装 .cjs 补丁 / asar 补丁（见 2.2 修订推荐组合） |

### 3.2 正常可用（确认无问题）

- 本地文档处理（docx/xlsx/pptx/pdf，离线插件已交付 offline-office-docs v0.1.0）
- 本地代码/文件读写、终端命令执行
- 自定义模型供应商（中转站 API 可达时可用）
- 任务库/任务历史（tasks-index.sqlite 本地持久化，崩溃可恢复）
- Browser Use 浏览器自动化：playwright-core 随包内置（**无 Chromium 下载动作**，原 3.1"浏览器下载类"表述有误）；桌面由宿主内置浏览器（IAB）承载，CLI 可接本机 Edge/Chrome CDP——面向内网站点的自动化完整可用
- 其余内置插件（documents/presentations/spreadsheets/pdf/skill-creator/plugin-creator/zcode-guide/browser-use/computer-use SDK 等）均为本地资源，不受内网影响

### 3.3 建议

1. 内网机首次部署时，将 `~/.zcode/cli/plugins/cache/` 整个目录从外网环境拷贝同步（含 computer-use 等内置插件）；
2. 关闭自动更新（设置中关闭，或补丁后注册计划任务防自动更新覆盖补丁）；
3. 如需完整离线体验，建议收集官方离线/私有化部署文档（企业版）；
4. 如需经企业代理放行白名单外网端点：设置键 `httpProxy` / `httpProxyNoProxy` / `httpProxyCaCertPath`，env `ZCODE_HTTP_PROXY` / `ZCODE_NO_PROXY` / `ZCODE_AGENT_CA_CERT`（自签 CA 注入）；远端运行时另有 `ZCODE_REMOTE_HTTP_PROXY` / `ZCODE_REMOTE_NO_PROXY` / `ZCODE_REMOTE_RUNTIME_NETWORK_AUTHORITY`；
5. 需在防火墙/代理上按需放行的官方域名清单：`api.z.ai`、`api.chatglm.site`、`bigmodel.cn`、`dev.bigmodel.cn`、`zcode.z.ai`、`chat.z.ai`、`cdn-zcode.z.ai`（插件市场 CDN）、`z.ai`；
6. 遥测如需静默：上报端点可用 `ZCODE_TELEMETRY_REPORT_ENDPOINT__*` 系列 env 覆盖/封堵（上报失败本就不影响功能）；
7. 自动更新配置位于安装目录 `resources/app-update.yml`（provider: generic，url 指向占位地址 http://localhost:8081，updaterCacheDirName `@zcodedesktop-updater`）——内网下更新检查静默失败，无额外风险。

---

## 四、交付物清单

| 产物 | 路径 | 用途 |
|---|---|---|
| 本报告 | REPORT.md | 结论与方案 |
| 使用说明 | README.md | 快速上手 |
| 一键启用脚本（Windows） | scripts/enable-dynamic-workflow.ps1 | 打补丁 + 备份/校验/回滚 + 计划任务 |
| 内网功能体检脚本 | scripts/check-intranet.ps1 | 一键排查 3.1 中各项功能 |
| 字节补丁包 | patches/disabled-to-ondemand.patch / patches/apply-patch.py | 等长替换与校验 |
| 配置建议 | config/ | computerUseComposerEntryHidden 等设置 |
| 日志备份位 | logs-backup/ | 补丁前现场备份 |

---

## 五、证据来源与可信度标注

| 结论 | 可信度 | 来源 |
|---|---|---|
| CUA enabled 判定（platform==darwin） | ✅ 已在本机包内源码核实 | ZCode 3.15.3 app.asar 解析 |
| computerUseComposerEntryHidden=true 默认 | ✅ 已核实 | setting.json + 源码 |
| computer-use 0.6.4 内置插件依赖 | ✅ 已核实 | 插件缓存目录 |
| 动态工作流 env 被 delete | ✅ 已核实 | 主进程 bundle grep + 用户截图 |
| isPreview 构建期烙死 "production" | ✅ 已核实 | 主进程 bundle grep + 用户截图 |
| delete env / isPreview 烙死 / 默认档位 disabled（三连） | ✅ 已在本机同代包逐字核实：`resolveDynamicWorkflowModeHostEnv` 返回 `{}`、`buildHostProcessEnv` 含 `delete m[bo]`、`Id=!To&&Xe==="preview"`、内核 `Ufr="disabled"` | app.asar @282797285 附近 + resources/glm/zcode.cjs |
| 偏移 280109971 | ⚠️ **安装特定值**：本机 340,515,489 字节 app.asar 同偏移处内容为 `allowlis`（toolDisallowlist 域），不可跨机/跨版本硬编码，必须"先定位、后打补丁" | 用户 Windows 实测 + 本机同代包反证 |
| CUA 平台判定（Windows 并非硬禁，走产品 helper 通道） | ✅ 已核实（见 1.4） | app.asar：shouldEnableDefaultCuaProductHelper / wxe / `\\.\pipe\zcode-cua-helper-` |
| 内核 CLI `--workflow-mode` 启动参数 | ✅ 已核实（2.2 途径 6/8 修订） | glm/zcode.cjs：`"workflow-mode":{type:"string"}`、报错文案原文 |

---

## 六、脚本与交付物复核结论（第二双盲审，2026-10-16）

复核方式：在本机同代 Linux 解包（squashfs-root）上对 REPORT 全部可验证结论做独立取证（grep/字节读取），并对 scripts/patches 逐行审读。

### 6.1 脚本复核（对照 16:07 更新版 enable-dynamic-workflow.ps1）

更新版已采纳复核意见的主流：**首选散装 zcode.cjs 锚点方案（幂等：已是 `Ufr="onDemand"` 即跳过）**、偏移标注"安装特定值"、加入 CLI 零修改提示。以下为**仍然存在**的问题：

- **P0（仅 asar 备选分支）**：L134-138 定点偏移校验失败时仍自动降级为"全文件第一个 `disabled` 盲替换"。同代内核文件中 `disabled` 字面量 160 处、asar 内更多，且 asar 分支在**重跑时**（偏移处已是 onDemand）必然触发该降级并改坏无关字符串；同时 L130-132 是"先备份后校验"，重跑产生的备份是已补丁状态，`-Rollback` 取最新备份等于恢复成已补丁版本。**修复要求**：锚点找不到就报错退出，禁止任何"第一个匹配"盲替换；asar 分支改为"先校验偏移、后备份、幂等跳过"。临时规避：不使用 `-UseAsarOnly`；确保 `resources\glm\zcode.cjs` 存在（首选路径幂等安全）。
- **P1（计划任务边界）**：登录自愈任务对 .cjs 路径已安全（幂等跳过），但若历史补丁是经 asar 分支打上的，自愈任务同样会触发 P0。修复 P0 后此条自动解除。

### 6.2 check-intranet.ps1 / apply-patch.py 复核（P1）

1. L105 `IndexOf("disabled", 280000000)` 命中的是 280000000 之后**第一个** disabled，不一定是门禁常量——检出的偏移可能误导 asar 打补丁；应精确偏移探测 + 前后上下文校验（首选 .cjs 锚点检测已加，此条影响有限）。
2. L27-29 的 `https://marketplace.zcode.local` / `https://config.zcode.local` 为虚构地址，探测无意义；应探测真实域名（见 3.3 第 5 条清单，如 `https://cdn-zcode.z.ai/zcode/official-plugin/marketplace.json`）。另外 `Invoke-WebRequest -Method Head` 在目标返回 404/405 时也走异常分支，"受限"判读需注明"不可达或拒绝 HEAD"。
3. 340MB 级 `ReadAllBytes` + `GetString` 全量解码内存开销大（约需 2-3 倍文件大小的瞬时内存），**必须在 x64 PowerShell 下运行**（32 位会 OOM）。
4. apply-patch.py（未随 16:07 更新）L47-52 的全文件降级仍是"第一个匹配"盲替换，需加与 6.1 相同的锚点定位；其 `.bak` 仅在不存在时创建的原始语义正确。

### 6.3 更稳的等长替换目标（16:07 更新版脚本已内置此方案，以下留作参考与独立手工操作）

内核散装文件 `resources\glm\zcode.cjs`（Windows 安装布局同名路径）是普通文本文件，不涉及 asar JSON 目录结构：
- 锚点 `Ufr="disabled"` 全文件唯一（同代包验证 count=1，紧邻 `Bfr="ZCODE_DYNAMIC_WORKFLOW_MODE"`）；
- 等长替换 `"onDemand"` 后无任何长度/索引副作用，回滚=还原该文件；
- 参考实现（幂等、锚点定位）：

```powershell
$path = "C:\...\resources\glm\zcode.cjs"
$bytes = [System.IO.File]::ReadAllBytes($path)
$ascii = [System.Text.Encoding]::ASCII
$anchor = $ascii.GetBytes('Ufr="disabled"'); $to = $ascii.GetBytes('Ufr="onDemand"')
function Find-Bytes([byte[]]$hay, [byte[]]$needle, [int]$start) {
  for ($i=$start; $i -le $hay.Length-$needle.Length; $i++) {
    $ok=$true; for ($j=0; $j -lt $needle.Length; $j++){ if ($hay[$i+$j] -ne $needle[$j]){ $ok=$false; break } }
    if ($ok) { return $i }
  }; return -1
}
$a = Find-Bytes $bytes $anchor 0
if ($a -lt 0) { throw '锚点 Ufr="disabled" 未找到（可能已补丁或版本不同）' }
$hit = Find-Bytes $bytes $to 0
if ($hit -ge 0) { Write-Host '已是 onDemand，幂等退出'; exit 0 }
for ($j=0; $j -lt $to.Length; $j++){ $bytes[$a+$j] = $to[$j] }
[System.IO.File]::WriteAllBytes($path, $bytes)
Write-Host "补丁完成 @ $a（锚点 Ufr=\"disabled\"，等长替换）"
```

（注：asar 内替换仍须以"先定位、后替换"为前提，偏移 280109971 仅在用户实测那台机器的那次安装上有效。）

### 6.4 复核结论汇总

| 项 | 复核结果 |
|---|---|
| REPORT 2.1 三条截图结论 | ✅ 全部在本机同代包逐字坐实 |
| REPORT 1.1/1.2 CUA "Windows 平台硬禁" | ⚠️ 修订：Windows 走产品 helper 通道（见 1.4） |
| REPORT 2.2 途径 6（启动参数） | ❌→✅ 修订：`--workflow-mode` 存在（途径 8） |
| REPORT 2.2 途径 3（云端接口） | ⚠️ 收窄：服务端形态/反代改写为官方通道注入点 |
| REPORT 第三节清单 | ✅ 已补登录/OAuth、飞书、WebFetch、Browser Use 修正、域名清单、代理/CA 键 |
| scripts/enable-dynamic-workflow.ps1（16:07 更新版） | ✅ 首选 .cjs 锚点+幂等已到位；⚠️ asar 备选分支仍有 1 个 P0 盲替换（6.1），规避：不用 `-UseAsarOnly`、确保 .cjs 存在 |
| scripts/check-intranet.ps1（16:07 更新版） | ⚠️ 3 个 P1（6.2），检出结果仅供参考 |
| patches/apply-patch.py | ⚠️ 盲替换降级同需锚点（6.2#4），.bak 语义正确 |
| 交付物 zip | 已同步刷新（dist/zcode-intranet-fixes-plugin-0.1.0.zip） |

### 6.5 第三轮复核（2026-10-16 晚）：P0/P1 修复落地与回归结论

本轮接手 6.1/6.2 意见的收尾修复，并对三个工具逐行复核。**所有 P0/P1 已关闭**：

| 项 | 6.1/6.2 原判 | 本轮处置 | 现状 |
|---|---|---|---|
| apply-patch.py 全文件盲替换降级（6.2#4，P0 级） | 需锚点定位 | 用户已改：定点校验失败报错退出（exit 1），无盲替换路径 | ✅ 已修复；实测 verify/apply/幂等/rollback/FAIL 五路径全通过（含真实同代 asar 上 FAIL 路径返回 `allowlis` 并退出 1） |
| enable-dynamic-workflow.ps1 asar 分支盲替换（6.1 P0） | 锚点找不到就报错退出；先校验后备份；幂等 | 用户已改基础流程；本轮重写固化：`Invoke-OffsetReplace`（FileStream seek 校验→写→复验，任意大小文件不再整读），asar 偏移处内容非 disabled/onDemand 即抛错退出 | ✅ 已修复 |
| **enable-dynamic-workflow.ps1 新发现 P0**（本轮引入审查时发现） | — | 首选 .cjs 分支调用 `Invoke-LiteralReplace` 未传 `-Offset`（默认 0），导致**首选分支必然在校验处抛错**并降级 asar | ✅ 已修复：重写为 `Find-AsciiOffsets` 锚点定位（要求 `Ufr="disabled"` 全文件唯一，count=1 才替换；count=0/>1 均不动手转 asar） |
| check-intranet.ps1 asar 检测 IndexOf 全文件扫描（6.2#1，P1） | 误导 | 改为 FileStream seek 精确读偏移 280109971 处 8 字节并报告实际内容（disabled/onDemand/其它原样显示），不再解码 340MB | ✅ 已修复 |
| check-intranet.ps1 x64 内存警告（6.2#3，P1） | 必须 x64 | 脚本头加 32 位进程警告（asar 已无需全量解码，.cjs 分支整读约 15MB×2 仍建议 x64） | ✅ 已修复 |
| check-intranet.ps1 虚构域名（6.2#2，P1） | 改真实域名 | 用户已改（cdn-zcode.z.ai 等），HEAD 判读注明"不可达或拒绝 HEAD" | ✅ 已修复 |
| 顺手修正的旧 bug（本轮复核发现） | — | ①`Write-Check "自动更新" $upd.AutoUpdate` 缺第三参数（Detail 恒空、状态色误判）→ 补全并按值判 OK/受限；②`$results +=` 在函数内闭包失效（汇总恒空）→ 改 `$script:results`（List）并在结尾输出汇总表与统计；③回滚只恢复"最新一份"备份（.cjs 与 asar 混补时漏恢复）→ 改为按文件种类分别恢复最新备份；④两路分支均未生效时仍提示"完成" → 加 `$dwEnabled` 状态标记，未生效时输出告警与替代方案 | ✅ 已修复 |

回归说明：
- `Find-AsciiOffsets` 用 **ASCII**（非 UTF8）解码定位——UTF8 解码非法序列会产生替换字符导致偏移漂移，纯 ASCII 锚点必须 1:1 解码；
- asar 分支全程 FileStream seek（读 8 字节校验 / 写 8 字节 / 读 8 字节复验），**内存占用 O(1)**，32 位 PowerShell 也可跑（仍建议 x64）；
- 计划任务自愈对两条路径均幂等（已是 onDemand 即跳过），6.1 的"计划任务触发 P0"边界随之解除；
- PowerShell 本机无运行环境（Linux），上述脚本经逐行静态复核 + 与 apply-patch.py 同构算法实测兜底；首次在目标机运行时建议先跑 `-SkipPlanTask` 观察输出。



## 七、内网绕过动态工作流禁用：完整方案矩阵（第三轮取证 + 官方文档核实，2026-10-16）

依据：本机同代 Linux 解包（squashfs-root，app.asar 340,515,489 字节）逐字取证 `out/main/index.js`（主进程）、`out/host/index.js`（宿主进程 bundle，asar 内起始偏移 279776654 / 1,681,899 字节）、`glm/zcode.cjs`（内核 15,127,039 字节），以及官方 Hooks 文档（zcode.z.ai/en/docs/hooks）与 QCode 中转文档（docs.qcode.cc）核实。

### 7.0 门禁链全图（本轮新证据，对第二节的再精确化）

```text
【桌面链路】（每一步均已逐字核实）
灰度数据源（out/host/index.js 内 BigModelCodingPlanSubscriptionProvider）：
  getDynamicWorkflowClientConfig():
    ① env 短路：if (normalizeDynamicWorkflowMode(process.env["ZCODE_DYNAMIC_WORKFLOW_MODE"]))
         → 直接以 env 值作为 offer（source=override），不发网络请求
       …… 但主进程 buildHostProcessEnv 构建 host 子进程 env 时 delete 该键（Vx 打包正式版返回 {}），
          故桌面下该短路不可达（与 2.1 实测一致，env 注入仍是死的）
    ② GET {endpoint}/api/v1/client/configs?app_version=…&platform=…
       → 消费响应 JSON 的 data.configs.dynamicWorkflow（.mode 字段）
       → 15 秒超时 / 1 小时缓存 / 失败日志"动态工作流灰度配置读取失败，按关闭处理"→ disabled 兜底
    endpoint 解析：ZCODE_ENV 烙死 "production"
       → ZCODE_BASE_URL ?? ZCODE_ENDPOINT_ORIGIN ?? ZCODE_PRODUCTION_BASE_URL
       （默认 https://zcode.z.ai；test 档为 https://zcode.chatglm.site）
host 侧门禁（out/host/index.js）：
  wa() resolveDynamicWorkflowGate：let g = await Gi(); return g ? gg(g, await ko()).mode : "disabled"
       ↑ 灰度 offer 读不到 → 兜底 "disabled"
  va() enqueueDynamicWorkflowPolicySync：wa() 结果 ≠ 当前 policy 状态时，向内核发
       workspace/updateDynamicWorkflowPolicy { workspace, enabled, mode }
  As() resolveDynamicWorkflowSessionFlags → createSession envelope 注入 dynamicWorkflowEnabled/Mode
内核侧（glm/zcode.cjs）：
  BZi resolveDynamicWorkflowOffer：env["ZCODE_DYNAMIC_WORKFLOW_MODE"](override)
       → remote.mode(remote) → Ufr="disabled"（default；count=1 本轮再次验证）
  Nai updateDynamicWorkflowPolicy（收到 host 请求）：写 appRuntimePreferences.{dynamicWorkflowEnabled,Mode}
  Lai resolveDynamicWorkflowRuntimeGate：请求带 dynamicWorkflowEnabled===true 或
       appRuntimePreferences.dynamicWorkflowEnabled===true → 启用（mode 缺省 alwaysOn）
       否则【显式】返回 dynamicWorkflowEnabled:false → 会话 config 显式 false
  工具注册门禁：includeDynamicWorkflow===false 时把 CreateWorkflow 工具族整组排除（dRo 集合）
TLS：
  内核 TLS 连接为标准校验（rejectUnauthorized!==false 且未授权 → 断开），无证书固定（pinning）；
  ZCODE_AGENT_CA_CERT（内核 ZCODE_AGENT_CA_CERT_ENV_KEY）可指定 CA 文件（readFileSync 加载）；
  NODE_EXTRA_CA_CERTS / SSL_CERT_FILE / REQUESTS_CA_BUNDLE / GIT_SSL_CAINFO 等在子进程透传白名单；
  桌面另有设置键 httpProxyCaCertPath → 企业 CA 注入是官方支持面
userMode 的覆盖限制（重要）：
  applyDynamicWorkflowUserMode（host 与内核同构实现）仅在 offer 非 disabled 时才应用用户本地档位
  appRuntimePreferences 为纯内存态（host 启动即 false），无盘上持久化点，仅能由
  workspace/updateDynamicWorkflowPolicy 请求翻转，而该请求由 host 门禁 wa() 驱动
```

三个直接推论（对 2.2 的修订）：
1. **偏移 280109971 的真实身份**（见 7.2）：用户实测补丁点位于 **host 侧 `wa()` 兜底 `"disabled"`**——之前"内核在云端不可达时按 onDemand 处理"的表述不精确，应为"host 在灰度读取失败/未命中时按补丁值处理"。
2. **途径9（内核 .cjs 的 Ufr 档位）对桌面 Composer 基本无效**（见 7.3 方案E）：桌面链路每次会话由 host 显式下发 `dynamicWorkflowEnabled:false`（Lai 强制返回、envelope 覆盖），内核默认档只对"无 host、无 remote"的直连内核场景生效；`Ufr` 补丁的实际价值收窄为 CLI 兜底，而 CLI 本就有零修改的方案A。
3. **途径10（预置本地运行偏好）判 ❌**：userMode 仅在 offer 非 disabled 时生效（`!o||!zfr(n)?n:…`），且 appRuntimePreferences 是内存态、内网下永远翻不成 true。

### 7.1 官方 Hooks 能否影响动态工作流门禁 —— 不能 ❌

- 官方文档（zcode.z.ai/en/docs/hooks）事件集：SessionStart / UserPromptSubmit / PreToolUse / PermissionRequest / PostToolUse / PostToolUseFailure / Stop；能力面为：additionalContext 上下文注入、updatedInput 工具输入替换（schema 复验）、permissionDecision（allow/ask/deny）、UserPromptSubmit 阻断、Stop `decision:block`（最多连续 3 次）。
- 档位判定发生在**会话创建之前**（7.0 全图：wa()/As() → updateDynamicWorkflowPolicy → appRuntimePreferences → Lai → 工具注册 BDa）；Hooks 全部运行在会话内工具环，文档能力清单中没有任何"改会话配置/特性开关/工具注册"项，且明示"显式 deny、Plan 模式写禁、硬性工具限制不可被 hook allow 绕过"。
- 源码旁证：CreateWorkflow 工具族在**注册期**即被排除（dRo），会话内不存在可被 PreToolUse 放行的 workflow 工具；Hooks 也无法触达 `workspace/updateDynamicWorkflowPolicy`（该请求由 host 内部门禁驱动，非工具调用）。
- 结论：官方 Hooks 体系**不能**注入 dynamicWorkflowEnabled 偏好、不能改内核档位、不能解锁动态工作流。

### 7.2 偏移 280109971 的定位方法（跨机/跨版本重定位标准流程）

- 身份：host bundle 内 `wa()`（resolveDynamicWorkflowGate）的兜底字符串。本机同代 Linux 包中该函数位于 asar 偏移 280108693（`out/host/index.js` 相对 332039），原文：
  `async function wa(){let g=await Gi();return g?gg(g,await ko()).mode:"disabled"}`；
  紧邻其后是 `enqueueDynamicWorkflowPolicySync` 与 `workspaceUpdateDynamicWorkflowPolicy` 字面量（本机 280108945）——与用户 Windows 包实测偏移 280109971 的布局同构（偏移值本身安装特定）。
- ⚠️ 同区域有**两个以上** `disabled` 字面量：除 wa() 兜底外，还有两处 carrier 初始值 `dynamicWorkflowPolicyMode:"disabled"`。**补错位置会把初始状态改成 onDemand 而 wa() 仍返回 disabled，导致每次会话下发 enabled:false（反向禁用）**。
- 标准定位流程（等长 8 字节替换 `"disabled"`→`"onDemand"`）：
  1. 在 app.asar 全文搜特征串 `).mode:"disabled"`（本机 host bundle 内 count=1，唯一）；
  2. 确认其上下文为 `await ko()).mode:"disabled"` 且其后不远处有 `i(wa,"resolveDynamicWorkflowGate")` 标注（minified 保留函数名，可作二次确认）；
  3. 记录该偏移 → seek 校验 8 字节原文为 `disabled` → 备份 → 定点替换 → 复验。
  check-intranet.ps1 与 enable-dynamic-workflow.ps1 的 asar 分支已内置"先校验、不符即拒"逻辑；跨机时将脚本顶部 `$asarOffset` 改为定位结果（apply-patch.py 同理改 `OFFSET`）。

### 7.3 方案矩阵（★=不需要改二进制的纯配置方案）

| # | 方案 | 原理 | 可行性 | 稳定性 | 自动更新覆盖风险 | 部署步骤 |
|---|---|---|---|---|---|---|
| A ★ | **CLI `--workflow-mode`（或 CLI 进程 env）** | 内核 CLI 不经桌面主进程，无 delete 逻辑；env→参数→override 档位 | ✅ 源码+官方报错文案双证 | ✅ 高 | 无（命令行用法） | `zcode --workflow-mode onDemand\|alwaysOn`（与 `--prompt`/`--target`/`tui` 同用）；直接跑 zcode.cjs 时也可预设 `ZCODE_DYNAMIC_WORKFLOW_MODE=alwaysOn`（CLI 进程 env 不被删）。局限：桌面 Composer 输入框不适用 |
| B ★★ | **`ZCODE_BASE_URL` + 企业 CA + 内网网关伪造灰度响应（重点推荐，纯配置解锁桌面 Composer）** | host 的灰度读取 `GET {ZCODE_BASE_URL}/api/v1/client/configs`，消费 `data.configs.dynamicWorkflow.mode`；`ZCODE_BASE_URL` 是用户级 env 且**不在 buildHostProcessEnv 删除名单**（仅删 6 个 ZCODE_TELEMETRY_* 与 ZCODE_DYNAMIC_WORKFLOW_MODE），`Uh`(applySelectedZCodeEnvLinks) 还会显式把该键传入 host env；TLS 为标准校验、可经 `ZCODE_AGENT_CA_CERT`/`NODE_EXTRA_CA_CERTS`/`httpProxyCaCertPath` 信任企业 CA | ✅ 静态取证充分（接口路径/响应字段/env 传递链均已核实）；⚠️ 需内网机实测确认端到端 | ⚠️→✅ 版本升级若改响应结构会失效，且有日志"灰度读取失败按关闭处理"可观测；env 与网关不受更新覆盖 | 无 | ① 用户级 setx `ZCODE_BASE_URL=https://zcode-config.intranet.corp`；② 内网网关/反代对该路径返回 `{"code":0,"success":true,"data":{"configs":{"dynamicWorkflow":{"mode":"alwaysOn"}}}}`（mode 亦可 onDemand；其余路径回源或 404）；③ 私网 HTTPS 用企业 CA 签发并在机器上信任。注意：该变量会同时重定向其它 zcode API 调用（登录/订阅等，内网本不可达，影响可接受；如需精细控制可在网关按路径分流） |
| C ★ | **官方 remote 运行时下发 remote.mode（standalone-server / desktop-attached-remote）** | 内核 BZi 第二优先级 `e.remote.mode`，即"云端档位"的官方通道；企业自建远端运行时报出 enabled 档位即等价云端灰度命中 | ✅ 通道真实存在 | ✅ 高 | 无 | 部署企业级远端运行时并下发运行时配置（含 mode）；成本高，适合已有企业平台的内网，个人内网机不推荐 |
| D | **asar 偏移补丁（host 侧 wa() 兜底）**——桌面已验证 | 见 7.2；wa() 返回 onDemand → 首次 policy sync 置内核偏好 → 全部桌面会话注入启用 | ✅ 用户内网实测 | ⚠️ 跨版本/跨机失效，须按 7.2 重新定位 | ❌ 会被更新覆盖（计划任务自愈兜底） | `enable-dynamic-workflow.ps1`（asar 备选分支自动介入，或 `-UseAsarOnly`；改 `$asarOffset` 适配本机） |
| E | **散装 zcode.cjs `Ufr="disabled"`→`"onDemand"`（内核默认档）** | 内核 BZi 第三优先级默认档 | ⚠️ 本轮**下调**：对 CLI 直连有效但与方案A等价；对桌面 Composer 静态证据表明**无效**（host 每会话显式下发 false，覆盖内核默认档；见 7.0 推论2） | ✅ 锚点唯一、等长安全 | ❌ 更新覆盖（计划任务自愈） | 脚本首选分支保留（幂等、无害、CLI 兜底）；**桌面 Composer 请以方案 B/D 为准**。若现场实测 E 对桌面有效（存在未发现的回退路径），欢迎回填本报告 |
| F | Preview 内测版安装包 | isPreview 构建期 alwaysOn | ✅ 如可获得仍是最干净路径 | ⚠️ 内测版稳定性 | — | 原判保留（2.2 途径4） |
| ✗ | 环境变量注入（桌面链路） | — | ❌ 已证死（delete + 实测） | — | — | — |
| ✗ | 官方 Hooks 解锁 | — | ❌（7.1） | — | — | — |
| ✗ | 预置本地运行偏好（原途径10） | — | ❌（7.0 推论3） | — | — | — |
| ✗ | 启动参数（桌面）/ settings 配置项 | — | ❌ 当前版本未发现 | — | — | — |

**推荐组合（第三轮修订）**：脚本/终端 → 方案A（零修改）；桌面 Composer → 首选方案B（纯配置，不怕升级）→ 备选方案D（补丁+计划任务自愈）→ 企业场景方案C；方案E 仅作 CLI 兜底保留。

### 7.4 Claude Code workflow.js 兼容性核实

- ZCode 内核**原生实现本地脚本编排**（非云端执行）：`.workflow.js` 落盘于 `cli/sessions/<sessionId>/workflows/`（persistInlineScript / resolveNamedWorkflowPath / createScriptWorkflowToolPort），原语同族——agent / parallel / pipeline / phase + hole 填洞；TS 编译、子代理编排、结果回传全部本地，模型调用走已配置供应商。**工作流运行期不依赖灰度/云端档位接口**。
- 但与 Claude Code 的 workflow 文件**不是同一格式**（facade 不同：hole/fill 语义、nodeRepl 桥、类型面、Amend/Resume 协议等），文件不可直接互换；"本地有 workflow 文件"本身**不绕过门禁**——CreateWorkflow 工具族在会话工具注册期即被排除（7.0），必须先让门禁开启（方案A/B/C/D 之一）。
- QCode 文档（docs.qcode.cc）核实：其"works out of the box"指 **Claude Code 客户端自带动态工作流、与模型供应商无关**（仅设 `ANTHROPIC_BASE_URL`/`ANTHROPIC_AUTH_TOKEN` 指路；另提示其每 key 5 并发与工作流默认 16 并发的 429 冲突，需下调 `CLAUDE_CODE_WORKFLOW_MAX_CONCURRENT_AGENTS`）。**该文档不构成"remote.mode 云端下发"的证据**——它印证的结论是：动态工作流是客户端本地编排能力，档位是 ZCode 自己的产品门禁（与本机取证一致）；官方云端下发通道的本体是方案C。

### 7.5 与 2.2 途径表的对账（摘要）

| 2.2 途径 | 第三轮结论 |
|---|---|
| 1 env 注入 / 6 启动参数（桌面） | ❌ 维持 |
| 2 isPreview 翻转 | ⚠️ 维持（不等长、结构风险，收益不及 B/D） |
| 3 云端灰度接口 | ✅ 细化升级为方案 B（ZCODE_BASE_URL 可配化 + 响应结构已核实）与方案 C（remote 运行时） |
| 4 Preview 版 | ✅ 维持（方案F） |
| 5 settings 配置项 | ❌ 维持（当前版本无） |
| 7 asar 偏移 | ✅ 语义修正为 host 侧 wa() 兜底补丁（方案D，7.2） |
| 8 CLI --workflow-mode | ✅ 方案A |
| 9 散装 .cjs | ⚠️ 下调为方案E（CLI 兜底；桌面 Composer 无效/存疑） |
| 10 预置本地偏好 | ❌ 下调（7.0 推论3） |
| 新增 | 官方 Hooks 解锁 ❌（7.1） |

## 八、最终建议（一句话版，第三轮更新）

脚本/终端要动态工作流：直接 `zcode --workflow-mode onDemand|alwaysOn`，零修改（方案A）。桌面 Composer 要动态工作流：优先验证纯配置方案 B（用户级 `ZCODE_BASE_URL` 指向内网网关 + 企业 CA + 网关伪造 `/api/v1/client/configs` 的 `data.configs.dynamicWorkflow.mode`，不改任何二进制、不怕升级）；退而求其次用 asar 偏移补丁（方案D，先按 7.2 用唯一特征串 `).mode:"disabled"` 定位再等长替换）；散装 .cjs 的 `Ufr` 补丁只对 CLI 兜底有效、对桌面 Composer 无效（方案E，降级保留）。官方 Hooks 不能解锁该门禁（7.1）；CUA 在 Windows 按 1.4 步骤开启（平台并非硬禁）。
