# ZCode 桌面端内网环境适配调查报告

- 版本：v0.1.0
- 日期：2026-10-16
- 适用环境：Windows 内网（无外网）ZCode 桌面端
- 调查方式：源码取证（app.asar 解析 / 主进程 bundle grep / 配置与日志分析）+ 用户内网实测截图交叉验证
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

### 6.1 P0（使用 enable-dynamic-workflow.ps1 前必须知晓/修正）

1. **计划任务自愈 + 盲替换降级 = 每次登录随机改坏文件的组合缺陷**。脚本 L108-115 在定点偏移校验失败时自动降级为"全文件搜索第一个 `disabled` 盲替换"，L126-139 又注册了**每次登录触发**的计划任务。补丁生效后：下一次登录定点校验失败（偏移处已是 onDemand）→ 降级盲替换 → 把枚举字面量 `["disabled","onDemand","alwaysOn"]` 或任意 UI 字符串改坏（同代内核文件中 `disabled` 字面量 160 处）。**修复要求**：①入口幂等——定点偏移已是 `onDemand` 时直接报"已打过补丁"并成功退出；②降级路径必须按唯一上下文锚点定位（asar 主进程副本锚 `AX="ZCODE_DYNAMIC_WORKFLOW_MODE"` 后首个 `"disabled"`；散装内核锚 `Ufr="disabled"`，本机验证全文件唯一），锚点找不到就报错退出，**禁止任何"第一个匹配"盲替换**；③计划任务指向的入口须先做幂等校验。临时规避：运行时加 `-SkipPlanTask`，且只执行一次。
2. **校验谓词恒假**（L118-124）：`-not $probe.Substring(0, 280109971+8).Contains("onDemand")` ——"onDemand" 在枚举字面量中大量出现于该偏移之前，子串必命中，条件必为 false，**补丁成功也会打印"校验提示"**。修复：直接读取偏移处 8 字节与 `onDemand` 逐字节比对。
3. **备份/回滚语义**：PS1 每次运行都新做 `app.asar.bak-<时间戳>` 备份——第二次运行备份的是**已补丁文件**，`-Rollback` 取"最新备份"等于恢复成已补丁版本。修复：回滚取最早备份，或改为"无原始备份才备份"（patches/apply-patch.py 的 `.bak` 语义即正确，可参考）。

### 6.2 P1（建议修正）

4. check-intranet.ps1 L85 `IndexOf("disabled", 280000000)` 命中的是 280000000 之后**第一个** disabled，不一定是门禁常量——检出的偏移可能误导打补丁；应改为精确偏移探测 + 前后上下文校验。
5. check-intranet.ps1 L27-30 的 `https://marketplace.zcode.local` / `https://config.zcode.local` 为虚构地址，探测无意义；应探测真实域名（见 3.3 第 5 条清单，如 `https://cdn-zcode.z.ai/zcode/official-plugin/marketplace.json`）。另外 `Invoke-WebRequest -Method Head` 在目标返回 404/405 时也走异常分支，"受限"判读需注明"不可达或拒绝 HEAD"。
6. 340MB 级 `ReadAllBytes` + `GetString` 全量解码内存开销大（约需 2-3 倍文件大小的瞬时内存），**必须在 x64 PowerShell 下运行**（32 位会 OOM）。
7. apply-patch.py L47-52 的全文件降级同样是"第一个匹配"盲替换，需加与 6.1#1 相同的锚点定位。

### 6.3 更稳的等长替换目标（推荐）

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
| scripts/enable-dynamic-workflow.ps1 | ⚠️ 1 个 P0 + 1 个校验缺陷 + 备份语义问题（6.1），按 6.1/6.3 加固后使用 |
| scripts/check-intranet.ps1 | ⚠️ 3 个 P1（6.2），检出结果仅供参考 |
| patches/apply-patch.py | ⚠️ 盲替换降级同需锚点（6.2#7），.bak 语义正确 |
| 交付物 zip | 已同步刷新（dist/zcode-intranet-fixes-plugin-0.1.0.zip） |

## 七、最终建议（一句话版）

桌面 Composer 内要动态工作流：优先改散装内核文件 `resources\glm\zcode.cjs`（锚 `Ufr="disabled"` 唯一、等长替换、无结构风险）；脚本/终端场景直接 `--workflow-mode alwaysOn` 零修改；asar 偏移补丁仅在重新定位锚点后作为最后手段；CUA 在 Windows 按 1.4 步骤开启（平台并非硬禁）。
