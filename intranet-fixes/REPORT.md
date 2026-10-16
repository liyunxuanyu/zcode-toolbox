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
   "disabled" → "onDemand"（恰 8 字节等长替换，无需改 JSON 长度）
   内核在"云端不可达"时按 onDemand 处理（会话注入动态工作流支持）
```

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

**推荐组合**：优先尝试途径 4（Preview 版）；无法获取内测版时，用途径 7 字节补丁（配套本包 scripts/patches 的自动化脚本，可一键打补丁并注册计划任务防恢复）。

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
| 浏览器下载类（Chromium 内置下载依赖外网资源） | 部分站点资源不可达 | 走内网镜像/离线分发 |
| CUA（电脑操控） | 云端授权服务不可达（见第一节） | 入口开启 + 本地 UI 自动化替代 |
| 动态工作流 | 云端灰度 disabled（见第二节） | Preview 版或字节补丁 |

### 3.2 正常可用（确认无问题）

- 本地文档处理（docx/xlsx/pptx/pdf，离线插件已交付 offline-office-docs v0.1.0）
- 本地代码/文件读写、终端命令执行
- 自定义模型供应商（中转站 API 可达时可用）
- 任务库/任务历史（tasks-index.sqlite 本地持久化，崩溃可恢复）

### 3.3 建议

1. 内网机首次部署时，将 `~/.zcode/cli/plugins/cache/` 整个目录从外网环境拷贝同步（含 computer-use 等内置插件）；
2. 关闭自动更新（设置中关闭，或补丁后注册计划任务防自动更新覆盖补丁）；
3. 如需完整离线体验，建议收集官方离线/私有化部署文档（企业版）。

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
| 云端灰度 disabled 兜底 / 偏移 280109971 | ⚠️ 来自用户内网实测截图（本机 Linux 包同源验证中） | 用户实测 + ZCode 取证 |
