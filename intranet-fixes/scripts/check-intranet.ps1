#requires -Version 5.1
<#
.SYNOPSIS
    内网环境 ZCode 功能体检脚本：一键排查哪些功能不可用/受限，并给出离线方案。
.NOTES
    纯 PowerShell 5.1，无需外网。只读检查，不改任何配置。
    请在 64 位 PowerShell 下运行（脚本内大文件字节读取在 32 位进程下可能 OOM，
    详见 REPORT.md 6.2 第 3 条；asar 探测已改为按需 seek 读 8 字节，不再全量解码）。
#>

$ErrorActionPreference = "Continue"

if (-not [Environment]::Is64BitProcess) {
    Write-Warning "当前是 32 位 PowerShell。本脚本涉及大文件读取（zcode.cjs 约 15MB 整读；历史版本曾对 app.asar 约 340MB 全量解码，32 位下必 OOM）。建议用 C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe（64 位）重跑。"
}

Write-Host "===== ZCode 内网功能体检 =====" -ForegroundColor Cyan
Write-Host ""

$zcodeHome = "$env:USERPROFILE\.zcode"
$script:results = New-Object System.Collections.Generic.List[object]

function Write-Check {
    param([string]$Name, [string]$Status, [string]$Detail)
    $script:results.Add([PSCustomObject]@{ 功能=$Name; 状态=$Status; 说明=$Detail }) | Out-Null
    $color = if ($Status -eq "OK") { "Green" } elseif ($Status -eq "受限") { "Yellow" } else { "Red" }
    Write-Host ("[{0}] {1}: {2}" -f $Status, $Name, $Detail) -ForegroundColor $color
}

# ---- 1. 网络可达性（真实官方域名，见 REPORT 3.3 第 5 条清单）----
Write-Host "--- 1. 外网/云端可达性 ---"
$hosts = @(
    @{ N="云端授权服务(CUA)"; U="https://z.ai" },
    @{ N="插件市场CDN"; U="https://cdn-zcode.z.ai" },
    @{ N="模型API(中转/官方)"; U="https://api.chatglm.site" }
)
$netOK = $false
foreach ($h in $hosts) {
    try {
        $r = Invoke-WebRequest -Uri $h.U -Method Head -TimeoutSec 4 -UseBasicParsing -ErrorAction Stop
        Write-Check $h.N "OK" "可达 (HTTP $($r.StatusCode))"
        $netOK = $true
    } catch {
        Write-Check $h.N "受限" "内网不可达或拒绝 HEAD（预期内）"
    }
}
if (-not $netOK) { Write-Host "  注：全不可达符合内网预期，以下功能按离线模式评估。" -ForegroundColor DarkGray }

# ---- 2. CUA 入口与插件 ----
Write-Host "--- 2. Computer Use (CUA) ---"
$setting = Join-Path $zcodeHome "v2\setting.json"
$cuaHidden = $null
if (Test-Path $setting) {
    try {
        $j = Get-Content $setting -Raw | ConvertFrom-Json
        $cuaHidden = $j.computerUseComposerEntryHidden
    } catch { }
}
if ($null -eq $cuaHidden) {
    Write-Check "CUA 入口" "受限" "computerUseComposerEntryHidden 未显式配置（默认为 true=隐藏）"
} elseif ($cuaHidden) {
    Write-Check "CUA 入口" "受限" "computerUseComposerEntryHidden=true（入口隐藏）"
} else {
    Write-Check "CUA 入口" "OK" "入口已开启（false）"
}
$cuaPlugin = Join-Path $zcodeHome "cli\plugins\cache\zcode-plugins-official\computer-use"
if (Test-Path $cuaPlugin) {
    $ver = (Get-ChildItem $cuaPlugin -Directory | Select-Object -First 1).Name
    Write-Check "CUA 内置插件" "OK" "computer-use 插件在位 ($ver)"
} else {
    Write-Check "CUA 内置插件" "缺失" "未找到 computer-use 插件缓存，需从外网环境拷贝"
}

# ---- 3. 动态工作流 ----
Write-Host "--- 3. 动态工作流 ---"
$envV = $env:ZCODE_DYNAMIC_WORKFLOW_MODE
if ($envV) {
    Write-Check "env 注入" "受限" "检测到 ZCODE_DYNAMIC_WORKFLOW_MODE=$envV，但官方版主进程会 delete 该变量，实际无效"
} else {
    Write-Check "env 注入" "受限" "未设置 ZCODE_DYNAMIC_WORKFLOW_MODE（即便设置官方版也会删除）"
}
# 3.1 首选锚点：散装 zcode.cjs
$cjs = $null
$cjsCandidates = @(
    "$env:USERPROFILE\.zcode\app\resources\glm\zcode.cjs",
    "$env:LOCALAPPDATA\Programs\zcode\resources\glm\zcode.cjs"
)
foreach ($c in $cjsCandidates) { if (Test-Path $c) { $cjs = $c; break } }
if ($cjs) {
    $txt = [System.IO.File]::ReadAllText($cjs)
    if ($txt.Contains('Ufr="onDemand"')) {
        Write-Check "zcode.cjs 档位" "OK" "已是 onDemand（动态工作流已启用）"
    } elseif ($txt.Contains('Ufr="disabled"')) {
        Write-Check "zcode.cjs 档位" "受限" "默认档位 disabled（可用 enable-dynamic-workflow.ps1 一键改为 onDemand）"
    } else {
        Write-Check "zcode.cjs 档位" "受限" "未找到 Ufr 锚点（版本不同）"
    }
} else {
    Write-Check "zcode.cjs 档位" "缺失" "未找到 resources\glm\zcode.cjs，无法检测"
}
# 3.2 备选锚点：app.asar 精确偏移探测（按需 seek 读 8 字节，绝不整文件解码）
$asar = $null
$asarCandidates = @(
    "$env:USERPROFILE\.zcode\app\resources\app.asar",
    "$env:LOCALAPPDATA\Programs\zcode\resources\app.asar"
)
foreach ($c in $asarCandidates) { if (Test-Path $c) { $asar = $c; break } }
$asarOffset = 280109971   # 用户 Windows 内网机实测偏移；安装特定值，跨机/跨版本必须重新定位
if ($asar) {
    $asarLen = (Get-Item $asar).Length
    if ($asarLen -le $asarOffset) {
        Write-Check "asar 补丁点" "受限" "app.asar 仅 $asarLen 字节，小于探测偏移 $asarOffset（版本不同），该偏移对本机无效"
    } else {
        $fs = [System.IO.File]::OpenRead($asar)
        try {
            $fs.Position = $asarOffset
            $buf = New-Object byte[] 8
            $read = 0
            while ($read -lt 8) {
                $n = $fs.Read($buf, $read, 8 - $read)
                if ($n -le 0) { break }
                $read += $n
            }
            $peek = [System.Text.Encoding]::ASCII.GetString($buf, 0, $read)
        } finally { $fs.Dispose() }
        if ($read -lt 8) {
            Write-Check "asar 补丁点" "受限" "偏移 $asarOffset 处不足 8 字节可读（实际 $read），文件异常"
        } elseif ($peek -eq "disabled") {
            Write-Check "asar 补丁点" "OK" "偏移 $asarOffset 处为 'disabled'（host 侧 resolveDynamicWorkflowGate 兜底，与用户实测吻合，可直接补丁）"
        } elseif ($peek -eq "onDemand") {
            Write-Check "asar 补丁点" "OK" "偏移 $asarOffset 处已是 'onDemand'（已补丁）"
        } else {
            Write-Check "asar 补丁点" "受限" "偏移 $asarOffset 处实际内容为 '$peek'（非 disabled/onDemand）——该偏移是安装特定值，本机无效；禁止盲替换，定位方法见 REPORT.md 7.2"
        }
    }
} else {
    Write-Check "asar 补丁点" "缺失" "未找到 app.asar，无法检测"
}
# 3.3 灰度接口指向（REPORT.md 7.3 方案B 的前提检查）
$baseUrl = $env:ZCODE_BASE_URL
if ($baseUrl) {
    Write-Check "灰度接口指向" "OK" "已设 ZCODE_BASE_URL=$baseUrl（host 的 /api/v1/client/configs 请求将指向该地址）"
} else {
    Write-Check "灰度接口指向" "受限" "未设 ZCODE_BASE_URL（默认官方 https://zcode.z.ai/api/v1/client/configs，内网不可达即兜底 disabled；纯配置方案见 REPORT.md 7.3 方案B）"
}
# 3.4 CLI 零修改方案提示
Write-Host "  提示：CLI/脚本会话可用零修改方案：zcode --workflow-mode onDemand|alwaysOn（见 REPORT.md 2.2 途径8）" -ForegroundColor DarkGray

# ---- 4. 自动更新 / 遥测 ----
Write-Host "--- 4. 自动更新与遥测 ---"
try {
    $upd = Get-ItemProperty "HKCU:\Software\ZCode" -Name AutoUpdate -ErrorAction Stop
    $updStatus = if ($upd.AutoUpdate) { "受限" } else { "OK" }
    Write-Check "自动更新" $updStatus "注册表 AutoUpdate=$($upd.AutoUpdate)（内网下更新接口不可达静默失败；已打补丁的机器建议保持关闭以防覆盖）"
} catch {
    Write-Check "自动更新" "受限" "未发现显式配置（默认行为待确认；内网下更新接口不可达，静默失败）"
}

# ---- 5. 本地功能确认 ----
Write-Host "--- 5. 本地功能 ---"
if (Test-Path (Join-Path $zcodeHome "v2\tasks-index.sqlite")) {
    Write-Check "任务库" "OK" "tasks-index.sqlite 在位（本地持久化）"
} else {
    Write-Check "任务库" "缺失" "未找到任务库"
}
if (Test-Path (Join-Path $zcodeHome "workspace\default\plugins")) {
    Write-Check "插件目录" "OK" "workspace 插件目录在位"
}

Write-Host ""
Write-Host "--- 检出汇总 ---"
$script:results | Format-Table -AutoSize | Out-Host
$okN = ($script:results | Where-Object { $_.状态 -eq "OK" }).Count
$warnN = ($script:results | Where-Object { $_.状态 -eq "受限" }).Count
$missN = ($script:results | Where-Object { $_.状态 -notin @("OK","受限") }).Count
Write-Host ("统计：OK={0} 受限={1} 缺失/其它={2}" -f $okN, $warnN, $missN)
Write-Host "===== 体检完成 =====" -ForegroundColor Cyan
Write-Host "受限/缺失项参考 REPORT.md 的离线方案处理；CUA 入口修复参考 config/ 下配置建议。"
