#requires -Version 5.1
<#
.SYNOPSIS
    内网环境 ZCode 功能体检脚本：一键排查哪些功能不可用/受限，并给出离线方案。
.NOTES
    纯 PowerShell 5.1，无需外网。只读检查，不改任何配置。
#>

$ErrorActionPreference = "Continue"

Write-Host "===== ZCode 内网功能体检 =====" -ForegroundColor Cyan
Write-Host ""

$zcodeHome = "$env:USERPROFILE\.zcode"
$results = @()

function Write-Check {
    param([string]$Name, [string]$Status, [string]$Detail)
    $results += [PSCustomObject]@{ 功能=$Name; 状态=$Status; 说明=$Detail }
    $color = if ($Status -eq "OK") { "Green" } elseif ($Status -eq "受限") { "Yellow" } else { "Red" }
    Write-Host ("[{0}] {1}: {2}" -f $Status, $Name, $Detail) -ForegroundColor $color
}

# ---- 1. 网络可达性 ----
Write-Host "--- 1. 外网/云端可达性 ---"
$hosts = @(
    @{ N="云端授权服务(CUA)"; U="https://z.ai" },
    @{ N="插件市场"; U="https://marketplace.zcode.local" },
    @{ N="灰度配置(动态工作流)"; U="https://config.zcode.local" }
)
$netOK = $false
foreach ($h in $hosts) {
    try {
        $r = Invoke-WebRequest -Uri $h.U -Method Head -TimeoutSec 4 -UseBasicParsing -ErrorAction Stop
        Write-Check $h.N "OK" "可达 (HTTP $($r.StatusCode))"
        $netOK = $true
    } catch {
        Write-Check $h.N "受限" "内网不可达（预期内）"
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
# 3.2 备选锚点：app.asar 偏移
$asar = $null
$asarCandidates = @(
    "$env:USERPROFILE\.zcode\app\resources\app.asar",
    "$env:LOCALAPPDATA\Programs\zcode\resources\app.asar"
)
foreach ($c in $asarCandidates) { if (Test-Path $c) { $asar = $c; break } }
if ($asar) {
    $bytes = [System.IO.File]::ReadAllBytes($asar)
    $s = [System.Text.Encoding]::ASCII.GetString($bytes)
    $off = $s.IndexOf("disabled", 280000000)
    if ($off -ge 0) {
        Write-Check "asar 补丁点" "OK" "找到 disabled 位于偏移 $off（用户实测 280109971；以本机为准）"
        if ($off -ne 280109971) {
            Write-Warning "  实际偏移 $off ≠ 用户实测 280109971，打补丁时请用此实际偏移"
        }
    } else {
        Write-Check "asar 补丁点" "受限" "未在预期区间找到 disabled（可能已打过补丁或版本不同）"
    }
} else {
    Write-Check "asar 补丁点" "缺失" "未找到 app.asar，无法检测"
}
# 3.3 CLI 零修改方案提示
Write-Host "  提示：CLI/脚本会话可用零修改方案：zcode --workflow-mode onDemand|alwaysOn（见 REPORT.md 2.2 途径8）" -ForegroundColor DarkGray

# ---- 4. 自动更新 / 遥测 ----
Write-Host "--- 4. 自动更新与遥测 ---"
try {
    $upd = Get-ItemProperty "HKCU:\Software\ZCode" -Name AutoUpdate -ErrorAction Stop
    Write-Check "自动更新" $upd.AutoUpdate
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
Write-Host "===== 体检完成 =====" -ForegroundColor Cyan
Write-Host "受限/缺失项参考 REPORT.md 的离线方案处理；CUA 入口修复参考 config/ 下配置建议。"
