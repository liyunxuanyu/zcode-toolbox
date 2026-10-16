#requires -Version 5.1
<#
.SYNOPSIS
    内网环境启用 ZCode 动态工作流（两路补丁：优先散装 zcode.cjs，备选 asar 偏移）
.DESCRIPTION
    官方正式版 ZCode 构建内核/host 子进程环境时删除 ZCODE_DYNAMIC_WORKFLOW_MODE，
    isPreview 构建期烙死 "production"，真正门禁为内核默认档位：
    env → 远端档位(remote.mode) → 默认 "disabled"（内网云端不可达即落到默认）。

    本脚本按推荐顺序执行：
      [首选] 途径9：resources\glm\zcode.cjs 中 Ufr="disabled"（全文件唯一锚点）
             等长替换为 Ufr="onDemand" —— 普通文本文件，无 asar/JSON 结构风险，回滚即还原。
      [备选] 途径7：app.asar 偏移 280109971 处 "disabled"→"onDemand"（用户实测偏移；
             本机如不命中自动转全文件搜索）。
    流程：备份 → 校验锚点 → 打补丁 → 校验 → 注册计划任务（防自动更新/异常恢复覆盖）。
.NOTES
    与 offline-office-docs 完全解耦，不影响其任何文件。
    若你主要用 CLI/脚本会话，还有零修改方案：内核 CLI --workflow-mode onDemand|alwaysOn。
#>

param(
    [string]$ZcodeRoot = "",          # ZCode 安装根目录（含 resources\glm\zcode.cjs）
    [switch]$Rollback,                # 回滚
    [switch]$SkipPlanTask,            # 跳过计划任务注册
    [switch]$UseAsarOnly              # 只用 asar 补丁（不优先 .cjs）
)

$ErrorActionPreference = "Stop"

# ---- 路径探测 ----
function Resolve-ZcodeRoot {
    param([string]$Hint)
    if ($Hint -and (Test-Path $Hint)) { return $Hint }
    $candidates = @(
        "$env:USERPROFILE\.zcode\app",
        "$env:LOCALAPPDATA\Programs\zcode",
        "$env:ProgramFiles\zcode",
        "C:\zcode"
    )
    foreach ($c in $candidates) { if (Test-Path (Join-Path $c "resources\glm\zcode.cjs")) { return $c } }
    throw "未找到 ZCode 安装目录（需含 resources\glm\zcode.cjs），请用 -ZcodeRoot 指定"
}

# ---- 等长 ASCII 字节替换 ----
function Invoke-LiteralReplace {
    param([string]$Path, [string]$From, [string]$To, [long]$Offset = -1)
    $enc = [System.Text.Encoding]::UTF8
    $bFrom = $enc.GetBytes($From); $bTo = $enc.GetBytes($To)
    if ($bFrom.Length -ne $bTo.Length) {
        throw "长度不一致（$From → $To），等长替换要求 ASCII 长度相同"
    }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($Offset -ge 0) {
        for ($i = 0; $i -lt $bFrom.Length; $i++) {
            if ($bytes[$Offset + $i] -ne $bFrom[$i]) { throw "偏移 $Offset 处非预期内容，请先 run check-intranet.ps1 定位" }
        }
        $bytes[$Offset..($Offset + $bFrom.Length - 1)] = $bTo
        Write-Host "定点替换成功 @ $Offset"
    } else {
        $found = $false
        for ($i = 0; $i -le $bytes.Length - $bFrom.Length; $i++) {
            $hit = $true
            for ($j = 0; $j -lt $bFrom.Length; $j++) {
                if ($bytes[$i + $j] -ne $bFrom[$j]) { $hit = $false; break }
            }
            if ($hit) { $bytes[$i..($i + $bFrom.Length - 1)] = $bTo; Write-Host "全文件命中并替换 @ $i"; $found = $true; break }
        }
        if (-not $found) { throw "未找到 '$From'（可能已补丁或版本不同）" }
    }
    [System.IO.File]::WriteAllBytes($Path, $bytes)
    Write-Host "写入完成：$Path"
}

# ---- 主流程 ----
$root = Resolve-ZcodeRoot -Hint $ZcodeRoot
Write-Host "ZCode 安装目录：$root"

$cjsPath = Join-Path $root "resources\glm\zcode.cjs"
$asarPath = Join-Path $root "resources\app.asar"
$backupDir = Join-Path $PSScriptRoot "..\logs-backup"
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"

# 回滚模式
if ($Rollback) {
    $latest = Get-ChildItem $backupDir -Filter "*.bak-*" | Sort-Object Name -Descending | Select-Object -First 1
    if (-not $latest) { throw "未找到备份，无法回滚" }
    $target = if ($latest.Name -like "zcode.cjs*") { $cjsPath } else { $asarPath }
    Copy-Item $latest.FullName $target -Force
    Write-Host "已回滚：$($latest.Name) → $target"
    exit 0
}

# 1) 首选：散装 zcode.cjs 补丁
$cjsPatched = $false
if (-not $UseAsarOnly -and (Test-Path $cjsPath)) {
    Write-Host ""
    Write-Host "=== 首选方案：散装 zcode.cjs（Ufr=disabled → onDemand）===" -ForegroundColor Cyan
    try {
        $bytes = [System.IO.File]::ReadAllBytes($cjsPath)
        $probe = [System.Text.Encoding]::UTF8.GetString($bytes)
        if ($probe.Contains('Ufr="onDemand"')) {
            Write-Host "zcode.cjs 已是 onDemand，跳过"
            $cjsPatched = $true
        } else {
            $bak = Join-Path $backupDir "zcode.cjs.bak-$stamp"
            Copy-Item $cjsPath $bak -Force
            Write-Host "已备份：$bak"
            Invoke-LiteralReplace -Path $cjsPath -From 'Ufr="disabled"' -To 'Ufr="onDemand"'
            $cjsPatched = $true
        }
        # 校验
        $bytes2 = [System.IO.File]::ReadAllBytes($cjsPath)
        $s2 = [System.Text.Encoding]::UTF8.GetString($bytes2)
        if ($s2.Contains('Ufr="onDemand"') -and -not $s2.Contains('Ufr="disabled"')) {
            Write-Host "校验通过：zcode.cjs 已启用 onDemand"
        } else {
            Write-Warning "校验未完全通过，请检查（可能同时存在 disabled 残留）"
        }
    } catch {
        Write-Warning "散装 .cjs 补丁失败：$($_.Exception.Message)"
        Write-Host "自动降级到 asar 补丁..."
    }
}

# 2) 备选：asar 偏移补丁
if (-not $cjsPatched -and (Test-Path $asarPath)) {
    Write-Host ""
    Write-Host "=== 备选方案：app.asar 偏移（disabled → onDemand）===" -ForegroundColor Cyan
    $bak = Join-Path $backupDir "app.asar.bak-$stamp"
    Copy-Item $asarPath $bak -Force
    Write-Host "已备份：$bak"
    try {
        Invoke-LiteralReplace -Path $asarPath -From "disabled" -To "onDemand" -Offset 280109971
    } catch {
        Write-Warning "定点失败：$($_.Exception.Message)，转全文件搜索..."
        Invoke-LiteralReplace -Path $asarPath -From "disabled" -To "onDemand"
    }
    Write-Host "asar 补丁完成"
}

# 3) 计划任务（防自动更新/异常恢复覆盖）
if (-not $SkipPlanTask) {
    $taskName = "ZCode-DynamicWorkflow-Patch-Keep"
    $scriptPath = $MyInvocation.MyCommand.Path
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
    try {
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
        Write-Host "已注册计划任务 $taskName（登录时自愈补丁）"
    } catch {
        Write-Warning "注册计划任务失败（可能需要管理员权限）：$($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "完成。请完全退出并重启 ZCode。"
Write-Host "回滚：powershell -File `"$($MyInvocation.MyCommand.Path)`" -Rollback"
Write-Host "CLI/脚本会话零修改方案：zcode --workflow-mode onDemand（详见 REPORT.md 2.2 途径8）"
