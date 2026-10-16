#requires -Version 5.1
<#
.SYNOPSIS
    内网环境一键启用 ZCode 动态工作流（字节补丁方案）
.DESCRIPTION
    官方正式版 ZCode 在构建内核/host 子进程环境时删除 ZCODE_DYNAMIC_WORKFLOW_MODE，
    且 isPreview 为构建期烙死的 "production"（正式版永远非 Preview），
    真正门禁为内核 host 内云端灰度档位：云端不可达（内网）时兜底 "disabled"。

    本脚本将 app.asar 中偏移 280109971 处的 "disabled" 原地等长替换为 "onDemand"
    （8 字节等长，无需调整 JSON 目录长度），使内核在"云端不可达"时按 onDemand 处理。
    流程：备份 → 校验定位 → 打补丁 → 校验结果 → 注册计划任务（防止自动更新覆盖）。
.NOTES
    适用范围：ZCode 3.15.3 linux-x64 解包版（squashfs-root 布局）。
    若你的安装是其它版本/平台，请先运行 check-intranet.ps1 定位实际偏移后再执行。
    本脚本为本地功能开关调整，不影响其它插件；与 offline-office-docs 完全解耦。
#>

param(
    [string]$AsarPath = "",
    [switch]$Rollback,
    [switch]$SkipPlanTask
)

$ErrorActionPreference = "Stop"

# ---- 路径自动探测 ----
function Resolve-AsarPath {
    param([string]$Hint)
    if ($Hint -and (Test-Path $Hint)) { return $Hint }
    $candidates = @(
        "$env:USERPROFILE\.zcode\app\resources\app.asar",
        "$env:LOCALAPPDATA\Programs\zcode\resources\app.asar",
        "$env:ProgramFiles\zcode\resources\app.asar",
        "C:\zcode\resources\app.asar"
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    throw "未找到 app.asar，请用 -AsarPath 指定完整路径"
}

# ---- 等长字节替换 ----
function Invoke-BytesPatch {
    param(
        [string]$Path,
        [string]$From,      # 原字符串（如 disabled）
        [string]$To,        # 目标字符串（如 onDemand，长度必须一致）
        [long]$Offset = -1
    )
    $bFrom = [System.Text.Encoding]::ASCII.GetBytes($From)
    $bTo   = [System.Text.Encoding]::ASCII.GetBytes($To)
    if ($bFrom.Length -ne $bTo.Length) {
        throw "目标字符串与源字符串长度不一致（$From → $To），等长替换要求长度相同"
    }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $found = $false
    if ($Offset -ge 0) {
        # 定点偏移
        if ($Offset + $bFrom.Length -gt $bytes.Length) { throw "偏移越界" }
        $hit = $true
        for ($i = 0; $i -lt $bFrom.Length; $i++) {
            if ($bytes[$Offset + $i] -ne $bFrom[$i]) { $hit = $false; break }
        }
        if (-not $hit) { throw "偏移 $Offset 处内容不是预期的 '$From'，可能版本不同，请勿强行打补丁" }
        $bytes[$Offset..($Offset + $bFrom.Length - 1)] = $bTo
        $found = $true
        Write-Host "定点替换成功 @ $Offset"
    } else {
        # 全文件搜索第一次出现
        for ($i = 0; $i -le $bytes.Length - $bFrom.Length; $i++) {
            $hit = $true
            for ($j = 0; $j -lt $bFrom.Length; $j++) {
                if ($bytes[$i + $j] -ne $bFrom[$j]) { $hit = $false; break }
            }
            if ($hit) {
                $bytes[$i..($i + $bFrom.Length - 1)] = $bTo
                Write-Host "全文件搜索命中并替换 @ $i"
                $found = $true
                break
            }
        }
        if (-not $found) { throw "全文件中未找到 '$From'，可能已打过补丁或版本不同" }
    }
    [System.IO.File]::WriteAllBytes($Path, $bytes)
    Write-Host "写入完成"
}

# ---- 主流程 ----
$asar = Resolve-AsarPath -Hint $AsarPath
Write-Host "使用 asar：$asar"

# 1) 备份现场
$backupDir = Join-Path $PSScriptRoot "..\logs-backup"
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = Join-Path $backupDir "app.asar.bak-$stamp"
Copy-Item $asar $backup -Force
Write-Host "已备份到：$backup"

# 2) 回滚模式
if ($Rollback) {
    $latest = Get-ChildItem $backupDir -Filter "app.asar.bak-*" | Sort-Object Name -Descending | Select-Object -First 1
    if (-not $latest) { throw "未找到任何备份，无法回滚" }
    Copy-Item $latest.FullName $asar -Force
    Write-Host "已回滚：$($latest.Name)"
    exit 0
}

# 3) 打补丁（默认定点偏移 280109971，版本不符时自动转全文件搜索）
try {
    Invoke-BytesPatch -Path $asar -From "disabled" -To "onDemand" -Offset 280109971
} catch {
    Write-Warning "定点补丁失败：$($_.Exception.Message)"
    Write-Host "尝试全文件搜索替换（等长 disabled→onDemand）..."
    Invoke-BytesPatch -Path $asar -From "disabled" -To "onDemand"
}

# 4) 校验
$bytes = [System.IO.File]::ReadAllBytes($asar)
$probe = [System.Text.Encoding]::ASCII.GetString($bytes)
if ($probe.Contains("onDemand") -and -not $probe.Substring(0, 280109971 + 8).Contains("onDemand")) {
    Write-Host "补丁校验通过：目标区已变为 onDemand"
} else {
    Write-Host "补丁已写入（校验提示：请在重启后确认动态工作流状态）"
}

# 5) 注册计划任务（防止自动更新/异常恢复覆盖补丁）
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
Write-Host "完成。请完全退出 ZCode 后重新启动，动态工作流应可按 onDemand 生效。"
Write-Host "如需回滚：powershell -File `"$($MyInvocation.MyCommand.Path)`" -Rollback"
