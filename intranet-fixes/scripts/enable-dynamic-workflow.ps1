#requires -Version 5.1
<#
.SYNOPSIS
    内网环境启用 ZCode 动态工作流（两路补丁：优先散装 zcode.cjs，备选 asar 偏移）
.DESCRIPTION
    官方正式版 ZCode 构建内核/host 子进程环境时删除 ZCODE_DYNAMIC_WORKFLOW_MODE，
    isPreview 构建期烙死 "production"，真正门禁为内核默认档位：
    env → 远端档位(remote.mode) → 默认 "disabled"（内网云端不可达即落到默认）。

    本脚本按推荐顺序执行（全程"先校验、后备份、定点替换"，任何校验不符即报错退出，禁止盲替换）：
      [首选] 途径9：resources\glm\zcode.cjs 中 Ufr="disabled"（全文件唯一锚点）
             等长替换为 Ufr="onDemand" —— 普通文本文件，无 asar/JSON 结构风险，回滚即还原。
             注意：锚点必须全文件唯一（count=1）；count=0 或 >1 均不替换并转 asar 分支。
      [备选] 途径7：app.asar 偏移 280109971 处（用户实测值，host 侧 resolveDynamicWorkflowGate
             兜底 "disabled"）—— 校验偏移处 8 字节：disabled → 补丁；onDemand → 幂等跳过；
             其它内容 → 报错退出（该偏移是安装特定值，须先用 check-intranet.ps1 / REPORT 7.2 定位）。
    流程：校验锚点 → 备份 → 定点替换 → 复验 → 注册计划任务（防自动更新/异常恢复覆盖）。
.NOTES
    与 offline-office-docs 完全解耦，不影响其任何文件。
    建议在 64 位 PowerShell 下运行：.cjs 分支整读约 15MB；asar 分支只做 seek 读/写 8 字节，
    不再整文件解码（历史版本曾 ReadAllBytes 340MB，32 位下会 OOM）。
    桌面 Composer 生效位在 host 侧（app.asar）；内核 .cjs 的 Ufr 档位主要影响 CLI 直连
    场景（CLI 另有零修改方案 --workflow-mode onDemand|alwaysOn）。见 REPORT.md 第七节。
#>

param(
    [string]$ZcodeRoot = "",          # ZCode 安装根目录（含 resources\glm\zcode.cjs）
    [switch]$Rollback,                # 回滚（按文件种类分别恢复最新备份）
    [switch]$SkipPlanTask,            # 跳过计划任务注册
    [switch]$UseAsarOnly              # 只用 asar 补丁（不优先 .cjs）
)

$ErrorActionPreference = "Stop"

if (-not [Environment]::Is64BitProcess) {
    Write-Warning "当前是 32 位 PowerShell：.cjs 分支需整文件读入约 15MB（bytes+ASCII 串各一份），32 位进程可用但余量小；建议用 64 位 powershell.exe 重跑。"
}

# ---- 路径探测 ----
function Resolve-ZcodeRoot {
    param([string]$Hint)
    if ($Hint -and (Test-Path $Hint)) { return $Hint }
    # 优先含散装 zcode.cjs 的布局；-UseAsarOnly 时也接受仅有 app.asar 的布局
    foreach ($c in @("$env:USERPROFILE\.zcode\app", "$env:LOCALAPPDATA\Programs\zcode", "$env:ProgramFiles\zcode", "C:\zcode")) {
        if ((Test-Path (Join-Path $c "resources\glm\zcode.cjs")) -or
            ($UseAsarOnly -and (Test-Path (Join-Path $c "resources\app.asar")))) { return $c }
    }
    throw "未找到 ZCode 安装目录（需含 resources\glm\zcode.cjs，或 -UseAsarOnly 时含 resources\app.asar），请用 -ZcodeRoot 指定"
}

# ---- ASCII 定位：返回 needle 的全部偏移（上限 $Limit 个）----
function Find-AsciiOffsets {
    param([byte[]]$Hay, [string]$Needle, [int]$Limit = 64)
    $s = [System.Text.Encoding]::ASCII.GetString($Hay)   # ASCII 解码保证字节↔字符 1:1，偏移不出错
    $list = New-Object System.Collections.Generic.List[long]
    $pos = 0
    while ($list.Count -lt $Limit) {
        $i = $s.IndexOf($Needle, $pos, [System.StringComparison]::Ordinal)
        if ($i -lt 0) { break }
        $list.Add([long]$i)
        $pos = $i + 1
    }
    return ,$list
}

# ---- 定点等长替换（FileStream seek 读写，任意大小文件均不整读）----
function Invoke-OffsetReplace {
    param([string]$Path, [long]$Offset, [string]$From, [string]$To)
    $enc = [System.Text.Encoding]::ASCII
    $bFrom = $enc.GetBytes($From); $bTo = $enc.GetBytes($To)
    if ($bFrom.Length -ne $bTo.Length) { throw "长度不一致（$From → $To），等长替换要求 ASCII 长度相同" }
    $buf = New-Object byte[] $bFrom.Length
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    try {
        # 幂等：已是目标值则跳过
        $fs.Position = $Offset
        $null = $fs.Read($buf, 0, $buf.Length)
        if ($enc.GetString($buf) -eq $To) { Write-Host "偏移 $Offset 已是 $To，幂等跳过"; return }
        # 校验：必须精确等于源值，否则报错退出（禁止盲替换）
        if ($enc.GetString($buf) -ne $From) {
            $actual = ($enc.GetString($buf) -replace '[^\x20-\x7E]', '?')
            throw "偏移 $Offset 处内容为 '$actual'（非 '$From' 也非 '$To'），禁止盲替换；请先用 check-intranet.ps1 定位本机实际锚点（REPORT.md 7.2）"
        }
        $fs.Position = $Offset
        $fs.Write($bTo, 0, $bTo.Length)
        $fs.Flush()
        # 复验
        $fs.Position = $Offset
        $null = $fs.Read($buf, 0, $buf.Length)
        if ($enc.GetString($buf) -ne $To) { throw "写入后复验失败：偏移 $Offset 处非 '$To'" }
        Write-Host "定点替换成功并复验通过 @ $Offset"
    } finally { $fs.Dispose() }
}

# ---- 主流程 ----
$root = Resolve-ZcodeRoot -Hint $ZcodeRoot
Write-Host "ZCode 安装目录：$root"

$cjsPath = Join-Path $root "resources\glm\zcode.cjs"
$asarPath = Join-Path $root "resources\app.asar"
$asarOffset = 280109971   # 用户 Windows 内网机实测偏移（host 侧 resolveDynamicWorkflowGate 兜底）；安装特定值
$backupDir = Join-Path $PSScriptRoot "..\logs-backup"
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"

# 回滚模式：按文件种类分别恢复最新一份"补丁前"备份
if ($Rollback) {
    $did = $false
    $cjsBak = Get-ChildItem $backupDir -Filter "zcode.cjs.bak-*" -File -ErrorAction SilentlyContinue |
              Sort-Object Name -Descending | Select-Object -First 1
    if ($cjsBak) {
        if (-not (Test-Path $cjsPath)) { throw "存在 .cjs 备份但未找到目标 $cjsPath，请用 -ZcodeRoot 指定安装目录" }
        Copy-Item $cjsBak.FullName $cjsPath -Force
        Write-Host "已回滚：$($cjsBak.Name) → $cjsPath"
        $did = $true
    }
    $asarBak = Get-ChildItem $backupDir -Filter "app.asar.bak-*" -File -ErrorAction SilentlyContinue |
               Sort-Object Name -Descending | Select-Object -First 1
    if ($asarBak) {
        if (-not (Test-Path $asarPath)) { throw "存在 asar 备份但未找到目标 $asarPath，请用 -ZcodeRoot 指定安装目录" }
        Copy-Item $asarBak.FullName $asarPath -Force
        Write-Host "已回滚：$($asarBak.Name) → $asarPath"
        $did = $true
    }
    if (-not $did) { throw "未找到可回滚的备份（$backupDir 下无 zcode.cjs.bak-* / app.asar.bak-*）" }
    Write-Host "回滚完成。请完全退出并重启 ZCode。"
    exit 0
}

# 1) 首选：散装 zcode.cjs 补丁（锚点全文件唯一才动手）
$cjsPatched = $false
$dwEnabled = $false
if (-not $UseAsarOnly -and (Test-Path $cjsPath)) {
    Write-Host ""
    Write-Host "=== 首选方案：散装 zcode.cjs（Ufr=`"disabled`" → Ufr=`"onDemand`"）===" -ForegroundColor Cyan
    try {
        $bytes = [System.IO.File]::ReadAllBytes($cjsPath)
        $anchor = 'Ufr="disabled"'
        $target = 'Ufr="onDemand"'
        $targetHits = @(Find-AsciiOffsets -Hay $bytes -Needle $target)
        if ($targetHits.Count -gt 0) {
            Write-Host "zcode.cjs 已是 onDemand（偏移 $([int]$targetHits[0])），幂等跳过"
            $cjsPatched = $true
            $dwEnabled = $true
        } else {
            $hits = @(Find-AsciiOffsets -Hay $bytes -Needle $anchor)
            if ($hits.Count -eq 0) {
                Write-Warning "zcode.cjs 未找到锚点 $anchor（版本不同或已被改动），转 asar 分支"
            } elseif ($hits.Count -gt 1) {
                Write-Warning ("zcode.cjs 锚点 {0} 不唯一（count={1}），禁止盲替换，转 asar 分支" -f $anchor, $hits.Count)
            } else {
                $off = $hits[0]
                $bak = Join-Path $backupDir "zcode.cjs.bak-$stamp"
                Copy-Item $cjsPath $bak -Force
                Write-Host "已备份：$bak"
                Invoke-OffsetReplace -Path $cjsPath -Offset $off -From $anchor -To $target
                $cjsPatched = $true
                $dwEnabled = $true
            }
        }
    } catch {
        Write-Warning "散装 .cjs 补丁失败：$($_.Exception.Message)"
        Write-Host "转 asar 备选分支..."
    }
}

# 2) 备选：asar 偏移补丁（先校验偏移处 8 字节，再备份，再 seek 定点写）
if (-not $cjsPatched -and (Test-Path $asarPath)) {
    Write-Host ""
    Write-Host "=== 备选方案：app.asar 偏移（disabled → onDemand）===" -ForegroundColor Cyan
    try {
        $len = (Get-Item $asarPath).Length
        if ($len -le $asarOffset) { throw "app.asar 仅 $len 字节，小于偏移 $asarOffset（版本不同），该偏移对本机无效" }
        # 先校验
        $peekBuf = New-Object byte[] 8
        $fs = [System.IO.File]::OpenRead($asarPath)
        try {
            $fs.Position = $asarOffset
            $null = $fs.Read($peekBuf, 0, 8)
        } finally { $fs.Dispose() }
        $peek = [System.Text.Encoding]::ASCII.GetString($peekBuf)
        if ($peek -eq "onDemand") {
            Write-Host "asar 偏移 $asarOffset 已是 onDemand，幂等跳过"
            $dwEnabled = $true
        } elseif ($peek -eq "disabled") {
            $bak = Join-Path $backupDir "app.asar.bak-$stamp"
            Copy-Item $asarPath $bak -Force
            Write-Host "已备份：$bak"
            Invoke-OffsetReplace -Path $asarPath -Offset $asarOffset -From "disabled" -To "onDemand"
            $dwEnabled = $true
        } else {
            $actual = ($peek -replace '[^\x20-\x7E]', '?')
            throw "偏移 $asarOffset 处内容为 '$actual'（非 disabled/onDemand）——该偏移是安装特定值，禁止盲替换；定位方法见 check-intranet.ps1 与 REPORT.md 7.2"
        }
    } catch {
        Write-Warning "asar 备选分支失败：$($_.Exception.Message)"
        Write-Host "提示：纯配置替代（桌面 Composer）见 REPORT.md 7.3 方案B（ZCODE_BASE_URL + 企业 CA）；CLI 场景直接 --workflow-mode 零修改。"
    }
}

# 3) 计划任务（防自动更新/异常恢复覆盖；补丁幂等，重跑安全）
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
if ($dwEnabled) {
    Write-Host "完成（动态工作流补丁在位）。请完全退出并重启 ZCode。" -ForegroundColor Green
} else {
    Write-Warning "未做任何更改：两路分支均未生效（.cjs 缺失/锚点不符，或 asar 缺失/偏移不符）。见上方告警；桌面 Composer 纯配置替代见 REPORT.md 7.3 方案B，CLI 直接 --workflow-mode 零修改。"
}
Write-Host "回滚：powershell -File `"$($MyInvocation.MyCommand.Path)`" -Rollback"
Write-Host "CLI/脚本会话零修改方案：zcode --workflow-mode onDemand（详见 REPORT.md 2.2 途径8、第七节方案A）"
