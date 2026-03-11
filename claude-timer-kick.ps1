<#
.SYNOPSIS
    Agent Timer Kicker - Claude Code / Codex CLI の5時間ローリングウィンドウを自動で整えます。
.DESCRIPTION
    Windows タスクスケジューラに軽量プロンプト送信タスクを登録し、
    ユーザーが回復していてほしい時刻の5時間前に kick を行います。
.NOTES
    要件: Windows 10+, PowerShell 5.1+, claude-code and/or codex CLI 認証済み。
    実行: powershell -ExecutionPolicy Bypass -File claude-timer-kick.ps1
    または launch.bat をダブルクリック。
#>

# ── 設定 ─────────────────────────────────────────────────────────────
$Script:AppName               = "Agent Recovery Scheduler"
$Script:ScriptDir             = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:ConfigDir             = Join-Path $Script:ScriptDir "data"
$Script:ConfigFile            = Join-Path $Script:ConfigDir "config.json"
$Script:LogFile               = Join-Path $Script:ConfigDir "kick.log"
$Script:TaskPrefix            = "AgentTimerKick"
$Script:DefaultPrompt         = "just say hi and nothing else."
$Script:ClaudeModel           = "haiku"
$Script:CodexModel            = "gpt-5-codex-mini"
$Script:WindowHours           = 5
$Script:WindowMinutes         = $Script:WindowHours * 60
$Script:RetryIntervalMinutes  = 30
$Script:RetryDurationMinutes  = 120
$Script:MinimumSpacingMinutes = $Script:WindowMinutes + 10
$Script:MinimumSpacingLabel   = "5時間10分"
$Script:WeekdayOptions        = @(
    [PSCustomObject]@{ Name = "Monday";    Label = "月" }
    [PSCustomObject]@{ Name = "Tuesday";   Label = "火" }
    [PSCustomObject]@{ Name = "Wednesday"; Label = "水" }
    [PSCustomObject]@{ Name = "Thursday";  Label = "木" }
    [PSCustomObject]@{ Name = "Friday";    Label = "金" }
    [PSCustomObject]@{ Name = "Saturday";  Label = "土" }
    [PSCustomObject]@{ Name = "Sunday";    Label = "日" }
)
$Script:DefaultDaysOfWeek     = @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
$Script:DefaultRecoveryTimes  = @("11:00", "16:10", "21:20")

# ── データディレクトリ確保 ───────────────────────────────────────────
if (-not (Test-Path $Script:ConfigDir)) { New-Item -ItemType Directory -Path $Script:ConfigDir -Force | Out-Null }

# ── 共通ヘルパー ─────────────────────────────────────────────────────
function ConvertTo-TimeString ([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    $trimmed = $value.Trim()
    if ($trimmed -match '^([01]?\d|2[0-3]):([0-5]\d)$') {
        return "{0:D2}:{1}" -f [int]$matches[1], $matches[2]
    }
    return $null
}

function ConvertTo-Minutes ([string]$timeString) {
    $normalized = ConvertTo-TimeString $timeString
    if (-not $normalized) { return $null }
    $parts = $normalized.Split(":")
    return ([int]$parts[0] * 60) + [int]$parts[1]
}

function ConvertTo-TimeLabel ([int]$totalMinutes) {
    $minutes = (($totalMinutes % 1440) + 1440) % 1440
    $hour = [int][math]::Floor($minutes / 60)
    $minute = [int]($minutes % 60)
    return "{0:D2}:{1:D2}" -f $hour, $minute
}

function Get-RetryOffsets {
    $offsets = @()
    for ($offset = 0; $offset -le $Script:RetryDurationMinutes; $offset += $Script:RetryIntervalMinutes) {
        $offsets += $offset
    }
    return $offsets
}

function Get-NormalizedDaysOfWeek ($days) {
    $map = @{
        "monday"    = "Monday"
        "mon"       = "Monday"
        "月"        = "Monday"
        "tuesday"   = "Tuesday"
        "tue"       = "Tuesday"
        "火"        = "Tuesday"
        "wednesday" = "Wednesday"
        "wed"       = "Wednesday"
        "水"        = "Wednesday"
        "thursday"  = "Thursday"
        "thu"       = "Thursday"
        "木"        = "Thursday"
        "friday"    = "Friday"
        "fri"       = "Friday"
        "金"        = "Friday"
        "saturday"  = "Saturday"
        "sat"       = "Saturday"
        "土"        = "Saturday"
        "sunday"    = "Sunday"
        "sun"       = "Sunday"
        "日"        = "Sunday"
    }

    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($day in (ConvertTo-StringArray $days)) {
        $key = $day.ToString().Trim().ToLowerInvariant()
        if ($map.ContainsKey($key)) { [void]$set.Add($map[$key]) }
    }

    return @(
        foreach ($option in $Script:WeekdayOptions) {
            if ($set.Contains($option.Name)) { $option.Name }
        }
    )
}

function Get-DayOfWeekLabels ($days) {
    $labels = @()
    $normalized = Get-NormalizedDaysOfWeek $days
    foreach ($option in $Script:WeekdayOptions) {
        if ($normalized -contains $option.Name) { $labels += $option.Label }
    }
    return $labels
}

function ConvertTo-StringArray ($value) {
    if ($null -eq $value) { return @() }
    if ($value -is [string]) { return @($value) }
    return @($value)
}

function ConvertTo-Bool ($value, [bool]$defaultValue = $false) {
    if ($null -eq $value) { return $defaultValue }
    return [bool]$value
}

function ConvertTo-PSSingleQuotedLiteral ([string]$value) {
    return "'" + ($value -replace "'", "''") + "'"
}

function ConvertTo-BashSingleQuotedLiteral ([string]$value) {
    return "'" + ($value -replace "'", "'""'""'") + "'"
}

function ConvertTo-PowerShellArrayLiteral ([string[]]$values) {
    if ($null -eq $values -or $values.Count -eq 0) { return "@()" }

    $literals = @()
    foreach ($value in $values) {
        $literals += (ConvertTo-PSSingleQuotedLiteral $value)
    }
    return "@(" + ($literals -join ", ") + ")"
}

function Escape-ForPowerShellDoubleQuotedString ([string]$value) {
    return ($value -replace '`', '``' -replace '"', '`"')
}

function Get-NormalizedRecoveryTimes ($times) {
    $unique = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($time in (ConvertTo-StringArray $times)) {
        $normalized = ConvertTo-TimeString $time
        if ($normalized) { [void]$unique.Add($normalized) }
    }

    return @($unique | Sort-Object { ConvertTo-Minutes $_ })
}

function Get-DefaultConfig {
    return [PSCustomObject]@{
        Claude = [PSCustomObject]@{
            RecoveryTimes = @($Script:DefaultRecoveryTimes)
            DaysOfWeek = @($Script:DefaultDaysOfWeek)
            UseWSL = $false
        }
        Codex = [PSCustomObject]@{
            RecoveryTimes = @($Script:DefaultRecoveryTimes)
            DaysOfWeek = @($Script:DefaultDaysOfWeek)
            UseWSL = $false
        }
    }
}

function Normalize-ToolConfig ($toolConfig) {
    if ($null -eq $toolConfig) {
        return [PSCustomObject]@{
            RecoveryTimes = @($Script:DefaultRecoveryTimes)
            DaysOfWeek = @($Script:DefaultDaysOfWeek)
            UseWSL = $false
        }
    }

    $times = Get-NormalizedRecoveryTimes $toolConfig.RecoveryTimes
    if ($times.Count -eq 0) { $times = @($Script:DefaultRecoveryTimes) }
    $days = Get-NormalizedDaysOfWeek $toolConfig.DaysOfWeek
    if ($days.Count -eq 0) { $days = @($Script:DefaultDaysOfWeek) }

    return [PSCustomObject]@{
        RecoveryTimes = @($times)
        DaysOfWeek = @($days)
        UseWSL = ConvertTo-Bool $toolConfig.UseWSL
    }
}

function Load-Config {
    $defaultConfig = Get-DefaultConfig
    if (Test-Path $Script:ConfigFile) {
        try {
            $raw = Get-Content $Script:ConfigFile -Raw | ConvertFrom-Json
            return [PSCustomObject]@{
                Claude = Normalize-ToolConfig $raw.Claude
                Codex  = Normalize-ToolConfig $raw.Codex
            }
        } catch { }
    }
    return $defaultConfig
}

function Save-Config ($cfg) {
    $payload = [PSCustomObject]@{
        Claude = [PSCustomObject]@{
            RecoveryTimes = @($cfg.Claude.RecoveryTimes)
            DaysOfWeek = @($cfg.Claude.DaysOfWeek)
            UseWSL = [bool]$cfg.Claude.UseWSL
        }
        Codex = [PSCustomObject]@{
            RecoveryTimes = @($cfg.Codex.RecoveryTimes)
            DaysOfWeek = @($cfg.Codex.DaysOfWeek)
            UseWSL = [bool]$cfg.Codex.UseWSL
        }
    }
    $payload | ConvertTo-Json -Depth 4 | Set-Content $Script:ConfigFile -Encoding UTF8
}

function Get-RecoveryPlan ($times, $daysOfWeek) {
    $normalized = Get-NormalizedRecoveryTimes $times
    $selectedDays = Get-NormalizedDaysOfWeek $daysOfWeek
    $dayLabels = Get-DayOfWeekLabels $selectedDays
    $issues = New-Object 'System.Collections.Generic.List[string]'
    $items = New-Object 'System.Collections.Generic.List[object]'

    if ($normalized.Count -eq 0) {
        $issues.Add("回復時刻を1件以上追加してください。")
        return [PSCustomObject]@{
            IsValid = $false
            RecoveryTimes = @()
            DaysOfWeek = @()
            DayLabels = @()
            Items = @()
            KickMinutes = @()
            Issues = @($issues)
        }
    }

    if ($selectedDays.Count -eq 0) {
        $issues.Add("曜日を1つ以上選択してください。")
    }

    $minutes = @()
    foreach ($time in $normalized) { $minutes += ConvertTo-Minutes $time }

    for ($i = 0; $i -lt $minutes.Count; $i++) {
        $current = $minutes[$i]
        $nextIndex = ($i + 1) % $minutes.Count
        $next = $minutes[$nextIndex]
        if ($nextIndex -eq 0) { $next += 1440 }
        $gap = $next - $current
        if ($gap -lt $Script:MinimumSpacingMinutes) {
            $a = ConvertTo-TimeLabel $current
            $b = ConvertTo-TimeLabel $next
            $issues.Add("$a と $b は近すぎます。回復時刻どうしは $($Script:MinimumSpacingLabel) 以上あけてください。")
        }
    }

    $allKickMinutes = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($recovery in $minutes) {
        $kickStart = $recovery - $Script:WindowMinutes
        $kickOffsets = Get-RetryOffsets
        $kickAttemptMinutes = @()
        foreach ($offset in $kickOffsets) {
            $attempt = (($kickStart + $offset) % 1440 + 1440) % 1440
            $kickAttemptMinutes += [int]$attempt
            [void]$allKickMinutes.Add([int]$attempt)
        }

        $items.Add([PSCustomObject]@{
            RecoveryTime = ConvertTo-TimeLabel $recovery
            RecoveryMinutes = $recovery
            KickStartTime = ConvertTo-TimeLabel $kickStart
            KickEndTime = ConvertTo-TimeLabel ($kickStart + $Script:RetryDurationMinutes)
            KickAttemptMinutes = @($kickAttemptMinutes)
            AttemptCount = $kickAttemptMinutes.Count
        })
    }

    return [PSCustomObject]@{
        IsValid = ($issues.Count -eq 0)
        RecoveryTimes = @($normalized)
        DaysOfWeek = @($selectedDays)
        DayLabels = @($dayLabels)
        Items = @($items | Sort-Object RecoveryMinutes)
        KickMinutes = @($allKickMinutes | Sort-Object)
        Issues = @($issues)
    }
}

function Get-RecoveryPlanText ($plan) {
    $lines = New-Object 'System.Collections.Generic.List[string]'

    if ($plan.RecoveryTimes.Count -eq 0) {
        $lines.Add("回復時刻を追加すると、5時間前から再試行 kick を自動計算します。")
    } else {
        if ($plan.DayLabels.Count -gt 0) {
            $lines.Add("曜日: $($plan.DayLabels -join ' ')")
            $lines.Add("")
        }
        foreach ($item in $plan.Items) {
            $lines.Add("$($item.KickStartTime)-$($item.KickEndTime) kick /$($Script:RetryIntervalMinutes)m -> $($item.RecoveryTime)")
        }
    }

    if (-not $plan.IsValid) {
        if ($lines.Count -gt 0) { $lines.Add("") }
        foreach ($issue in $plan.Issues) { $lines.Add("警告: $issue") }
    } else {
        if ($lines.Count -gt 0) {
            $lines.Add("")
            $lines.Add("各目標時刻の5時間前から $($Script:RetryDurationMinutes) 分間、$($Script:RetryIntervalMinutes) 分おきに再試行します。")
            $lines.Add("回復時刻どうしは $($Script:MinimumSpacingLabel) 以上あいています。")
        }
    }

    return ($lines -join "`r`n")
}

function Set-ListBoxTimes ($listBox, $times) {
    $listBox.BeginUpdate()
    $listBox.Items.Clear()
    foreach ($time in (Get-NormalizedRecoveryTimes $times)) {
        [void]$listBox.Items.Add($time)
    }
    $listBox.EndUpdate()
}

function Get-ListBoxTimes ($listBox) {
    $times = @()
    foreach ($item in $listBox.Items) { $times += [string]$item }
    return @(Get-NormalizedRecoveryTimes $times)
}

# ── ログ ─────────────────────────────────────────────────────────────
function Add-LogLine ([string]$path, [string]$line) {
    $encoding = New-Object System.Text.UTF8Encoding($false)
    for ($attempt = 0; $attempt -lt 5; $attempt++) {
        try {
            $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            try {
                $writer = New-Object System.IO.StreamWriter($stream, $encoding)
                $writer.WriteLine($line)
                $writer.Flush()
            } finally {
                if ($writer) { $writer.Dispose() }
                $stream.Dispose()
            }
            return
        } catch [System.IO.IOException] {
            Start-Sleep -Milliseconds 100
        }
    }

    throw "ログファイルに書き込めません: $path"
}

function Write-Log ([string]$tool, [string]$message) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts [$tool] $message"
    Add-LogLine $Script:LogFile $line
    if ($Script:LogBox) {
        $Script:LogBox.AppendText("$line`r`n")
        $Script:LogBox.SelectionStart = $Script:LogBox.TextLength
        $Script:LogBox.ScrollToCaret()
    }
}

function Get-CommandStatus ([scriptblock]$action) {
    $global:LASTEXITCODE = $null
    try {
        & $action | Out-Null
        if ($null -ne $global:LASTEXITCODE) {
            if ($global:LASTEXITCODE -eq 0) { return "OK" }
            return "FAIL(exit=$global:LASTEXITCODE)"
        }
        return "OK"
    } catch {
        return "ERROR: $($_.Exception.Message)"
    }
}

# ── 実行コマンド構築 ─────────────────────────────────────────────────
function Get-CommandParts ([string]$tool) {
    switch ($tool) {
        "Claude" {
            return [PSCustomObject]@{
                FilePath = "claude"
                Arguments = @("-p", "--no-session-persistence", "--model", $Script:ClaudeModel, $Script:DefaultPrompt)
            }
        }
        "Codex" {
            return [PSCustomObject]@{
                FilePath = "codex"
                Arguments = @("exec", "--ephemeral", "--model", $Script:CodexModel, $Script:DefaultPrompt)
            }
        }
        default {
            throw "Unknown tool: $tool"
        }
    }
}

function Build-Command ([string]$tool, [bool]$useWSL) {
    $parts = Get-CommandParts $tool
    $quotedArgs = @()
    foreach ($arg in $parts.Arguments) {
        if ($arg -match '^[A-Za-z0-9._:/=-]+$') {
            $quotedArgs += $arg
        } else {
            $quotedArgs += (ConvertTo-BashSingleQuotedLiteral $arg)
        }
    }
    $inner = ($parts.FilePath + " " + ($quotedArgs -join " ")).Trim()

    if ($useWSL) { return "wsl bash -lic " + (ConvertTo-BashSingleQuotedLiteral $inner) }
    return $inner
}

# ── スケジュールタスクが呼び出すラッパースクリプト ────────────────────
function Create-WrapperScript ([string]$tool, [bool]$useWSL) {
    $ps1Path = Join-Path $Script:ConfigDir "kick-$($tool.ToLower()).ps1"
    $vbsPath = Join-Path $Script:ConfigDir "kick-$($tool.ToLower()).vbs"
    $command = Get-CommandParts $tool
    if ($useWSL) {
        $inner = Build-Command $tool $false
        $filePathLiteral = ConvertTo-PSSingleQuotedLiteral "wsl"
        $argumentsLiteral = ConvertTo-PowerShellArrayLiteral @("bash", "-lic", $inner)
    } else {
        $filePathLiteral = ConvertTo-PSSingleQuotedLiteral $command.FilePath
        $argumentsLiteral = ConvertTo-PowerShellArrayLiteral $command.Arguments
    }
    $logPathLiteral = ConvertTo-PSSingleQuotedLiteral $Script:LogFile

    $ps1Content = @"
`$ErrorActionPreference = 'Stop'
`$ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
`$global:LASTEXITCODE = `$null
`$filePath = $filePathLiteral
`$arguments = $argumentsLiteral
try {
    & `$filePath @arguments 2>&1 | Out-Null
    if (`$null -ne `$global:LASTEXITCODE) {
        if (`$global:LASTEXITCODE -eq 0) {
            `$status = 'OK'
        } else {
            `$status = "FAIL(exit=`$global:LASTEXITCODE)"
        }
    } else {
        `$status = 'OK'
    }
} catch {
    `$status = "ERROR: `$(`$_.Exception.Message)"
}
`$encoding = New-Object System.Text.UTF8Encoding(`$false)
for (`$attempt = 0; `$attempt -lt 5; `$attempt++) {
    try {
        `$stream = [System.IO.File]::Open($logPathLiteral, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        try {
            `$writer = New-Object System.IO.StreamWriter(`$stream, `$encoding)
            `$writer.WriteLine("`$ts [$tool] `$status")
            `$writer.Flush()
        } finally {
            if (`$writer) { `$writer.Dispose() }
            `$stream.Dispose()
        }
        break
    } catch [System.IO.IOException] {
        Start-Sleep -Milliseconds 100
    }
}
"@
    Set-Content -Path $ps1Path -Value $ps1Content -Encoding UTF8

    $vbsContent = "CreateObject(""WScript.Shell"").Run ""powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """"$ps1Path"""""", 0, False"
    Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII

    return $vbsPath
}

# ── タスクスケジューラ管理 ───────────────────────────────────────────
function Get-TaskName ([string]$tool) {
    return "$($Script:TaskPrefix)-$tool"
}

function Test-TaskExists ([string]$tool) {
    $name = Get-TaskName $tool
    $task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
    return ($null -ne $task)
}

function Register-KickTask ([string]$tool, $recoveryTimes, $daysOfWeek, [bool]$useWSL) {
    $plan = Get-RecoveryPlan $recoveryTimes $daysOfWeek
    if (-not $plan.IsValid) {
        throw ($plan.Issues -join " ")
    }

    $name = Get-TaskName $tool
    Unregister-KickTask $tool

    $wrapper = Create-WrapperScript $tool $useWSL
    $action  = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$wrapper`""

    $triggers = @()
    foreach ($kickMinute in $plan.KickMinutes) {
        $at = (Get-Date).Date.AddMinutes([int]$kickMinute)
        $triggers += New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek $plan.DaysOfWeek -At $at
    }

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5)

    Register-ScheduledTask `
        -TaskName $name `
        -Action $action `
        -Trigger $triggers `
        -Settings $settings `
        -Description "$Script:AppName - $tool recovery-target kick" `
        -Force | Out-Null

    $targets = @()
    foreach ($item in $plan.Items) {
        $targets += "$($item.KickStartTime)-$($item.KickEndTime)->$($item.RecoveryTime)"
    }
    $targetSummary = $targets -join ", "
    $daySummary = $plan.DayLabels -join ""
    Write-Log $tool "タスク登録 ($targetSummary, 曜日=$daySummary, WSL=$useWSL)"
}

function Unregister-KickTask ([string]$tool) {
    $name = Get-TaskName $tool
    if (Test-TaskExists $tool) {
        Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log $tool "タスク解除"
    }
}

function Invoke-KickNow ([string]$tool, [bool]$useWSL) {
    $preview = Build-Command $tool $useWSL
    $parts = Get-CommandParts $tool
    if ($useWSL) {
        $filePath = "wsl"
        $arguments = @("bash", "-lic", (Build-Command $tool $false))
    } else {
        $filePath = $parts.FilePath
        $arguments = $parts.Arguments
    }

    Write-Log $tool "手動キック: $preview"
    $status = Get-CommandStatus { & $filePath @arguments 2>&1 | Out-Null }
    Write-Log $tool $status
}

# ── GUI ──────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text            = $Script:AppName
$form.Size            = New-Object System.Drawing.Size(520, 840)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

$cfg = Load-Config

$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Text      = "回復していてほしい時刻と曜日を指定すると、その5時間前から再試行 kick を自動で登録します。"
$lblDesc.Location  = New-Object System.Drawing.Point(14, 10)
$lblDesc.Size      = New-Object System.Drawing.Size(480, 20)
$lblDesc.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$form.Controls.Add($lblDesc)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text      = "迷ったら平日の 11:00 / 16:10 / 21:20。各時刻の5時間前から2時間、30分おきに再試行します。"
$lblHint.Location  = New-Object System.Drawing.Point(14, 30)
$lblHint.Size      = New-Object System.Drawing.Size(480, 20)
$lblHint.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$form.Controls.Add($lblHint)

function New-ToolSection ([string]$tool, [int]$top, $parentForm, $cfgSection) {
    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text     = "$tool"
    $grp.Location = New-Object System.Drawing.Point(12, $top)
    $grp.Size     = New-Object System.Drawing.Size(480, 255)
    $parentForm.Controls.Add($grp)

    $lblTime = New-Object System.Windows.Forms.Label
    $lblTime.Text     = "回復時刻:"
    $lblTime.Location = New-Object System.Drawing.Point(14, 28)
    $lblTime.AutoSize = $true
    $grp.Controls.Add($lblTime)

    $timePicker = New-Object System.Windows.Forms.DateTimePicker
    $timePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $timePicker.CustomFormat = "HH:mm"
    $timePicker.ShowUpDown = $true
    $timePicker.Width = 80
    $timePicker.Location = New-Object System.Drawing.Point(82, 24)
    $timePicker.Value = (Get-Date).Date.AddHours(11)
    $grp.Controls.Add($timePicker)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text     = "追加"
    $btnAdd.Location = New-Object System.Drawing.Point(172, 22)
    $btnAdd.Size     = New-Object System.Drawing.Size(70, 28)
    $grp.Controls.Add($btnAdd)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text     = "削除"
    $btnRemove.Location = New-Object System.Drawing.Point(248, 22)
    $btnRemove.Size     = New-Object System.Drawing.Size(70, 28)
    $btnRemove.Enabled  = $false
    $grp.Controls.Add($btnRemove)

    $chkWSL = New-Object System.Windows.Forms.CheckBox
    $chkWSL.Text     = "WSL"
    $chkWSL.Location = New-Object System.Drawing.Point(390, 26)
    $chkWSL.AutoSize = $true
    $chkWSL.Checked  = $cfgSection.UseWSL
    $grp.Controls.Add($chkWSL)

    $lblDays = New-Object System.Windows.Forms.Label
    $lblDays.Text     = "曜日:"
    $lblDays.Location = New-Object System.Drawing.Point(14, 60)
    $lblDays.AutoSize = $true
    $grp.Controls.Add($lblDays)

    $dayCheckboxes = @()
    $dayLeft = 58
    foreach ($option in $Script:WeekdayOptions) {
        $chkDay = New-Object System.Windows.Forms.CheckBox
        $chkDay.Text = $option.Label
        $chkDay.Tag = $option.Name
        $chkDay.AutoSize = $true
        $chkDay.Location = New-Object System.Drawing.Point($dayLeft, 58)
        $chkDay.Checked = ($cfgSection.DaysOfWeek -contains $option.Name)
        $grp.Controls.Add($chkDay)
        $dayCheckboxes += $chkDay
        $dayLeft += 42
    }

    $lstTimes = New-Object System.Windows.Forms.ListBox
    $lstTimes.Location = New-Object System.Drawing.Point(14, 88)
    $lstTimes.Size     = New-Object System.Drawing.Size(110, 95)
    $grp.Controls.Add($lstTimes)
    Set-ListBoxTimes $lstTimes $cfgSection.RecoveryTimes

    $txtPreview = New-Object System.Windows.Forms.TextBox
    $txtPreview.Location   = New-Object System.Drawing.Point(136, 88)
    $txtPreview.Size       = New-Object System.Drawing.Size(330, 95)
    $txtPreview.Multiline  = $true
    $txtPreview.ReadOnly   = $true
    $txtPreview.ScrollBars = "Vertical"
    $txtPreview.Font       = New-Object System.Drawing.Font("Consolas", 8.5)
    $grp.Controls.Add($txtPreview)

    $lblSt = New-Object System.Windows.Forms.Label
    $lblSt.Location = New-Object System.Drawing.Point(14, 194)
    $lblSt.Size     = New-Object System.Drawing.Size(310, 18)
    $grp.Controls.Add($lblSt)

    $btnOn = New-Object System.Windows.Forms.Button
    $btnOn.Text     = "保存して有効化"
    $btnOn.Location = New-Object System.Drawing.Point(14, 218)
    $btnOn.Size     = New-Object System.Drawing.Size(120, 28)
    $grp.Controls.Add($btnOn)

    $btnOff = New-Object System.Windows.Forms.Button
    $btnOff.Text     = "無効化"
    $btnOff.Location = New-Object System.Drawing.Point(142, 218)
    $btnOff.Size     = New-Object System.Drawing.Size(100, 28)
    $grp.Controls.Add($btnOff)

    $btnNow = New-Object System.Windows.Forms.Button
    $btnNow.Text     = "今すぐ kick"
    $btnNow.Location = New-Object System.Drawing.Point(250, 218)
    $btnNow.Size     = New-Object System.Drawing.Size(100, 28)
    $grp.Controls.Add($btnNow)

    $lblCmd = New-Object System.Windows.Forms.Label
    $lblCmd.Location  = New-Object System.Drawing.Point(356, 218)
    $lblCmd.Size      = New-Object System.Drawing.Size(110, 28)
    $lblCmd.ForeColor = [System.Drawing.Color]::Gray
    $lblCmd.Font      = New-Object System.Drawing.Font("Consolas", 7.0)
    $lblCmd.TextAlign = "MiddleLeft"
    $grp.Controls.Add($lblCmd)

    $getSelectedDays = {
        $selected = @()
        foreach ($chkDay in $dayCheckboxes) {
            if ($chkDay.Checked) { $selected += [string]$chkDay.Tag }
        }
        return @(Get-NormalizedDaysOfWeek $selected)
    }.GetNewClosure()

    $updatePreview = {
        try {
            $plan = Get-RecoveryPlan (Get-ListBoxTimes $lstTimes) (& $getSelectedDays)
            $txtPreview.Text = Get-RecoveryPlanText $plan
            if ($plan.IsValid) {
                $txtPreview.ForeColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
            } else {
                $txtPreview.ForeColor = [System.Drawing.Color]::Firebrick
            }
            return $plan
        } catch {
            $txtPreview.Text = "警告: プレビュー更新に失敗しました。`r`n$($_.Exception.Message)"
            $txtPreview.ForeColor = [System.Drawing.Color]::Firebrick
            return [PSCustomObject]@{
                IsValid = $false
                RecoveryTimes = @()
                DaysOfWeek = @()
                DayLabels = @()
                Items = @()
                KickMinutes = @()
                Issues = @($_.Exception.Message)
            }
        }
    }.GetNewClosure()

    $updateStatus = {
        $plan = & $updatePreview
        $exists = Test-TaskExists $tool
        if ($exists) {
            $lblSt.Text      = "回復予約: ON"
            $lblSt.ForeColor = [System.Drawing.Color]::Green
            $btnOff.Enabled  = $true
        } else {
            $lblSt.Text      = "回復予約: OFF"
            $lblSt.ForeColor = [System.Drawing.Color]::Gray
            $btnOff.Enabled  = $false
        }

        if (-not $plan.IsValid) {
            $lblSt.Text = "$($lblSt.Text) / 設定を見直してください"
            $lblSt.ForeColor = [System.Drawing.Color]::Firebrick
        }

        $lblCmd.Text = Build-Command $tool $chkWSL.Checked
    }.GetNewClosure()

    & $updateStatus

    $btnAdd.Add_Click({
        $time = $timePicker.Value.ToString("HH:mm")
        $current = @(Get-ListBoxTimes $lstTimes)
        if ($current -contains $time) {
            [System.Windows.Forms.MessageBox]::Show("同じ回復時刻は追加できません。", $tool, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $updated = @()
        $updated += $current
        $updated += $time
        Set-ListBoxTimes $lstTimes $updated
        & $updateStatus
    }.GetNewClosure())

    $btnRemove.Add_Click({
        if ($lstTimes.SelectedItem) {
            $items = @(Get-ListBoxTimes $lstTimes | Where-Object { $_ -ne [string]$lstTimes.SelectedItem })
            Set-ListBoxTimes $lstTimes $items
            & $updateStatus
        }
    }.GetNewClosure())

    $lstTimes.Add_SelectedIndexChanged({
        $btnRemove.Enabled = ($null -ne $lstTimes.SelectedItem)
    }.GetNewClosure())

    foreach ($chkDay in $dayCheckboxes) {
        $chkDay.Add_CheckedChanged({ & $updateStatus }.GetNewClosure())
    }

    $chkWSL.Add_CheckedChanged({ & $updateStatus }.GetNewClosure())

    $btnOn.Add_Click({
        $times = Get-ListBoxTimes $lstTimes
        $days = & $getSelectedDays
        $plan = Get-RecoveryPlan $times $days
        if (-not $plan.IsValid) {
            [System.Windows.Forms.MessageBox]::Show(($plan.Issues -join "`r`n"), "$tool の設定を修正してください", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        try {
            Register-KickTask $tool $times $days $chkWSL.Checked
            $config = Load-Config
            $config.$tool.RecoveryTimes = @($plan.RecoveryTimes)
            $config.$tool.DaysOfWeek = @($plan.DaysOfWeek)
            $config.$tool.UseWSL = $chkWSL.Checked
            Save-Config $config
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "$tool の有効化に失敗しました", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
        & $updateStatus
    }.GetNewClosure())

    $btnOff.Add_Click({
        Unregister-KickTask $tool
        & $updateStatus
    }.GetNewClosure())

    $btnNow.Add_Click({
        Invoke-KickNow $tool $chkWSL.Checked
    }.GetNewClosure())
}

New-ToolSection "Claude" 60  $form $cfg.Claude
New-ToolSection "Codex"  335 $form $cfg.Codex

$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text     = "ログ"
$grpLog.Location = New-Object System.Drawing.Point(12, 610)
$grpLog.Size     = New-Object System.Drawing.Size(480, 160)
$form.Controls.Add($grpLog)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location   = New-Object System.Drawing.Point(10, 22)
$logBox.Size       = New-Object System.Drawing.Size(458, 100)
$logBox.Multiline  = $true
$logBox.ReadOnly   = $true
$logBox.ScrollBars = "Vertical"
$logBox.Font       = New-Object System.Drawing.Font("Consolas", 8.5)
$logBox.BackColor  = [System.Drawing.Color]::FromArgb(30, 30, 30)
$logBox.ForeColor  = [System.Drawing.Color]::FromArgb(200, 220, 200)
$grpLog.Controls.Add($logBox)
$Script:LogBox = $logBox

if (Test-Path $Script:LogFile) {
    $recent = Get-Content $Script:LogFile -Tail 30 -ErrorAction SilentlyContinue
    if ($recent) { $logBox.Text = ($recent -join "`r`n") + "`r`n" }
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
}

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text     = "ログ消去"
$btnClear.Location = New-Object System.Drawing.Point(10, 128)
$btnClear.Size     = New-Object System.Drawing.Size(100, 26)
$grpLog.Controls.Add($btnClear)
$btnClear.Add_Click({
    $logBox.Clear()
    if (Test-Path $Script:LogFile) { Remove-Item $Script:LogFile -Force }
})

$btnTaskSch = New-Object System.Windows.Forms.Button
$btnTaskSch.Text     = "タスクスケジューラ"
$btnTaskSch.Location = New-Object System.Drawing.Point(118, 128)
$btnTaskSch.Size     = New-Object System.Drawing.Size(140, 26)
$grpLog.Controls.Add($btnTaskSch)
$btnTaskSch.Add_Click({
    Start-Process "taskschd.msc"
})

$btnGitHub = New-Object System.Windows.Forms.Button
$btnGitHub.Text     = "GitHub"
$btnGitHub.Location = New-Object System.Drawing.Point(266, 128)
$btnGitHub.Size     = New-Object System.Drawing.Size(80, 26)
$grpLog.Controls.Add($btnGitHub)
$btnGitHub.Add_Click({
    Start-Process "https://github.com/t-suzuki/agent-recovery-scheduler"
})

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text      = "平日デフォルト / 5時間前から再試行 / Claude: $($Script:ClaudeModel) / Codex: $($Script:CodexModel)"
$lblInfo.Location  = New-Object System.Drawing.Point(12, 782)
$lblInfo.Size      = New-Object System.Drawing.Size(480, 20)
$lblInfo.ForeColor = [System.Drawing.Color]::Gray
$lblInfo.Font      = New-Object System.Drawing.Font("Consolas", 7.5)
$form.Controls.Add($lblInfo)

[void]$form.ShowDialog()
