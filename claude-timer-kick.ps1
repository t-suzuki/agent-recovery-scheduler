<#
.SYNOPSIS
    Agent Timer Kicker - Claude Code / Codex CLI の5時間ローリングウィンドウを自動キックします。
.DESCRIPTION
    Windows タスクスケジューラに軽量プロンプトを定期送信するタスクを登録し、
    使用量ウィンドウのタイマーを早めにリフレッシュさせます。
.NOTES
    要件: Windows 10+, PowerShell 5.1+, claude-code and/or codex CLI 認証済み。
    実行: powershell -ExecutionPolicy Bypass -File claude-timer-kick.ps1
    または launch.bat をダブルクリック。
#>

# ── 設定 ─────────────────────────────────────────────────────────────
$Script:AppName        = "Agent Timer Kicker"
$Script:ScriptDir      = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:ConfigDir      = Join-Path $Script:ScriptDir "data"
$Script:ConfigFile     = Join-Path $Script:ConfigDir "config.json"
$Script:LogFile        = Join-Path $Script:ConfigDir "kick.log"
$Script:TaskPrefix     = "AgentTimerKick"
$Script:DefaultPrompt  = "just say hi and nothing else."
$Script:ClaudeModel    = "haiku"
$Script:CodexModel     = "gpt-5-codex-mini"

# ── データディレクトリ確保 ───────────────────────────────────────────
if (-not (Test-Path $Script:ConfigDir)) { New-Item -ItemType Directory -Path $Script:ConfigDir -Force | Out-Null }

# ── 設定ヘルパー ─────────────────────────────────────────────────────
function Load-Config {
    if (Test-Path $Script:ConfigFile) {
        try { return Get-Content $Script:ConfigFile -Raw | ConvertFrom-Json }
        catch { }
    }
    return [PSCustomObject]@{
        Claude = [PSCustomObject]@{ IntervalMin = 30; UseWSL = $false }
        Codex  = [PSCustomObject]@{ IntervalMin = 30; UseWSL = $false }
    }
}
function Save-Config ($cfg) {
    $cfg | ConvertTo-Json -Depth 4 | Set-Content $Script:ConfigFile -Encoding UTF8
}

# ── ログ ─────────────────────────────────────────────────────────────
function Write-Log ([string]$tool, [string]$message) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts [$tool] $message"
    Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8
    if ($Script:LogBox) {
        $Script:LogBox.AppendText("$line`r`n")
        $Script:LogBox.SelectionStart = $Script:LogBox.TextLength
        $Script:LogBox.ScrollToCaret()
    }
}

# ── 実行結果の判定 ───────────────────────────────────────────────────
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
function Build-Command ([string]$tool, [bool]$useWSL) {
    $prompt = $Script:DefaultPrompt
    switch ($tool) {
        "Claude" {
            $inner = "claude -p --no-session-persistence --model $($Script:ClaudeModel) `"$prompt`""
        }
        "Codex" {
            $inner = "codex exec --ephemeral --model $($Script:CodexModel) `"$prompt`""
        }
    }
    if ($useWSL) { return "wsl bash -lic '$inner'" }
    return $inner
}

# ── スケジュールタスクが呼び出すラッパースクリプト ────────────────────
function Create-WrapperScript ([string]$tool, [bool]$useWSL) {
    $ps1Path = Join-Path $Script:ConfigDir "kick-$($tool.ToLower()).ps1"
    $vbsPath = Join-Path $Script:ConfigDir "kick-$($tool.ToLower()).vbs"
    $cmd = Build-Command $tool $useWSL
    $logPath = $Script:LogFile

    # PS1 本体
    $ps1Content = @"
`$ErrorActionPreference = 'Stop'
`$ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
`$global:LASTEXITCODE = `$null
try {
    Invoke-Expression '$cmd' | Out-Null
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
Add-Content -Path '$logPath' -Value "`$ts [$tool] `$status" -Encoding UTF8
"@
    Set-Content -Path $ps1Path -Value $ps1Content -Encoding UTF8

    # VBS ランチャー（ウィンドウ完全非表示）
    $vbsContent = "CreateObject(""WScript.Shell"").Run ""powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """"$ps1Path"""""", 0, False"
    Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII

    return $vbsPath
}

# ── タスクスケジューラ管理 ───────────────────────────────────────────
function Get-TaskName ([string]$tool) { return "$($Script:TaskPrefix)-$tool" }

function Test-TaskExists ([string]$tool) {
    $name = Get-TaskName $tool
    $t = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
    return ($null -ne $t)
}

function Register-KickTask ([string]$tool, [int]$intervalMin, [bool]$useWSL) {
    $name = Get-TaskName $tool
    Unregister-KickTask $tool

    $wrapper = Create-WrapperScript $tool $useWSL
    $action  = New-ScheduledTaskAction `
        -Execute "wscript.exe" `
        -Argument "`"$wrapper`""

    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date)

    $rep = New-CimInstance -CimClass (
        Get-CimClass -Namespace "Root/Microsoft/Windows/TaskScheduler" -ClassName "MSFT_TaskRepetitionPattern"
    ) -ClientOnly
    $rep.Interval  = "PT$($intervalMin)M"
    $rep.Duration  = ""
    $rep.StopAtDurationEnd = $false
    $trigger.Repetition = $rep

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5)

    Register-ScheduledTask `
        -TaskName $name `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Description "$Script:AppName - $tool CLI 定期キック" `
        -Force | Out-Null

    Write-Log $tool "タスク登録 (${intervalMin}分間隔, WSL=$useWSL)"
}

function Unregister-KickTask ([string]$tool) {
    $name = Get-TaskName $tool
    if (Test-TaskExists $tool) {
        Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log $tool "タスク解除"
    }
}

function Invoke-KickNow ([string]$tool, [bool]$useWSL) {
    $cmd = Build-Command $tool $useWSL
    Write-Log $tool "手動キック: $cmd"
    $status = Get-CommandStatus { Invoke-Expression $cmd }
    Write-Log $tool $status
}

# ── GUI ──────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text            = $Script:AppName
$form.Size            = New-Object System.Drawing.Size(520, 620)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

$cfg = Load-Config

# ── 説明ラベル ──
$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Text      = "5時間ローリングウィンドウを定期的にキックし、レート制限のリフレッシュを早めます。"
$lblDesc.Location  = New-Object System.Drawing.Point(14, 10)
$lblDesc.Size      = New-Object System.Drawing.Size(480, 20)
$lblDesc.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$form.Controls.Add($lblDesc)

# ── ヘルパー: ツールセクション生成 ───────────────────────────────────
function New-ToolSection ([string]$tool, [int]$top, $parentForm, $cfgSection) {
    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text     = "$tool"
    $grp.Location = New-Object System.Drawing.Point(12, $top)
    $grp.Size     = New-Object System.Drawing.Size(480, 140)
    $parentForm.Controls.Add($grp)

    # ── 行1: 間隔 + WSL ──
    $lblInt = New-Object System.Windows.Forms.Label
    $lblInt.Text     = "間隔 (分):"
    $lblInt.Location = New-Object System.Drawing.Point(14, 28)
    $lblInt.AutoSize = $true
    $grp.Controls.Add($lblInt)

    $numInt = New-Object System.Windows.Forms.NumericUpDown
    $numInt.Location = New-Object System.Drawing.Point(100, 25)
    $numInt.Size     = New-Object System.Drawing.Size(70, 25)
    $numInt.Minimum  = 1
    $numInt.Maximum  = 600
    $numInt.Value    = $cfgSection.IntervalMin
    $grp.Controls.Add($numInt)

    $lblMin = New-Object System.Windows.Forms.Label
    $lblMin.Text     = "分"
    $lblMin.Location = New-Object System.Drawing.Point(174, 28)
    $lblMin.AutoSize = $true
    $grp.Controls.Add($lblMin)

    $chkWSL = New-Object System.Windows.Forms.CheckBox
    $chkWSL.Text     = "WSL"
    $chkWSL.Location = New-Object System.Drawing.Point(250, 26)
    $chkWSL.AutoSize = $true
    $chkWSL.Checked  = $cfgSection.UseWSL
    $grp.Controls.Add($chkWSL)

    # ── 行2: ステータス ──
    $lblSt = New-Object System.Windows.Forms.Label
    $lblSt.Location = New-Object System.Drawing.Point(14, 62)
    $lblSt.AutoSize = $true
    $grp.Controls.Add($lblSt)

    # ── 行3: ボタン ──
    $btnOn = New-Object System.Windows.Forms.Button
    $btnOn.Text     = [char]0x25B6 + " 有効化"
    $btnOn.Location = New-Object System.Drawing.Point(14, 95)
    $btnOn.Size     = New-Object System.Drawing.Size(100, 30)
    $grp.Controls.Add($btnOn)

    $btnOff = New-Object System.Windows.Forms.Button
    $btnOff.Text     = [char]0x23F9 + " 無効化"
    $btnOff.Location = New-Object System.Drawing.Point(122, 95)
    $btnOff.Size     = New-Object System.Drawing.Size(100, 30)
    $grp.Controls.Add($btnOff)

    $btnNow = New-Object System.Windows.Forms.Button
    $btnNow.Text     = "今すぐ実行"
    $btnNow.Location = New-Object System.Drawing.Point(230, 95)
    $btnNow.Size     = New-Object System.Drawing.Size(100, 30)
    $grp.Controls.Add($btnNow)

    # ── コマンドプレビュー ──
    $lblCmd = New-Object System.Windows.Forms.Label
    $lblCmd.Location  = New-Object System.Drawing.Point(338, 95)
    $lblCmd.Size      = New-Object System.Drawing.Size(132, 30)
    $lblCmd.ForeColor = [System.Drawing.Color]::Gray
    $lblCmd.Font      = New-Object System.Drawing.Font("Consolas", 7.5)
    $lblCmd.TextAlign = "MiddleLeft"
    $grp.Controls.Add($lblCmd)

    # ── ステータス更新 ──
    $updateStatus = {
        $exists = Test-TaskExists $tool
        if ($exists) {
            $lblSt.Text      = "定期キック:  ON"
            $lblSt.ForeColor = [System.Drawing.Color]::Green
            $btnOn.Enabled   = $false
            $btnOff.Enabled  = $true
        } else {
            $lblSt.Text      = "定期キック:  OFF"
            $lblSt.ForeColor = [System.Drawing.Color]::Gray
            $btnOn.Enabled   = $true
            $btnOff.Enabled  = $false
        }
        $lblCmd.Text = Build-Command $tool $chkWSL.Checked
    }.GetNewClosure()

    & $updateStatus

    $chkWSL.Add_CheckedChanged({ & $updateStatus }.GetNewClosure())

    # ── ボタンハンドラー ──
    $btnOn.Add_Click({
        $intVal  = [int]$numInt.Value
        $wslVal  = $chkWSL.Checked
        Register-KickTask $tool $intVal $wslVal
        $c = Load-Config
        $c.$tool.IntervalMin = $intVal
        $c.$tool.UseWSL      = $wslVal
        Save-Config $c
        & $updateStatus
    }.GetNewClosure())

    $btnOff.Add_Click({
        Unregister-KickTask $tool
        & $updateStatus
    }.GetNewClosure())

    $btnNow.Add_Click({
        $wslVal = $chkWSL.Checked
        Invoke-KickNow $tool $wslVal
    }.GetNewClosure())
}

# ── 2つのセクション作成 ──────────────────────────────────────────────
New-ToolSection "Claude" 36  $form $cfg.Claude
New-ToolSection "Codex"  186 $form $cfg.Codex

# ── ログ領域 ─────────────────────────────────────────────────────────
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text     = "ログ"
$grpLog.Location = New-Object System.Drawing.Point(12, 336)
$grpLog.Size     = New-Object System.Drawing.Size(480, 190)
$form.Controls.Add($grpLog)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location   = New-Object System.Drawing.Point(10, 22)
$logBox.Size       = New-Object System.Drawing.Size(458, 120)
$logBox.Multiline  = $true
$logBox.ReadOnly   = $true
$logBox.ScrollBars = "Vertical"
$logBox.Font       = New-Object System.Drawing.Font("Consolas", 8.5)
$logBox.BackColor  = [System.Drawing.Color]::FromArgb(30, 30, 30)
$logBox.ForeColor  = [System.Drawing.Color]::FromArgb(200, 220, 200)
$grpLog.Controls.Add($logBox)
$Script:LogBox = $logBox

# 直近のログ読み込み
if (Test-Path $Script:LogFile) {
    $recent = Get-Content $Script:LogFile -Tail 30 -ErrorAction SilentlyContinue
    if ($recent) { $logBox.Text = ($recent -join "`r`n") + "`r`n" }
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
}

# ── ログクリアボタン ──
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text     = "ログ消去"
$btnClear.Location = New-Object System.Drawing.Point(10, 150)
$btnClear.Size     = New-Object System.Drawing.Size(100, 28)
$grpLog.Controls.Add($btnClear)
$btnClear.Add_Click({
    $logBox.Clear()
    if (Test-Path $Script:LogFile) { Remove-Item $Script:LogFile -Force }
})

# ── タスクスケジューラを開くボタン ──
$btnTaskSch = New-Object System.Windows.Forms.Button
$btnTaskSch.Text     = "タスクスケジューラ"
$btnTaskSch.Location = New-Object System.Drawing.Point(118, 150)
$btnTaskSch.Size     = New-Object System.Drawing.Size(140, 28)
$grpLog.Controls.Add($btnTaskSch)
$btnTaskSch.Add_Click({
    Start-Process "taskschd.msc"
})

# ── GitHubボタン ──
$btnGitHub = New-Object System.Windows.Forms.Button
$btnGitHub.Text     = "GitHub"
$btnGitHub.Location = New-Object System.Drawing.Point(266, 150)
$btnGitHub.Size     = New-Object System.Drawing.Size(80, 28)
$grpLog.Controls.Add($btnGitHub)
$btnGitHub.Add_Click({
    Start-Process "https://github.com/t-suzuki/agent-timer-kicker"
})

# ── 情報ラベル ──
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text      = "Claude: $($Script:ClaudeModel)  Codex: $($Script:CodexModel)  プロンプト: `"$($Script:DefaultPrompt)`""
$lblInfo.Location  = New-Object System.Drawing.Point(12, 540)
$lblInfo.Size      = New-Object System.Drawing.Size(480, 20)
$lblInfo.ForeColor = [System.Drawing.Color]::Gray
$lblInfo.Font      = New-Object System.Drawing.Font("Consolas", 7.5)
$form.Controls.Add($lblInfo)

# ── 表示 ─────────────────────────────────────────────────────────────
[void]$form.ShowDialog()
