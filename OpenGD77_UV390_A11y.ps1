param(
    [switch]$SelfTest
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:BackendExe = Join-Path $script:BaseDir "dist\\OpenGD77_UV390_BackendCLI.exe"
$script:ProcessRunning = $false
$script:CurrentProcess = $null
$script:BackendExitHandled = $true
$script:LogQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$script:EventQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$script:CurrentOperationType = ""
$script:SuccessChirpWaveBytes = $null
$script:StartupSoundPath = Join-Path $script:BaseDir "assets\\startup_radio_shortwave_cc0.wav"
$script:AppVersion = "2026.03.01.7"
$script:ProgramUpdateConfigPath = Join-Path $script:BaseDir "program_update_config.json"
$script:DefaultProgramUpdateManifestUrl = "https://kazpar.pl/opengd77-updater/latest.json"
$script:UiLanguage = "pl"
$script:SkipCloseConfirmation = $false
$script:StartupProgramUpdateTimer = $null
$script:LogDir = Join-Path $script:BaseDir "logs"
if (-not (Test-Path $script:LogDir)) { [void](New-Item -ItemType Directory -Path $script:LogDir) }
$script:SessionStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:DebugLog = Join-Path $script:LogDir ("OpenGD77_A11y_" + $script:SessionStamp + ".log")

function Write-DebugLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    try {
        $line = ("[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"), $Level, $Message)
        Add-Content -Path $script:DebugLog -Value $line -Encoding UTF8
    } catch {
        # swallow logging failures
    }
}

trap {
    Write-DebugLog -Level "ERROR" -Message ("UNHANDLED: " + $_.Exception.Message + " | " + $_.ScriptStackTrace)
    throw
}

Write-DebugLog "A11y launcher start. BaseDir=$script:BaseDir"

function Normalize-UiLanguage {
    param([string]$Language)
    if ([string]::IsNullOrWhiteSpace($Language)) { return "pl" }
    $v = $Language.Trim().ToLowerInvariant()
    if ($v -in @("pl", "en")) { return $v }
    return "pl"
}

$script:I18n = @{
    pl = @{
        form_title = "OpenGD77 Updater - TYT UV390 Plus 10W (A11y)"
        title = "OpenGD77 Updater dla TYT UV390 Plus 10W"
        options = "Opcje"
        check_only = "Tylko sprawdź wersję (bez flashowania)"
        auto_driver = "Po braku DFU spróbuj automatycznej instalacji sterownika WinUSB"
        timeout = "Timeout DFU (sekundy):"
        ui_language = "Język interfejsu:"
        btn_check = "&Sprawdź wersję"
        btn_start = "&Start aktualizacji"
        btn_program_update = "&Aktualizuj program"
        btn_close = "&Zamknij"
        hints = "Skróty: Alt+S Sprawdź, Alt+T Start, Alt+A Aktualizuj program, Alt+L Log, Alt+D Timeout, Alt+F4 Zamknij, F1 Pomoc."
        status_group = "Status"
        status_label = "Bieżący status:"
        status_ready = "Gotowe do uruchomienia"
        last_message_label = "Ostatni komunikat:"
        status_no_messages = "Brak komunikatów."
        log_group = "Log"
        status_busy = "Trwa operacja..."
        status_success = "Zakończono pomyślnie"
        status_error_code_fmt = "Błąd (kod {0})"
        status_checking_program_update = "Sprawdzam aktualizację programu..."
        status_program_up_to_date = "Program jest aktualny"
        status_program_update_canceled = "Anulowano aktualizację programu"
        status_downloading_program_update = "Pobieram aktualizację programu..."
        status_restart_program_update = "Restart do aktualizacji programu..."
        status_program_update_error = "Błąd aktualizacji programu"
        cap_program_update = "Aktualizacja programu"
        cap_error = "Błąd"
        cap_dfu = "DFU"
        cap_confirmation = "Potwierdzenie"
        cap_help = "Pomoc"
        cap_update_in_progress = "Aktualizacja w toku"
        cap_confirm_close = "Potwierdzenie zamknięcia"
        msg_latest_program_fmt = "Masz najnowszą wersję programu: {0}"
        msg_program_new_auto_fmt = "Wykryto nową wersję programu: {0}`nObecna wersja: {1}`n`nRozpoczynam aktualizację programu.`nProgram zamknie się i uruchomi ponownie."
        msg_program_new_question_fmt = "Dostępna nowa wersja programu: {0}`nObecna wersja: {1}`n`nPobrać i zainstalować teraz?`nProgram zamknie się i uruchomi ponownie."
        msg_program_update_started = "Aktualizacja programu rozpoczęta.`nProgram zamknie się i uruchomi ponownie."
        msg_program_update_error_fmt = "Błąd aktualizacji programu:`n{0}"
        msg_backend_missing_fmt = "Brak backendu: {0}"
        msg_dfu_instructions = "Ustaw radio w DFU:`n1) Wyłącz radio`n2) Przytrzymaj PTT + S1`n3) Włącz radio (czarny ekran)`n4) Podłącz USB`n`nKliknij OK, gdy jesteś gotowy."
        msg_start_update_now = "Rozpocząć aktualizację teraz?"
        msg_help = "Sterowanie:`n- Alt+S: Sprawdź wersję`n- Alt+T: Start aktualizacji`n- Alt+A: Aktualizuj program`n- Alt+L: Fokus log`n- Alt+D: Fokus timeout`n- Alt+F4: Zamknij"
        msg_cannot_close_during_update = "Nie można zamknąć programu podczas aktualizacji."
        msg_confirm_close = "Czy na pewno chcesz zamknąć program?"
        log_program_local_ver_fmt = "Program version (lokalna): {0}"
        log_program_remote_ver_fmt = "Program version (zdalna): {0}"
        log_manifest_url_fmt = "Manifest URL: {0}"
        log_auto_check_up_to_date = "Auto-check: program jest aktualny."
        log_auto_check_found_new = "Auto-check: wykryto nową wersję, start aktualizacji programu."
        log_download_pkg_fmt = "Pobieranie paczki: {0}"
        log_downloaded_fmt = "Pobrano: {0}"
        log_sha_ok = "SHA256 paczki: OK"
        log_manifest_no_sha = "Manifest bez SHA256 - pomijam weryfikacje hash."
        log_launch_helper_fmt = "Uruchomiono helper aktualizacji: {0}"
        log_autocheck_error_fmt = "Auto-check update błąd: {0}"
        log_language_changed_fmt = "Zmieniono język interfejsu na: {0}"
        lang_name_pl = "Polski"
        lang_name_en = "Angielski"
        error_empty_version = "Pusty numer wersji."
        error_parse_version_fmt = "Nie można sparsować wersji: {0}"
        error_manifest_no_version = "Manifest nie zawiera pola 'version'."
        error_manifest_no_package_url = "Manifest nie zawiera pola 'package_url'."
        error_sha_mismatch_fmt = "SHA256 nie pasuje. expected={0} actual={1}"
    }
    en = @{
        form_title = "OpenGD77 Updater - TYT UV390 Plus 10W (A11y)"
        title = "OpenGD77 Updater for TYT UV390 Plus 10W"
        options = "Options"
        check_only = "Check version only (no flashing)"
        auto_driver = "If DFU is missing, try automatic WinUSB driver install"
        timeout = "DFU timeout (seconds):"
        ui_language = "Interface language:"
        btn_check = "&Check version"
        btn_start = "&Start update"
        btn_program_update = "&Update app"
        btn_close = "&Close"
        hints = "Shortcuts: Alt+S Check, Alt+T Start, Alt+A Update app, Alt+L Log, Alt+D Timeout, Alt+F4 Close, F1 Help."
        status_group = "Status"
        status_label = "Current status:"
        status_ready = "Ready"
        last_message_label = "Last message:"
        status_no_messages = "No messages."
        log_group = "Log"
        status_busy = "Operation in progress..."
        status_success = "Completed successfully"
        status_error_code_fmt = "Error (code {0})"
        status_checking_program_update = "Checking app update..."
        status_program_up_to_date = "App is up to date"
        status_program_update_canceled = "App update canceled"
        status_downloading_program_update = "Downloading app update..."
        status_restart_program_update = "Restarting for app update..."
        status_program_update_error = "App update error"
        cap_program_update = "App update"
        cap_error = "Error"
        cap_dfu = "DFU"
        cap_confirmation = "Confirmation"
        cap_help = "Help"
        cap_update_in_progress = "Update in progress"
        cap_confirm_close = "Close confirmation"
        msg_latest_program_fmt = "You already have the latest app version: {0}"
        msg_program_new_auto_fmt = "A new app version was detected: {0}`nCurrent version: {1}`n`nStarting app update.`nThe app will close and restart."
        msg_program_new_question_fmt = "A new app version is available: {0}`nCurrent version: {1}`n`nDownload and install now?`nThe app will close and restart."
        msg_program_update_started = "App update started.`nThe app will close and restart."
        msg_program_update_error_fmt = "App update error:`n{0}"
        msg_backend_missing_fmt = "Backend not found: {0}"
        msg_dfu_instructions = "Put the radio into DFU:`n1) Turn radio off`n2) Hold PTT + S1`n3) Turn radio on (black screen)`n4) Connect USB`n`nClick OK when ready."
        msg_start_update_now = "Start update now?"
        msg_help = "Controls:`n- Alt+S: Check version`n- Alt+T: Start update`n- Alt+A: Update app`n- Alt+L: Log focus`n- Alt+D: Timeout focus`n- Alt+F4: Close"
        msg_cannot_close_during_update = "You cannot close the app during update."
        msg_confirm_close = "Are you sure you want to close the app?"
        log_program_local_ver_fmt = "Program version (local): {0}"
        log_program_remote_ver_fmt = "Program version (remote): {0}"
        log_manifest_url_fmt = "Manifest URL: {0}"
        log_auto_check_up_to_date = "Auto-check: app is up to date."
        log_auto_check_found_new = "Auto-check: new version detected, starting app update."
        log_download_pkg_fmt = "Downloading package: {0}"
        log_downloaded_fmt = "Downloaded: {0}"
        log_sha_ok = "Package SHA256: OK"
        log_manifest_no_sha = "Manifest has no SHA256 - hash verification skipped."
        log_launch_helper_fmt = "Started update helper: {0}"
        log_autocheck_error_fmt = "Auto-check update error: {0}"
        log_language_changed_fmt = "Interface language changed to: {0}"
        lang_name_pl = "Polish"
        lang_name_en = "English"
        error_empty_version = "Empty version string."
        error_parse_version_fmt = "Cannot parse version: {0}"
        error_manifest_no_version = "Manifest does not contain field 'version'."
        error_manifest_no_package_url = "Manifest does not contain field 'package_url'."
        error_sha_mismatch_fmt = "SHA256 mismatch. expected={0} actual={1}"
    }
}

function T {
    param([string]$Key)
    $lang = Normalize-UiLanguage $script:UiLanguage
    if ($script:I18n.ContainsKey($lang) -and $script:I18n[$lang].ContainsKey($Key)) {
        return [string]$script:I18n[$lang][$Key]
    }
    if ($script:I18n["pl"].ContainsKey($Key)) {
        return [string]$script:I18n["pl"][$Key]
    }
    return $Key
}

function TF {
    param(
        [string]$Key,
        [Parameter(ValueFromRemainingArguments = $true)]
        [Object[]]$Args
    )
    return [string]::Format((T $Key), $Args)
}

function Get-LanguageDisplayName {
    param([string]$LangCode)
    $lang = Normalize-UiLanguage $LangCode
    if ($lang -eq "en") { return (T "lang_name_en") }
    return (T "lang_name_pl")
}

function Append-Log {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    Write-DebugLog -Message ("UI_LOG: " + $Text)
    $script:txtLog.AppendText($Text + [Environment]::NewLine)
    $script:txtLog.SelectionStart = $script:txtLog.TextLength
    $script:txtLog.ScrollToCaret()
    $script:txtLastMessage.Text = $Text
}

function Set-Status {
    param([string]$Text)
    Write-DebugLog -Message ("STATUS: " + $Text)
    $script:txtStatus.Text = $Text
    $script:txtLastMessage.Text = $Text
}

function Set-Busy {
    param([bool]$Busy, [string]$Status)
    $script:ProcessRunning = $Busy
    Set-Status $Status
    $script:btnStart.Enabled = -not $Busy
    $script:btnCheck.Enabled = -not $Busy
    $script:btnProgramUpdate.Enabled = -not $Busy
    $script:btnClose.Enabled = -not $Busy
    $script:chkCheckOnly.Enabled = -not $Busy
    $script:chkAutoDriver.Enabled = -not $Busy
    $script:numTimeout.Enabled = -not $Busy
}

function Get-ProgramUpdateConfig {
    $cfg = @{
        app_version = $script:AppVersion
        manifest_url = $script:DefaultProgramUpdateManifestUrl
        auto_check_on_start = $true
        ui_language = "pl"
    }

    if (-not (Test-Path $script:ProgramUpdateConfigPath)) {
        return $cfg
    }

    try {
        $raw = Get-Content -Raw -Path $script:ProgramUpdateConfigPath -Encoding UTF8
        $fileCfg = $raw | ConvertFrom-Json
        if ($fileCfg.app_version) { $cfg.app_version = [string]$fileCfg.app_version }
        if ($fileCfg.manifest_url) { $cfg.manifest_url = [string]$fileCfg.manifest_url }
        if ($null -ne $fileCfg.auto_check_on_start) {
            $cfg.auto_check_on_start = [System.Convert]::ToBoolean($fileCfg.auto_check_on_start)
        }
        if ($fileCfg.ui_language) {
            $cfg.ui_language = Normalize-UiLanguage ([string]$fileCfg.ui_language)
        }
    } catch {
        Write-DebugLog -Level "WARN" -Message ("PROGRAM_UPDATE config load failed: " + $_.Exception.Message)
    }

    return $cfg
}

function Save-ProgramUpdateConfig {
    param(
        [string]$UiLanguage = ""
    )

    try {
        $cfg = Get-ProgramUpdateConfig
        if (-not [string]::IsNullOrWhiteSpace($UiLanguage)) {
            $cfg.ui_language = Normalize-UiLanguage $UiLanguage
        } else {
            $cfg.ui_language = Normalize-UiLanguage $cfg.ui_language
        }

        $toSave = [ordered]@{
            app_version = [string]$cfg.app_version
            manifest_url = [string]$cfg.manifest_url
            auto_check_on_start = [bool]$cfg.auto_check_on_start
            ui_language = [string]$cfg.ui_language
        }
        $json = $toSave | ConvertTo-Json -Depth 3
        [System.IO.File]::WriteAllText($script:ProgramUpdateConfigPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        Write-DebugLog -Level "WARN" -Message ("PROGRAM_UPDATE config save failed: " + $_.Exception.Message)
    }
}

function Convert-ToComparableVersion {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw (T "error_empty_version")
    }

    $parts = @()
    foreach ($chunk in ($Value -split "[^0-9]+")) {
        if ($chunk -match "^[0-9]+$") {
            $parts += [int]$chunk
        }
    }

    if ($parts.Count -eq 0) {
        throw (TF "error_parse_version_fmt" $Value)
    }

    while ($parts.Count -lt 4) { $parts += 0 }
    if ($parts.Count -gt 4) { $parts = $parts[0..3] }
    return [Version]::new($parts[0], $parts[1], $parts[2], $parts[3])
}

function Compare-VersionStrings {
    param(
        [string]$Current,
        [string]$Remote
    )
    $currentVer = Convert-ToComparableVersion -Value $Current
    $remoteVer = Convert-ToComparableVersion -Value $Remote
    return $currentVer.CompareTo($remoteVer)
}

function Get-ProgramUpdateManifest {
    param([string]$ManifestUrl)

    $resp = Invoke-WebRequest -Uri $ManifestUrl -UseBasicParsing -TimeoutSec 25
    $obj = $resp.Content | ConvertFrom-Json
    if (-not $obj.version) {
        throw (T "error_manifest_no_version")
    }
    if (-not $obj.package_url) {
        throw (T "error_manifest_no_package_url")
    }
    return $obj
}

function Start-ProgramSelfUpdate {
    param(
        [switch]$SilentWhenUpToDate,
        [switch]$SilentOnError,
        [switch]$AutoTriggered
    )

    if ($script:ProcessRunning) { return }
    Write-DebugLog -Message ("PROGRAM_UPDATE start auto=" + $AutoTriggered.IsPresent + " silent=" + $SilentWhenUpToDate.IsPresent)

    Set-Busy $true (T "status_checking_program_update")
    try {
        $cfg = Get-ProgramUpdateConfig
        $currentVersion = [string]$cfg.app_version
        $manifestUrl = [string]$cfg.manifest_url

        Append-Log (TF "log_program_local_ver_fmt" $currentVersion)
        Append-Log (TF "log_manifest_url_fmt" $manifestUrl)

        $manifest = Get-ProgramUpdateManifest -ManifestUrl $manifestUrl
        $remoteVersion = [string]$manifest.version
        $packageUrl = [string]$manifest.package_url
        $expectedHash = ""
        if ($manifest.sha256) { $expectedHash = ([string]$manifest.sha256).ToLowerInvariant() }

        Append-Log (TF "log_program_remote_ver_fmt" $remoteVersion)
        $cmp = Compare-VersionStrings -Current $currentVersion -Remote $remoteVersion
        if ($cmp -ge 0) {
            Set-Busy $false (T "status_program_up_to_date")
            if ($SilentWhenUpToDate) {
                Append-Log (T "log_auto_check_up_to_date")
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    (TF "msg_latest_program_fmt" $currentVersion),
                    (T "cap_program_update"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            }
            return
        }

        if ($AutoTriggered) {
            Append-Log (T "log_auto_check_found_new")
            [System.Windows.Forms.MessageBox]::Show(
                (TF "msg_program_new_auto_fmt" $remoteVersion $currentVersion),
                (T "cap_program_update"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } else {
            $question = (TF "msg_program_new_question_fmt" $remoteVersion $currentVersion)
            $ans = [System.Windows.Forms.MessageBox]::Show(
                $question,
                (T "cap_program_update"),
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button1
            )
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
                Set-Busy $false (T "status_program_update_canceled")
                return
            }
        }

        Set-Status (T "status_downloading_program_update")
        $updatesDir = Join-Path $script:BaseDir "downloads\\program_updates"
        if (-not (Test-Path $updatesDir)) { [void](New-Item -ItemType Directory -Path $updatesDir) }
        $safeVersion = ($remoteVersion -replace "[^0-9A-Za-z._-]", "_")
        $zipPath = Join-Path $updatesDir ("OpenGD77_UV390_A11y_" + $safeVersion + ".zip")
        Append-Log (TF "log_download_pkg_fmt" $packageUrl)
        Invoke-WebRequest -Uri $packageUrl -OutFile $zipPath -UseBasicParsing -TimeoutSec 300
        Append-Log (TF "log_downloaded_fmt" $zipPath)

        if (-not [string]::IsNullOrWhiteSpace($expectedHash)) {
            $actualHash = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($actualHash -ne $expectedHash) {
                throw (TF "error_sha_mismatch_fmt" $expectedHash $actualHash)
            }
            Append-Log (T "log_sha_ok")
        } else {
            Append-Log (T "log_manifest_no_sha")
        }

        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $helperPath = Join-Path $updatesDir ("apply_program_update_" + $stamp + ".ps1")
        $helperLogPath = Join-Path $script:LogDir ("ProgramSelfUpdate_" + $stamp + ".log")
        $restartVbs = Join-Path $script:BaseDir "start_OpenGD77_UV390_A11y.vbs"

        $helperScript = @'
param(
    [string]$ZipPath,
    [string]$TargetDir,
    [string]$RestartVbs,
    [int]$MainPid,
    [string]$LogPath
)
$ErrorActionPreference = "Stop"

function Write-UpdateLog {
    param([string]$Message)
    $line = ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

try {
    Write-UpdateLog "Updater helper started."
    Write-UpdateLog ("ZipPath=" + $ZipPath)
    Write-UpdateLog ("TargetDir=" + $TargetDir)
    Write-UpdateLog ("MainPid=" + $MainPid)

    if ($MainPid -gt 0) {
        for ($i = 0; $i -lt 240; $i++) {
            $p = Get-Process -Id $MainPid -ErrorAction SilentlyContinue
            if (-not $p) { break }
            Start-Sleep -Milliseconds 500
        }
    }

    if (-not (Test-Path $ZipPath)) {
        throw ("Nie znaleziono paczki update: " + $ZipPath)
    }

    $stagingRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("OpenGD77_ProgramUpdate_" + [Guid]::NewGuid().ToString("N"))
    [void](New-Item -ItemType Directory -Path $stagingRoot)
    Expand-Archive -Path $ZipPath -DestinationPath $stagingRoot -Force
    Write-UpdateLog ("Expanded to: " + $stagingRoot)

    $sourceDir = $stagingRoot
    $rootDirs = @(Get-ChildItem -Path $stagingRoot -Directory)
    $rootFiles = @(Get-ChildItem -Path $stagingRoot -File)
    if (($rootDirs.Count -eq 1) -and ($rootFiles.Count -eq 0)) {
        $sourceDir = $rootDirs[0].FullName
    }
    Write-UpdateLog ("SourceDir=" + $sourceDir)

    $null = robocopy $sourceDir $TargetDir /E /R:2 /W:1 /NFL /NDL /NP /NJH /NJS
    $rc = $LASTEXITCODE
    Write-UpdateLog ("Robocopy exit code: " + $rc)
    if ($rc -ge 8) {
        throw ("Robocopy failed with code " + $rc)
    }

    if (Test-Path $RestartVbs) {
        Start-Process -FilePath "wscript.exe" -ArgumentList @("""$RestartVbs""")
        Write-UpdateLog ("Restarted via VBS: " + $RestartVbs)
    } else {
        $mainPs1 = Join-Path $TargetDir "OpenGD77_UV390_A11y.ps1"
        Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoProfile","-ExecutionPolicy","Bypass","-File",$mainPs1)
        Write-UpdateLog ("Restarted via PS1: " + $mainPs1)
    }

    Write-UpdateLog "Updater helper completed successfully."
} catch {
    Write-UpdateLog ("ERROR: " + $_.Exception.Message)
    throw
}
'@
        Set-Content -Path $helperPath -Value $helperScript -Encoding UTF8

        $launchArgs = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$helperPath`" -ZipPath `"$zipPath`" -TargetDir `"$script:BaseDir`" -RestartVbs `"$restartVbs`" -MainPid $PID -LogPath `"$helperLogPath`""
        [void](Start-Process -FilePath "powershell.exe" -ArgumentList $launchArgs -WindowStyle Hidden -PassThru)
        Append-Log (TF "log_launch_helper_fmt" $helperPath)

        Set-Busy $false (T "status_restart_program_update")
        $script:SkipCloseConfirmation = $true
        [System.Windows.Forms.MessageBox]::Show(
            (T "msg_program_update_started"),
            (T "cap_program_update"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        $script:form.Close()
    } catch {
        Write-DebugLog -Level "ERROR" -Message ("PROGRAM_UPDATE failed: " + $_.Exception.Message)
        Set-Busy $false (T "status_program_update_error")
        if ($SilentOnError) {
            Append-Log (TF "log_autocheck_error_fmt" $_.Exception.Message)
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                (TF "msg_program_update_error_fmt" $_.Exception.Message),
                (T "cap_error"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    }
}

function Play-StartupRadioSound {
    Write-DebugLog -Message ("SOUND_STARTUP: requested file=" + $script:StartupSoundPath)
    try {
        if (Test-Path $script:StartupSoundPath) {
            $player = New-Object System.Media.SoundPlayer($script:StartupSoundPath)
            try {
                $player.Load()
                $player.PlaySync()
                Write-DebugLog -Message "SOUND_STARTUP: played external CC0 radio sound"
                return
            } finally {
                $player.Dispose()
            }
        }
        Write-DebugLog -Level "WARN" -Message "SOUND_STARTUP: file missing, using fallback chirp"
    } catch {
        Write-DebugLog -Level "WARN" -Message ("SOUND_STARTUP failed: " + $_.Exception.Message)
    }

    Play-RadioSuccessChirp
}

function Get-TriangleSuccessWaveBytes {
    param([int]$SampleRate = 22050)

    $tones = @(
        # Wcześniejsza wersja 3-tonowa.
        @{ Freq = 880; DurationMs = 220 },
        @{ Freq = 1320; DurationMs = 260 },
        @{ Freq = 1760; DurationMs = 340 }
    )
    $silenceMs = 70
    $channels = 1
    $bitsPerSample = 16
    $blockAlign = [int]($channels * ($bitsPerSample / 8))
    $byteRate = [int]($SampleRate * $blockAlign)
    $silenceSamples = [int][Math]::Round($SampleRate * ($silenceMs / 1000.0))
    $amplitude = 15000
    $fadeSamples = [Math]::Max(1, [int][Math]::Round($SampleRate * 0.006))

    $totalSamples = 0
    for ($idx = 0; $idx -lt $tones.Count; $idx++) {
        $toneSamples = [int][Math]::Round($SampleRate * ($tones[$idx].DurationMs / 1000.0))
        $totalSamples += $toneSamples
        if ($idx -lt ($tones.Count - 1)) {
            $totalSamples += $silenceSamples
        }
    }
    $dataSize = [int]($totalSamples * $blockAlign)
    $riffSize = [int](36 + $dataSize)

    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms, [System.Text.Encoding]::ASCII, $true)
    try {
        $bw.Write([System.Text.Encoding]::ASCII.GetBytes("RIFF"))
        $bw.Write($riffSize)
        $bw.Write([System.Text.Encoding]::ASCII.GetBytes("WAVE"))
        $bw.Write([System.Text.Encoding]::ASCII.GetBytes("fmt "))
        $bw.Write([int]16)
        $bw.Write([int16]1)
        $bw.Write([int16]$channels)
        $bw.Write([int]$SampleRate)
        $bw.Write([int]$byteRate)
        $bw.Write([int16]$blockAlign)
        $bw.Write([int16]$bitsPerSample)
        $bw.Write([System.Text.Encoding]::ASCII.GetBytes("data"))
        $bw.Write([int]$dataSize)

        for ($idx = 0; $idx -lt $tones.Count; $idx++) {
            $freq = [double]$tones[$idx].Freq
            $toneSamples = [int][Math]::Round($SampleRate * ($tones[$idx].DurationMs / 1000.0))
            for ($n = 0; $n -lt $toneSamples; $n++) {
                $phase = (($n * $freq) / [double]$SampleRate) % 1.0
                $triangle = (2.0 * [Math]::Abs((2.0 * $phase) - 1.0)) - 1.0

                $env = 1.0
                if ($n -lt $fadeSamples) {
                    $env = $n / [double]$fadeSamples
                }
                $fromEnd = ($toneSamples - 1) - $n
                if ($fromEnd -lt $fadeSamples) {
                    $tail = $fromEnd / [double]$fadeSamples
                    if ($tail -lt $env) { $env = $tail }
                }

                $sample = [int16][Math]::Round($triangle * $env * $amplitude)
                $bw.Write($sample)
            }

            if ($idx -lt ($tones.Count - 1)) {
                for ($s = 0; $s -lt $silenceSamples; $s++) {
                    $bw.Write([int16]0)
                }
            }
        }

        $bw.Flush()
        return $ms.ToArray()
    } finally {
        $bw.Dispose()
        $ms.Dispose()
    }
}

function Play-RadioSuccessChirp {
    Write-DebugLog -Message "SOUND: success chirp requested"
    try {
        if (-not $script:SuccessChirpWaveBytes) {
            $script:SuccessChirpWaveBytes = Get-TriangleSuccessWaveBytes
        }
        $stream = New-Object System.IO.MemoryStream(, $script:SuccessChirpWaveBytes)
        $player = New-Object System.Media.SoundPlayer($stream)
        try {
            $player.Load()
            $player.PlaySync()
            Write-DebugLog -Message "SOUND: success chirp played as triangle wave"
            return
        } finally {
            $player.Dispose()
            $stream.Dispose()
        }
    } catch {
        Write-DebugLog -Level "WARN" -Message ("SOUND triangle-wave failed: " + $_.Exception.Message)
    }

    try {
        [Console]::Beep(880, 220)
        Start-Sleep -Milliseconds 70
        [Console]::Beep(1320, 260)
        Start-Sleep -Milliseconds 70
        [Console]::Beep(1760, 340)
        Write-DebugLog -Message "SOUND: fallback chirp played via Console.Beep"
        return
    } catch {
        Write-DebugLog -Level "WARN" -Message ("SOUND beep failed: " + $_.Exception.Message)
    }

    try {
        [System.Media.SystemSounds]::Asterisk.Play()
        Start-Sleep -Milliseconds 120
        [System.Media.SystemSounds]::Asterisk.Play()
        Write-DebugLog -Message "SOUND: fallback SystemSounds played"
    } catch {
        Write-DebugLog -Level "WARN" -Message ("SOUND fallback failed: " + $_.Exception.Message)
    }
}

function Start-BackendProcess {
    param(
        [string]$Arguments
    )

    if ($script:ProcessRunning) { return }
    Write-DebugLog -Message ("START_BACKEND requested with args: " + $Arguments)
    if (-not (Test-Path $script:BackendExe)) {
        Write-DebugLog -Level "ERROR" -Message ("Backend missing: " + $script:BackendExe)
        [System.Windows.Forms.MessageBox]::Show(
            (TF "msg_backend_missing_fmt" $script:BackendExe),
            (T "cap_error"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    Append-Log ("Start: " + $Arguments)
    Set-Busy $true (T "status_busy")

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $p.StartInfo.FileName = $script:BackendExe
    $p.StartInfo.Arguments = $Arguments
    $p.StartInfo.WorkingDirectory = $script:BaseDir
    $p.StartInfo.UseShellExecute = $false
    $p.StartInfo.RedirectStandardOutput = $true
    $p.StartInfo.RedirectStandardError = $true
    $p.StartInfo.CreateNoWindow = $true
    $p.EnableRaisingEvents = $true

    $p.add_OutputDataReceived({
        param($sender, $eventArgs)
        if ($eventArgs.Data) {
            Write-DebugLog -Message ("BACKEND_OUT: " + $eventArgs.Data)
            $script:LogQueue.Enqueue($eventArgs.Data)
        }
    })
    $p.add_ErrorDataReceived({
        param($sender, $eventArgs)
        if ($eventArgs.Data) {
            Write-DebugLog -Level "WARN" -Message ("BACKEND_ERR: " + $eventArgs.Data)
            $script:LogQueue.Enqueue("[ERR] " + $eventArgs.Data)
        }
    })
    $p.add_Exited({
        param($sender, $eventArgs)
        Write-DebugLog -Message ("BACKEND_EXIT code=" + $sender.ExitCode)
        $script:EventQueue.Enqueue("EXIT:" + $sender.ExitCode)
    })

    $script:CurrentProcess = $p
    $script:BackendExitHandled = $false
    [void]$p.Start()
    Write-DebugLog -Message ("BACKEND_PID=" + $p.Id)
    $p.BeginOutputReadLine()
    $p.BeginErrorReadLine()
}

function Run-SelfTest {
    Write-DebugLog "SELFTEST_START"
    Write-Output "[SELFTEST] Start"
    if (-not (Test-Path $script:BackendExe)) {
        Write-DebugLog -Level "ERROR" -Message ("SELFTEST_FAIL backend missing: " + $script:BackendExe)
        Write-Output "[SELFTEST] FAIL: backend missing"
        return 2
    }

    $output = @()
    try {
        $output = & $script:BackendExe --check-only 2>&1
        $rc = $LASTEXITCODE
    } catch {
        Write-DebugLog -Level "ERROR" -Message ("SELFTEST_EXCEPTION: " + $_.Exception.Message)
        Write-Output "[SELFTEST] FAIL: exception " + $_.Exception.Message
        return 3
    }

    foreach ($line in $output) {
        Write-DebugLog -Message ("SELFTEST_OUT: " + $line)
    }
    Write-DebugLog -Message ("SELFTEST_EXIT_CODE: " + $rc)

    if ($rc -ne 0) {
        Write-Output "[SELFTEST] FAIL: backend exit code $rc"
        return 4
    }

    if (-not ($output -match "Najnowszy release")) {
        Write-Output "[SELFTEST] WARN: expected text missing"
        Write-DebugLog -Level "WARN" -Message "SELFTEST_WARN expected text missing"
    }

    Write-Output "[SELFTEST] OK"
    Write-DebugLog "SELFTEST_OK"
    return 0
}

if ($SelfTest) {
    $selfTestCode = Run-SelfTest
    exit $selfTestCode
}

try {
    $startupCfg = Get-ProgramUpdateConfig
    $script:UiLanguage = Normalize-UiLanguage $startupCfg.ui_language
} catch {
    $script:UiLanguage = "pl"
}

$form = New-Object System.Windows.Forms.Form
$form.Text = (T "form_title")
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(980, 700)
$form.MinimumSize = New-Object System.Drawing.Size(860, 580)
$form.KeyPreview = $true
Write-DebugLog "FORM_INITIALIZED"

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = (T "title")
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(12, 12)
$form.Controls.Add($lblTitle)

$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Text = (T "options")
$grpOptions.Location = New-Object System.Drawing.Point(12, 40)
$grpOptions.Size = New-Object System.Drawing.Size(940, 110)
$form.Controls.Add($grpOptions)

$chkCheckOnly = New-Object System.Windows.Forms.CheckBox
$chkCheckOnly.Text = (T "check_only")
$chkCheckOnly.Location = New-Object System.Drawing.Point(16, 26)
$chkCheckOnly.AutoSize = $true
$chkCheckOnly.TabIndex = 0
$grpOptions.Controls.Add($chkCheckOnly)

$chkAutoDriver = New-Object System.Windows.Forms.CheckBox
$chkAutoDriver.Text = (T "auto_driver")
$chkAutoDriver.Location = New-Object System.Drawing.Point(16, 52)
$chkAutoDriver.AutoSize = $true
$chkAutoDriver.TabIndex = 1
$grpOptions.Controls.Add($chkAutoDriver)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = (T "timeout")
$lblTimeout.Location = New-Object System.Drawing.Point(16, 78)
$lblTimeout.AutoSize = $true
$grpOptions.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(180, 76)
$numTimeout.Minimum = 5
$numTimeout.Maximum = 900
$numTimeout.Value = 90
$numTimeout.TabIndex = 2
$grpOptions.Controls.Add($numTimeout)

$lblLanguage = New-Object System.Windows.Forms.Label
$lblLanguage.Text = (T "ui_language")
$lblLanguage.Location = New-Object System.Drawing.Point(620, 78)
$lblLanguage.AutoSize = $true
$grpOptions.Controls.Add($lblLanguage)

$cmbLanguage = New-Object System.Windows.Forms.ComboBox
$cmbLanguage.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbLanguage.Location = New-Object System.Drawing.Point(760, 74)
$cmbLanguage.Size = New-Object System.Drawing.Size(160, 24)
$cmbLanguage.TabIndex = 3
[void]$cmbLanguage.Items.Add("Polski")
[void]$cmbLanguage.Items.Add("English")
if ($script:UiLanguage -eq "en") {
    $cmbLanguage.SelectedIndex = 1
} else {
    $cmbLanguage.SelectedIndex = 0
}
$grpOptions.Controls.Add($cmbLanguage)

$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Text = (T "btn_check")
$btnCheck.Location = New-Object System.Drawing.Point(12, 162)
$btnCheck.Size = New-Object System.Drawing.Size(160, 32)
$btnCheck.TabIndex = 4
$form.Controls.Add($btnCheck)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = (T "btn_start")
$btnStart.Location = New-Object System.Drawing.Point(180, 162)
$btnStart.Size = New-Object System.Drawing.Size(160, 32)
$btnStart.TabIndex = 5
$form.Controls.Add($btnStart)

$btnProgramUpdate = New-Object System.Windows.Forms.Button
$btnProgramUpdate.Text = (T "btn_program_update")
$btnProgramUpdate.Location = New-Object System.Drawing.Point(348, 162)
$btnProgramUpdate.Size = New-Object System.Drawing.Size(160, 32)
$btnProgramUpdate.TabIndex = 6
$form.Controls.Add($btnProgramUpdate)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = (T "btn_close")
$btnClose.Location = New-Object System.Drawing.Point(516, 162)
$btnClose.Size = New-Object System.Drawing.Size(120, 32)
$btnClose.TabIndex = 7
$form.Controls.Add($btnClose)

$lblHints = New-Object System.Windows.Forms.Label
$lblHints.Text = (T "hints")
$lblHints.AutoSize = $true
$lblHints.Location = New-Object System.Drawing.Point(12, 204)
$form.Controls.Add($lblHints)

$grpStatus = New-Object System.Windows.Forms.GroupBox
$grpStatus.Text = (T "status_group")
$grpStatus.Location = New-Object System.Drawing.Point(12, 228)
$grpStatus.Size = New-Object System.Drawing.Size(940, 88)
$form.Controls.Add($grpStatus)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = (T "status_label")
$lblStatus.Location = New-Object System.Drawing.Point(12, 26)
$lblStatus.AutoSize = $true
$grpStatus.Controls.Add($lblStatus)

$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Location = New-Object System.Drawing.Point(120, 22)
$txtStatus.Size = New-Object System.Drawing.Size(800, 24)
$txtStatus.ReadOnly = $true
$txtStatus.TabStop = $true
$txtStatus.TabIndex = 8
$txtStatus.Text = (T "status_ready")
$grpStatus.Controls.Add($txtStatus)

$lblLastMessage = New-Object System.Windows.Forms.Label
$lblLastMessage.Text = (T "last_message_label")
$lblLastMessage.Location = New-Object System.Drawing.Point(12, 56)
$lblLastMessage.AutoSize = $true
$grpStatus.Controls.Add($lblLastMessage)

$txtLastMessage = New-Object System.Windows.Forms.TextBox
$txtLastMessage.Location = New-Object System.Drawing.Point(120, 52)
$txtLastMessage.Size = New-Object System.Drawing.Size(800, 24)
$txtLastMessage.ReadOnly = $true
$txtLastMessage.TabStop = $true
$txtLastMessage.TabIndex = 9
$txtLastMessage.Text = (T "status_no_messages")
$grpStatus.Controls.Add($txtLastMessage)

$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = (T "log_group")
$grpLog.Location = New-Object System.Drawing.Point(12, 324)
$grpLog.Size = New-Object System.Drawing.Size(940, 320)
$form.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Location = New-Object System.Drawing.Point(12, 24)
$txtLog.Size = New-Object System.Drawing.Size(916, 286)
$txtLog.TabIndex = 10
$grpLog.Controls.Add($txtLog)

# expose controls to helper functions
$script:lblTitle = $lblTitle
$script:grpOptions = $grpOptions
$script:chkCheckOnly = $chkCheckOnly
$script:chkAutoDriver = $chkAutoDriver
$script:lblTimeout = $lblTimeout
$script:lblLanguage = $lblLanguage
$script:cmbLanguage = $cmbLanguage
$script:numTimeout = $numTimeout
$script:btnCheck = $btnCheck
$script:btnStart = $btnStart
$script:btnProgramUpdate = $btnProgramUpdate
$script:btnClose = $btnClose
$script:lblHints = $lblHints
$script:grpStatus = $grpStatus
$script:lblStatus = $lblStatus
$script:lblLastMessage = $lblLastMessage
$script:grpLog = $grpLog
$script:txtStatus = $txtStatus
$script:txtLastMessage = $txtLastMessage
$script:txtLog = $txtLog
$script:form = $form

function Apply-UiLanguage {
    if (-not $script:form) { return }

    $script:form.Text = (T "form_title")
    $script:lblTitle.Text = (T "title")
    $script:grpOptions.Text = (T "options")
    $script:chkCheckOnly.Text = (T "check_only")
    $script:chkAutoDriver.Text = (T "auto_driver")
    $script:lblTimeout.Text = (T "timeout")
    $script:lblLanguage.Text = (T "ui_language")
    $script:btnCheck.Text = (T "btn_check")
    $script:btnStart.Text = (T "btn_start")
    $script:btnProgramUpdate.Text = (T "btn_program_update")
    $script:btnClose.Text = (T "btn_close")
    $script:lblHints.Text = (T "hints")
    $script:grpStatus.Text = (T "status_group")
    $script:lblStatus.Text = (T "status_label")
    $script:lblLastMessage.Text = (T "last_message_label")
    $script:grpLog.Text = (T "log_group")

    if ([string]::IsNullOrWhiteSpace($script:txtStatus.Text) -or
        $script:txtStatus.Text -in @($script:I18n.pl.status_ready, $script:I18n.en.status_ready)) {
        $script:txtStatus.Text = (T "status_ready")
    }
    if ([string]::IsNullOrWhiteSpace($script:txtLastMessage.Text) -or
        $script:txtLastMessage.Text -in @($script:I18n.pl.status_no_messages, $script:I18n.en.status_no_messages)) {
        $script:txtLastMessage.Text = (T "status_no_messages")
    }
}

Apply-UiLanguage

$cmbLanguage.Add_SelectedIndexChanged({
    if ($script:cmbLanguage.SelectedIndex -eq 1) {
        $newLang = "en"
    } else {
        $newLang = "pl"
    }
    if ($newLang -eq $script:UiLanguage) { return }

    $script:UiLanguage = $newLang
    Apply-UiLanguage
    Save-ProgramUpdateConfig -UiLanguage $newLang
    Append-Log (TF "log_language_changed_fmt" (Get-LanguageDisplayName $newLang))
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 100
$timer.Add_Tick({
    try {
        $line = $null
        while ($script:LogQueue.TryDequeue([ref]$line)) {
            Append-Log $line
        }

        $evt = $null
        while ($script:EventQueue.TryDequeue([ref]$evt)) {
            if ($evt -like "EXIT:*") {
                $code = [int]($evt.Split(":")[1])
                if ($code -eq 0) {
                    Set-Busy $false (T "status_success")
                    if ($script:CurrentOperationType -eq "flash") {
                        Play-RadioSuccessChirp
                    }
                } else {
                    Set-Busy $false (TF "status_error_code_fmt" $code)
                }
                $script:BackendExitHandled = $true
            }
        }

        if ($script:ProcessRunning -and $script:CurrentProcess -and $script:CurrentProcess.HasExited -and (-not $script:BackendExitHandled)) {
            $fallbackCode = $script:CurrentProcess.ExitCode
            Write-DebugLog -Level "WARN" -Message ("BACKEND_EXIT_FALLBACK code=" + $fallbackCode)
            if ($fallbackCode -eq 0) {
                Set-Busy $false (T "status_success")
                if ($script:CurrentOperationType -eq "flash") {
                    Play-RadioSuccessChirp
                }
            } else {
                Set-Busy $false (TF "status_error_code_fmt" $fallbackCode)
            }
            $script:BackendExitHandled = $true
        }
    } catch {
        Write-DebugLog -Level "ERROR" -Message ("TIMER_TICK_ERR: " + $_.Exception.Message)
    }
})
$timer.Start()

$btnCheck.Add_Click({
    if ($script:ProcessRunning) { return }
    Write-DebugLog "UI_CLICK: Sprawdź wersję"
    $script:CurrentOperationType = "check"
    Start-BackendProcess "--check-only"
})

$btnStart.Add_Click({
    if ($script:ProcessRunning) { return }
    Write-DebugLog "UI_CLICK: Start aktualizacji"

    if ($script:chkCheckOnly.Checked) {
        Write-DebugLog "UI_OPTION: check-only selected in start path"
        $script:CurrentOperationType = "check"
        Start-BackendProcess "--check-only"
        return
    }

    $ok = [System.Windows.Forms.MessageBox]::Show(
        (T "msg_dfu_instructions"),
        (T "cap_dfu"),
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    if ($ok -ne [System.Windows.Forms.DialogResult]::OK) { return }
    Write-DebugLog "UI_CONFIRM: DFU dialog OK"

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        (T "msg_start_update_now"),
        (T "cap_confirmation"),
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    Write-DebugLog "UI_CONFIRM: Start flash YES"

    $args = "--force --auto-confirm --dfu-timeout " + [int]$script:numTimeout.Value
    if ($script:chkAutoDriver.Checked) {
        $args += " --auto-driver-install"
    }
    $script:CurrentOperationType = "flash"
    Start-BackendProcess $args
})

$btnProgramUpdate.Add_Click({
    if ($script:ProcessRunning) { return }
    Write-DebugLog "UI_CLICK: Aktualizuj program"
    Start-ProgramSelfUpdate
})

$btnClose.Add_Click({
    Write-DebugLog "UI_CLICK: Zamknij"
    $form.Close()
})

$form.Add_KeyDown({
    param($sender, $e)
    if ($e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::S) {
        $btnCheck.PerformClick()
        $e.SuppressKeyPress = $true
    } elseif ($e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::T) {
        $btnStart.PerformClick()
        $e.SuppressKeyPress = $true
    } elseif ($e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $btnProgramUpdate.PerformClick()
        $e.SuppressKeyPress = $true
    } elseif ($e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::L) {
        $txtLog.Focus()
        $e.SuppressKeyPress = $true
    } elseif ($e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::D) {
        $numTimeout.Focus()
        $e.SuppressKeyPress = $true
    } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::F1) {
        [System.Windows.Forms.MessageBox]::Show(
            (T "msg_help"),
            (T "cap_help"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        $e.SuppressKeyPress = $true
    }
})

$form.Add_FormClosing({
    param($sender, $e)
    Write-DebugLog "FORM_CLOSING requested"
    if ($script:SkipCloseConfirmation) {
        Write-DebugLog "FORM_CLOSING skip confirmation flag active"
        return
    }
    if ($script:ProcessRunning) {
        Write-DebugLog -Level "WARN" -Message "FORM_CLOSING blocked: update in progress"
        [System.Windows.Forms.MessageBox]::Show(
            (T "msg_cannot_close_during_update"),
            (T "cap_update_in_progress"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        $e.Cancel = $true
        return
    }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        (T "msg_confirm_close"),
        (T "cap_confirm_close"),
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-DebugLog "FORM_CLOSING canceled by user"
        $e.Cancel = $true
    } else {
        Write-DebugLog "FORM_CLOSING confirmed by user"
    }
})

$form.Add_Shown({
    Write-DebugLog "FORM_SHOWN"
    Play-StartupRadioSound
    try {
        $cfg = Get-ProgramUpdateConfig
        if ($cfg.auto_check_on_start) {
            Write-DebugLog "PROGRAM_UPDATE auto-check scheduled on startup"
            $script:StartupProgramUpdateTimer = New-Object System.Windows.Forms.Timer
            $script:StartupProgramUpdateTimer.Interval = 1200
            $script:StartupProgramUpdateTimer.Add_Tick({
                $script:StartupProgramUpdateTimer.Stop()
                $script:StartupProgramUpdateTimer.Dispose()
                $script:StartupProgramUpdateTimer = $null
                Start-ProgramSelfUpdate -SilentWhenUpToDate -SilentOnError -AutoTriggered
            })
            $script:StartupProgramUpdateTimer.Start()
        } else {
            Write-DebugLog "PROGRAM_UPDATE auto-check disabled in config"
        }
    } catch {
        Write-DebugLog -Level "WARN" -Message ("PROGRAM_UPDATE startup scheduling failed: " + $_.Exception.Message)
    }
})

$form.Add_FormClosed({
    Write-DebugLog "FORM_CLOSED"
})

$btnCheck.Focus()
[void]$form.ShowDialog()
Write-DebugLog "A11y launcher exit"
