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
        btn_channels_file = "Kanały z &pliku"
        btn_channels_url = "Przemienniki z &internetu"
        btn_close = "&Zamknij"
        btn_ok = "OK"
        btn_cancel = "Anuluj"
        hints = "Skróty: Alt+S Sprawdź, Alt+T Start, Alt+A Aktualizuj program, Alt+K Kanały z pliku, Alt+U Przemienniki z internetu, Alt+L Log, Alt+D Timeout, Alt+F4 Zamknij, F1 Pomoc."
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
        cap_channels_import = "Import kanałów/przemienników"
        cap_channels_url = "Źródło internetowe"
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
        msg_help = "Sterowanie:`n- Alt+S: Sprawdź wersję`n- Alt+T: Start aktualizacji`n- Alt+A: Aktualizuj program`n- Alt+K: Kanały z pliku`n- Alt+U: Przemienniki z internetu`n- Alt+L: Fokus log`n- Alt+D: Fokus timeout`n- Alt+F4: Zamknij"
        msg_channels_url_prompt = "Wklej URL do pliku CSV lub JSON z kanałami/przemiennikami:"
        msg_channels_no_data = "Nie znaleziono wpisów kanałów/przemienników."
        msg_channels_saved_fmt = "Zapisano {0} wpisów do: {1}"
        msg_channels_saved_with_cps_fmt = "Zapisano {0} wpisów do: {1}`nUtworzono pakiet OpenGD77 CPS: {2}`nW CPS użyj File -> CSV -> Append CSV."
        msg_channels_error_fmt = "Błąd importu kanałów/przemienników:`n{0}"
        msg_cannot_close_during_update = "Nie można zamknąć programu podczas aktualizacji."
        msg_confirm_close = "Czy na pewno chcesz zamknąć program?"
        filter_channels_open = "Kanały/Przemienniki (*.csv;*.json)|*.csv;*.json|CSV (*.csv)|*.csv|JSON (*.json)|*.json|Wszystkie pliki (*.*)|*.*"
        filter_csv_save = "CSV (*.csv)|*.csv|Wszystkie pliki (*.*)|*.*"
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
        log_channels_source_fmt = "Źródło listy kanałów: {0}"
        log_channels_loaded_fmt = "Wczytano wpisów: {0}"
        log_channels_saved_fmt = "Zapisano listę kanałów: {0}"
        log_channels_cps_saved_fmt = "Zapisano pakiet OpenGD77 CPS: {0}"
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
        btn_channels_file = "Channels from &file"
        btn_channels_url = "Repeaters from &internet"
        btn_close = "&Close"
        btn_ok = "OK"
        btn_cancel = "Cancel"
        hints = "Shortcuts: Alt+S Check, Alt+T Start, Alt+A Update app, Alt+K Channels from file, Alt+U Repeaters from internet, Alt+L Log, Alt+D Timeout, Alt+F4 Close, F1 Help."
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
        cap_channels_import = "Channels/repeaters import"
        cap_channels_url = "Internet source"
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
        msg_help = "Controls:`n- Alt+S: Check version`n- Alt+T: Start update`n- Alt+A: Update app`n- Alt+K: Channels from file`n- Alt+U: Repeaters from internet`n- Alt+L: Log focus`n- Alt+D: Timeout focus`n- Alt+F4: Close"
        msg_channels_url_prompt = "Paste URL to CSV or JSON file with channels/repeaters:"
        msg_channels_no_data = "No channel/repeater entries found."
        msg_channels_saved_fmt = "Saved {0} entries to: {1}"
        msg_channels_saved_with_cps_fmt = "Saved {0} entries to: {1}`nCreated OpenGD77 CPS bundle: {2}`nIn CPS use File -> CSV -> Append CSV."
        msg_channels_error_fmt = "Channel/repeater import error:`n{0}"
        msg_cannot_close_during_update = "You cannot close the app during update."
        msg_confirm_close = "Are you sure you want to close the app?"
        filter_channels_open = "Channels/Repeaters (*.csv;*.json)|*.csv;*.json|CSV (*.csv)|*.csv|JSON (*.json)|*.json|All files (*.*)|*.*"
        filter_csv_save = "CSV (*.csv)|*.csv|All files (*.*)|*.*"
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
        log_channels_source_fmt = "Channel list source: {0}"
        log_channels_loaded_fmt = "Loaded entries: {0}"
        log_channels_saved_fmt = "Saved channel list: {0}"
        log_channels_cps_saved_fmt = "Saved OpenGD77 CPS bundle: {0}"
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
    if ($script:btnChannelsFile) { $script:btnChannelsFile.Enabled = -not $Busy }
    if ($script:btnChannelsUrl) { $script:btnChannelsUrl.Enabled = -not $Busy }
}

function Get-NormalizedObjectValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Object,
        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    if ($null -eq $Object) { return $null }
    $map = @{}
    foreach ($prop in $Object.PSObject.Properties) {
        $key = ($prop.Name.ToLowerInvariant() -replace "[^a-z0-9]", "")
        if (-not $map.ContainsKey($key)) {
            $map[$key] = $prop.Value
        }
    }

    foreach ($name in $Names) {
        $key = ($name.ToLowerInvariant() -replace "[^a-z0-9]", "")
        if ($map.ContainsKey($key)) {
            return $map[$key]
        }
    }
    return $null
}

function Convert-ToNullableFrequency {
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $s = $s.Trim()
    $s = $s -replace "(?i)mhz", ""
    $s = $s -replace "[^0-9,.\-]", ""
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }

    if ($s.Contains(",") -and -not $s.Contains(".")) {
        $s = $s.Replace(",", ".")
    } elseif ($s.Contains(",") -and $s.Contains(".")) {
        $s = $s.Replace(",", "")
    }

    $n = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
        return [double]$n
    }
    return $null
}

function Format-NullableFrequency {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return ([double]$Value).ToString("0.00000", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-RepeaterRowsFromJsonObject {
    param([object]$JsonObject)

    if ($null -eq $JsonObject) { return @() }
    if ($JsonObject -is [System.Array]) { return @($JsonObject) }

    foreach ($key in @("items", "data", "results", "repeaters", "channels", "list")) {
        $child = Get-NormalizedObjectValue -Object $JsonObject -Names @($key)
        if ($child -is [System.Array]) {
            return @($child)
        }
    }

    return @($JsonObject)
}

function Get-RepeaterRowsFromText {
    param(
        [string]$Text,
        [string]$SourceHint
    )

    $trimmed = $Text.TrimStart()
    $isJson = $false
    if (-not [string]::IsNullOrWhiteSpace($SourceHint) -and $SourceHint.ToLowerInvariant().EndsWith(".json")) { $isJson = $true }
    if ($trimmed.StartsWith("{") -or $trimmed.StartsWith("[")) { $isJson = $true }

    if ($isJson) {
        $jsonObj = $Text | ConvertFrom-Json -Depth 30
        return Get-RepeaterRowsFromJsonObject -JsonObject $jsonObj
    }

    $lines = @($Text -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -eq 0) { return @() }
    $header = $lines[0]
    $delimiter = ","
    if ($header.Contains(";") -and ($header.Split(";").Count -gt $header.Split(",").Count)) {
        $delimiter = ";"
    }

    return @($Text | ConvertFrom-Csv -Delimiter $delimiter)
}

function Convert-RepeaterRowsToNormalized {
    param(
        [object[]]$Rows,
        [string]$SourceLabel
    )

    $normalized = New-Object System.Collections.Generic.List[object]
    $rowIndex = 1
    foreach ($row in $Rows) {
        if ($null -eq $row) { continue }

        $name = Get-NormalizedObjectValue -Object $row -Names @(
            "name", "channel", "channelname", "callsign", "repeater", "label", "title"
        )
        $rx = Convert-ToNullableFrequency (Get-NormalizedObjectValue -Object $row -Names @(
            "rx", "rxmhz", "rxfrequency", "receive", "frequency", "freq", "output", "downlink"
        ))
        $tx = Convert-ToNullableFrequency (Get-NormalizedObjectValue -Object $row -Names @(
            "tx", "txmhz", "txfrequency", "transmit", "input", "uplink"
        ))
        $offset = Convert-ToNullableFrequency (Get-NormalizedObjectValue -Object $row -Names @(
            "offset", "offsetmhz", "shift"
        ))
        $tone = Get-NormalizedObjectValue -Object $row -Names @(
            "tone", "ctcss", "pl", "tonehz", "access", "accesstone"
        )
        $mode = Get-NormalizedObjectValue -Object $row -Names @(
            "mode", "modulation", "system", "digital", "analog"
        )
        $location = Get-NormalizedObjectValue -Object $row -Names @(
            "city", "qth", "location", "region", "state", "country"
        )
        $comment = Get-NormalizedObjectValue -Object $row -Names @(
            "comment", "comments", "note", "notes", "info", "description"
        )
        $latitude = Convert-ToNullableFrequency (Get-NormalizedObjectValue -Object $row -Names @(
            "lat", "latitude", "geo_lat", "y"
        ))
        $longitude = Convert-ToNullableFrequency (Get-NormalizedObjectValue -Object $row -Names @(
            "lon", "lng", "longitude", "geo_lon", "x"
        ))

        if ($null -eq $tx -and $null -ne $rx -and $null -ne $offset) {
            $tx = $rx + $offset
        }
        if ($null -eq $offset -and $null -ne $rx -and $null -ne $tx) {
            $offset = $tx - $rx
        }

        if ($null -eq $rx -and $null -eq $tx) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$name)) {
            $name = ("CH-{0:000}" -f $rowIndex)
        }

        $normalized.Add([pscustomobject]@{
            Name = [string]$name
            RX_MHz = Format-NullableFrequency $rx
            TX_MHz = Format-NullableFrequency $tx
            Offset_MHz = Format-NullableFrequency $offset
            Tone = [string]$tone
            Mode = [string]$mode
            Location = [string]$location
            Comment = [string]$comment
            Latitude = Format-NullableFrequency $latitude
            Longitude = Format-NullableFrequency $longitude
            Source = [string]$SourceLabel
        })
        $rowIndex++
    }

    return $normalized.ToArray()
}

function Show-TextInputDialog {
    param(
        [string]$Title,
        [string]$Prompt,
        [string]$DefaultValue = ""
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.StartPosition = "CenterParent"
    $dlg.Size = New-Object System.Drawing.Size(700, 170)
    $dlg.MinimumSize = New-Object System.Drawing.Size(560, 170)
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.KeyPreview = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Prompt
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size = New-Object System.Drawing.Size(660, 32)
    $dlg.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(12, 48)
    $txt.Size = New-Object System.Drawing.Size(660, 24)
    $txt.Text = $DefaultValue
    $txt.TabIndex = 0
    $dlg.Controls.Add($txt)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = (T "btn_ok")
    $btnOk.Location = New-Object System.Drawing.Point(500, 86)
    $btnOk.Size = New-Object System.Drawing.Size(80, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOk.TabIndex = 1
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = (T "btn_cancel")
    $btnCancel.Location = New-Object System.Drawing.Point(592, 86)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.TabIndex = 2
    $dlg.Controls.Add($btnCancel)

    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    [void]$txt.Focus()
    [void]($txt.SelectionStart = 0)
    [void]($txt.SelectionLength = $txt.TextLength)

    try {
        $result = $dlg.ShowDialog($script:form)
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $txt.Text.Trim()
        }
        return $null
    } finally {
        $dlg.Dispose()
    }
}

function Save-NormalizedRepeaterRows {
    param([object[]]$Rows)

    $targetDir = Join-Path $script:BaseDir "downloads\\channel_lists"
    if (-not (Test-Path $targetDir)) { [void](New-Item -ItemType Directory -Path $targetDir) }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = (T "cap_channels_import")
    $dlg.Filter = (T "filter_csv_save")
    $dlg.InitialDirectory = $targetDir
    $dlg.FileName = ("OpenGD77_channels_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
    $dlg.OverwritePrompt = $true

    try {
        if ($dlg.ShowDialog($script:form) -ne [System.Windows.Forms.DialogResult]::OK) {
            return $null
        }
        $Rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        return [string]$dlg.FileName
    } finally {
        $dlg.Dispose()
    }
}

function Normalize-ToneForOpenGd77Cps {
    param([string]$Tone)

    if ([string]::IsNullOrWhiteSpace($Tone)) { return "None" }
    $t = $Tone.Trim()
    if ([string]::IsNullOrWhiteSpace($t)) { return "None" }
    if ($t -match "^(?i:none|off|csq|carrier|n/a|-)$") { return "None" }
    $t = $t -replace "(?i)\s*hz", ""
    $t = $t -replace ",", "."
    return $t
}

function Get-OpenGd77ChannelTypeFromText {
    param([string]$Mode, [string]$Comment)

    $text = (([string]$Mode) + " " + ([string]$Comment)).ToLowerInvariant()
    if ($text -match "dmr|digital|d-star|dstar|nxdn|p25|ysf|fusion|m17") {
        return "Digital"
    }
    return "Analogue"
}

function Get-OpenGd77BandwidthFromText {
    param([string]$Mode, [string]$Comment, [string]$ChannelType)

    if ($ChannelType -eq "Digital") { return "" }
    $text = (([string]$Mode) + " " + ([string]$Comment)).ToLowerInvariant()
    if ($text -match "narrow|12\\.5|12,5|nfm") { return "12.5" }
    return "25"
}

function Try-ExtractOpenGd77ColorCode {
    param([string]$Mode, [string]$Comment)

    $text = (([string]$Mode) + " " + ([string]$Comment))
    $m = [regex]::Match($text, "(?i)\bcc\s*[:=]?\s*(1[0-5]|[0-9])\b")
    if ($m.Success) { return [string]$m.Groups[1].Value }
    return ""
}

function Try-ExtractOpenGd77Timeslot {
    param([string]$Mode, [string]$Comment)

    $text = (([string]$Mode) + " " + ([string]$Comment))
    $m = [regex]::Match($text, "(?i)\bts\s*[:=]?\s*([12])\b")
    if ($m.Success) { return [string]$m.Groups[1].Value }
    if ($text -match "(?i)\bts1\b") { return "1" }
    if ($text -match "(?i)\bts2\b") { return "2" }
    return ""
}

function Convert-NormalizedToOpenGd77Channels {
    param([object[]]$Rows)

    $channels = New-Object System.Collections.Generic.List[object]
    $channelNumber = 1
    foreach ($row in $Rows) {
        if ($null -eq $row) { continue }

        $rx = [string]$row.RX_MHz
        $tx = [string]$row.TX_MHz
        if ([string]::IsNullOrWhiteSpace($rx) -and [string]::IsNullOrWhiteSpace($tx)) { continue }

        $name = [string]$row.Name
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = ("CH-{0:000}" -f $channelNumber)
        }
        if ($name.Length -gt 16) {
            $name = $name.Substring(0, 16)
        }

        $mode = [string]$row.Mode
        $comment = [string]$row.Comment
        $tone = Normalize-ToneForOpenGd77Cps ([string]$row.Tone)
        $latitude = [string]$row.Latitude
        $longitude = [string]$row.Longitude
        $useLocation = "No"
        if ((-not [string]::IsNullOrWhiteSpace($latitude)) -and (-not [string]::IsNullOrWhiteSpace($longitude))) {
            $useLocation = "Yes"
        }
        $channelType = Get-OpenGd77ChannelTypeFromText -Mode $mode -Comment $comment
        $bandwidth = Get-OpenGd77BandwidthFromText -Mode $mode -Comment $comment -ChannelType $channelType

        $colorCode = ""
        $timeslot = ""
        $rxTone = $tone
        $txTone = $tone
        if ($channelType -eq "Digital") {
            $colorCode = Try-ExtractOpenGd77ColorCode -Mode $mode -Comment $comment
            if ([string]::IsNullOrWhiteSpace($colorCode)) { $colorCode = "1" }
            $timeslot = Try-ExtractOpenGd77Timeslot -Mode $mode -Comment $comment
            if ([string]::IsNullOrWhiteSpace($timeslot)) { $timeslot = "1" }
            $rxTone = "None"
            $txTone = "None"
        }

        $channels.Add([pscustomobject][ordered]@{
            "Channel Number" = $channelNumber
            "Channel Name" = $name
            "Channel Type" = $channelType
            "Rx Frequency" = $rx
            "Tx Frequency" = $tx
            "Bandwidth (kHz)" = $bandwidth
            "Colour Code" = $colorCode
            "Timeslot" = $timeslot
            "Contact" = "None"
            "TG List" = "None"
            "DMR ID" = "None"
            "TS1_TA_Tx ID" = "Off"
            "TS2_TA_Tx ID" = "Off"
            "RX Tone" = $rxTone
            "TX Tone" = $txTone
            "Squelch" = "Disabled"
            "Power" = "Master"
            "Rx Only" = "No"
            "Zone Skip" = "No"
            "All Skip" = "No"
            "TOT" = "0"
            "VOX" = "No"
            "No Beep" = "No"
            "No Eco" = "No"
            "APRS" = "No"
            "Latitude" = $latitude
            "Longitude" = $longitude
            "Use Location" = $useLocation
        })
        $channelNumber++
    }

    return $channels.ToArray()
}

function Convert-OpenGd77ChannelsToZones {
    param([object[]]$ChannelRows)

    $zones = New-Object System.Collections.Generic.List[object]
    if ($null -eq $ChannelRows -or $ChannelRows.Count -eq 0) {
        return $zones.ToArray()
    }

    $zoneCapacity = 80
    $zoneCount = [int][Math]::Ceiling($ChannelRows.Count / [double]$zoneCapacity)
    for ($z = 0; $z -lt $zoneCount; $z++) {
        $zoneName = if ($zoneCount -gt 1) { ("Imported {0:000}" -f ($z + 1)) } else { "Imported" }
        $zone = [ordered]@{
            "Zone Name" = $zoneName
        }
        for ($i = 1; $i -le $zoneCapacity; $i++) {
            $idx = ($z * $zoneCapacity) + ($i - 1)
            if ($idx -lt $ChannelRows.Count) {
                $zone["Channel$i"] = [string]$ChannelRows[$idx]."Channel Name"
            } else {
                $zone["Channel$i"] = ""
            }
        }
        $zones.Add([pscustomobject]$zone)
    }

    return $zones.ToArray()
}

function Save-OpenGd77CpsBundle {
    param([object[]]$Rows)

    $channels = Convert-NormalizedToOpenGd77Channels -Rows $Rows
    if ($channels.Count -le 0) { return $null }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $bundleDir = Join-Path $script:BaseDir ("downloads\\channel_lists\\OpenGD77_CPS_" + $timestamp)
    if (-not (Test-Path $bundleDir)) { [void](New-Item -ItemType Directory -Path $bundleDir) }

    $channelsPath = Join-Path $bundleDir "Channels.csv"
    $zonesPath = Join-Path $bundleDir "Zones.csv"
    $readmePath = Join-Path $bundleDir "README.txt"

    $channelHeaders = @(
        "Channel Number", "Channel Name", "Channel Type", "Rx Frequency", "Tx Frequency", "Bandwidth (kHz)",
        "Colour Code", "Timeslot", "Contact", "TG List", "DMR ID", "TS1_TA_Tx ID", "TS2_TA_Tx ID",
        "RX Tone", "TX Tone", "Squelch", "Power", "Rx Only", "Zone Skip", "All Skip", "TOT", "VOX",
        "No Beep", "No Eco", "APRS", "Latitude", "Longitude", "Use Location"
    )
    $channels | Select-Object -Property $channelHeaders | Export-Csv -Path $channelsPath -NoTypeInformation -Encoding UTF8

    $zones = Convert-OpenGd77ChannelsToZones -ChannelRows $channels
    $zoneHeaders = @("Zone Name")
    foreach ($i in 1..80) { $zoneHeaders += ("Channel" + $i) }
    $zones | Select-Object -Property $zoneHeaders | Export-Csv -Path $zonesPath -NoTypeInformation -Encoding UTF8

    $readme = @(
        "OpenGD77 CPS CSV bundle",
        "",
        "Generated: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
        "Entries: " + $channels.Count,
        "",
        "How to import in OpenGD77 CPS:",
        "1) Open OpenGD77 CPS",
        "2) Use File -> CSV -> Append CSV",
        "3) Select this folder and import Channels.csv + Zones.csv"
    ) -join "`r`n"
    [System.IO.File]::WriteAllText($readmePath, $readme, (New-Object System.Text.UTF8Encoding($false)))

    return [pscustomobject]@{
        BundleDir = $bundleDir
        ChannelsPath = $channelsPath
        ZonesPath = $zonesPath
        Count = $channels.Count
    }
}

function Start-ChannelsImport {
    param([switch]$FromUrl)

    if ($script:ProcessRunning) { return }

    try {
        $sourceSpec = $null
        $sourceLabel = ""
        $textData = ""

        if ($FromUrl) {
            $sourceSpec = Show-TextInputDialog -Title (T "cap_channels_url") -Prompt (T "msg_channels_url_prompt") -DefaultValue "https://"
            if ($sourceSpec -is [System.Array]) {
                $sourceSpec = ($sourceSpec | Select-Object -Last 1)
            }
            $sourceSpec = [string]$sourceSpec
            if ([string]::IsNullOrWhiteSpace($sourceSpec)) { return }
            $sourceSpec = $sourceSpec.Trim()
            $sourceLabel = $sourceSpec
            Append-Log (TF "log_channels_source_fmt" $sourceLabel)
            $response = Invoke-WebRequest -Uri $sourceSpec -UseBasicParsing -TimeoutSec 40
            $textData = [string]$response.Content
        } else {
            $openDlg = New-Object System.Windows.Forms.OpenFileDialog
            $openDlg.Title = (T "cap_channels_import")
            $openDlg.Filter = (T "filter_channels_open")
            $openDlg.Multiselect = $false
            $openDlg.InitialDirectory = (Join-Path $script:BaseDir "downloads")
            try {
                if ($openDlg.ShowDialog($script:form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
                $sourceSpec = $openDlg.FileName
            } finally {
                $openDlg.Dispose()
            }
            $sourceLabel = $sourceSpec
            Append-Log (TF "log_channels_source_fmt" $sourceLabel)
            $textData = Get-Content -Path $sourceSpec -Raw -Encoding UTF8
        }

        $rows = Get-RepeaterRowsFromText -Text $textData -SourceHint $sourceSpec
        $normalized = Convert-RepeaterRowsToNormalized -Rows $rows -SourceLabel $sourceLabel
        Append-Log (TF "log_channels_loaded_fmt" $normalized.Count)

        if ($normalized.Count -le 0) {
            [System.Windows.Forms.MessageBox]::Show(
                (T "msg_channels_no_data"),
                (T "cap_channels_import"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $savedPath = Save-NormalizedRepeaterRows -Rows $normalized
        if ([string]::IsNullOrWhiteSpace($savedPath)) { return }

        Append-Log (TF "log_channels_saved_fmt" $savedPath)
        $cpsBundlePath = ""
        try {
            $cpsBundle = Save-OpenGd77CpsBundle -Rows $normalized
            if ($cpsBundle -and -not [string]::IsNullOrWhiteSpace([string]$cpsBundle.BundleDir)) {
                $cpsBundlePath = [string]$cpsBundle.BundleDir
                Append-Log (TF "log_channels_cps_saved_fmt" $cpsBundlePath)
            }
        } catch {
            Write-DebugLog -Level "WARN" -Message ("CHANNELS_IMPORT CPS bundle generation failed: " + $_.Exception.Message)
        }

        $doneMessage = ""
        if ([string]::IsNullOrWhiteSpace($cpsBundlePath)) {
            $doneMessage = (TF "msg_channels_saved_fmt" $normalized.Count $savedPath)
        } else {
            $doneMessage = (TF "msg_channels_saved_with_cps_fmt" $normalized.Count $savedPath $cpsBundlePath)
        }
        [System.Windows.Forms.MessageBox]::Show(
            $doneMessage,
            (T "cap_channels_import"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        Write-DebugLog -Level "ERROR" -Message ("CHANNELS_IMPORT failed: " + $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            (TF "msg_channels_error_fmt" $_.Exception.Message),
            (T "cap_error"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
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
$btnCheck.Size = New-Object System.Drawing.Size(148, 32)
$btnCheck.TabIndex = 4
$form.Controls.Add($btnCheck)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = (T "btn_start")
$btnStart.Location = New-Object System.Drawing.Point(168, 162)
$btnStart.Size = New-Object System.Drawing.Size(148, 32)
$btnStart.TabIndex = 5
$form.Controls.Add($btnStart)

$btnProgramUpdate = New-Object System.Windows.Forms.Button
$btnProgramUpdate.Text = (T "btn_program_update")
$btnProgramUpdate.Location = New-Object System.Drawing.Point(324, 162)
$btnProgramUpdate.Size = New-Object System.Drawing.Size(148, 32)
$btnProgramUpdate.TabIndex = 6
$form.Controls.Add($btnProgramUpdate)

$btnChannelsFile = New-Object System.Windows.Forms.Button
$btnChannelsFile.Text = (T "btn_channels_file")
$btnChannelsFile.Location = New-Object System.Drawing.Point(480, 162)
$btnChannelsFile.Size = New-Object System.Drawing.Size(148, 32)
$btnChannelsFile.TabIndex = 7
$form.Controls.Add($btnChannelsFile)

$btnChannelsUrl = New-Object System.Windows.Forms.Button
$btnChannelsUrl.Text = (T "btn_channels_url")
$btnChannelsUrl.Location = New-Object System.Drawing.Point(636, 162)
$btnChannelsUrl.Size = New-Object System.Drawing.Size(148, 32)
$btnChannelsUrl.TabIndex = 8
$form.Controls.Add($btnChannelsUrl)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = (T "btn_close")
$btnClose.Location = New-Object System.Drawing.Point(792, 162)
$btnClose.Size = New-Object System.Drawing.Size(120, 32)
$btnClose.TabIndex = 9
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
$txtStatus.TabIndex = 10
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
$txtLastMessage.TabIndex = 11
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
$txtLog.TabIndex = 12
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
$script:btnChannelsFile = $btnChannelsFile
$script:btnChannelsUrl = $btnChannelsUrl
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
    $script:btnChannelsFile.Text = (T "btn_channels_file")
    $script:btnChannelsUrl.Text = (T "btn_channels_url")
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

$btnChannelsFile.Add_Click({
    if ($script:ProcessRunning) { return }
    Write-DebugLog "UI_CLICK: Import kanałów z pliku"
    Start-ChannelsImport
})

$btnChannelsUrl.Add_Click({
    if ($script:ProcessRunning) { return }
    Write-DebugLog "UI_CLICK: Import przemienników z internetu"
    Start-ChannelsImport -FromUrl
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
    } elseif ($e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::K) {
        $btnChannelsFile.PerformClick()
        $e.SuppressKeyPress = $true
    } elseif ($e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::U) {
        $btnChannelsUrl.PerformClick()
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
