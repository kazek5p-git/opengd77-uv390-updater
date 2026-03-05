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
$script:AppVersion = "2026.03.05.12"
$script:ProgramUpdateConfigPath = Join-Path $script:BaseDir "program_update_config.json"
$script:DefaultProgramUpdateManifestUrl = "https://kazpar.pl/opengd77-updater/latest.json"
$script:UiLanguage = "pl"
$script:SkipCloseConfirmation = $false
$script:StartupProgramUpdateTimer = $null
$script:CpsExecutableCandidates = @(
    "C:\\Program Files (x86)\\OpenGD77CPS\\OpenGD77CPS.exe",
    "C:\\Program Files\\OpenGD77CPS\\OpenGD77CPS.exe"
)
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
        btn_channels_write = "Kanały + &wgraj"
        btn_close = "&Zamknij"
        btn_ok = "OK"
        btn_cancel = "Anuluj"
        hints = "Skróty: Alt+S Sprawdź, Alt+T Start, Alt+A Aktualizuj program, Alt+K Kanały z pliku, Alt+U Przemienniki z internetu, Alt+W Kanały i wgraj, Alt+L Log, Alt+D Timeout, Alt+F4 Zamknij, F1 Pomoc."
        status_group = "Status"
        status_label = "Bieżący status:"
        status_ready = "Gotowe do uruchomienia"
        last_message_label = "Ostatni komunikat:"
        status_no_messages = "Brak komunikatów."
        log_group = "Log"
        status_busy = "Trwa operacja..."
        status_channels_write = "Import i zapis kanałów do radia..."
        status_channels_write_error = "Błąd zapisu kanałów do radia"
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
        cap_channels_write = "Import i zapis kanałów"
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
        msg_help = "Sterowanie:`n- Alt+S: Sprawdź wersję`n- Alt+T: Start aktualizacji`n- Alt+A: Aktualizuj program`n- Alt+K: Kanały z pliku`n- Alt+U: Przemienniki z internetu`n- Alt+W: Kanały i wgraj`n- Alt+L: Fokus log`n- Alt+D: Fokus timeout`n- Alt+F4: Zamknij"
        msg_channels_url_prompt = "Wklej URL do pliku CSV, JSON lub ICF (ICOM) z kanałami/przemiennikami:"
        msg_channels_no_data = "Nie znaleziono wpisów kanałów/przemienników."
        msg_channels_saved_fmt = "Zapisano {0} wpisów do: {1}"
        msg_channels_saved_with_cps_fmt = "Zapisano {0} wpisów do: {1}`nUtworzono pakiet OpenGD77 CPS: {2}`nW CPS użyj File -> CSV -> Import CSV."
        msg_channels_write_confirm = "Program zaimportuje kanały i zapisze je do radia przez OpenGD77 CPS.`nUpewnij się, że radio jest podłączone i włączone.`n`nKontynuować?"
        msg_channels_write_done_fmt = "Zapis kanałów zakończony.`nKanały: {0}`nStrefy: {1}`nWeryfikacja eksportu: {2} kanałów.`nFolder roboczy: {3}"
        msg_channels_write_error_fmt = "Błąd automatycznego zapisu kanałów do radia:`n{0}"
        msg_cps_missing_fmt = "Nie znaleziono OpenGD77 CPS. Zainstaluj CPS i spróbuj ponownie.`nSzukane lokalizacje:`n{0}"
        msg_cps_window_not_ready = "OpenGD77 CPS uruchomił się, ale okno główne nie jest gotowe."
        msg_channels_verify_failed_fmt = "Weryfikacja po zapisie nie powiodła się. Oczekiwano {0} kanałów, odczytano {1}."
        msg_channels_verify_missing_names_fmt = "Brakujące nazwy kanałów po weryfikacji: {0}."
        msg_channels_error_fmt = "Błąd importu kanałów/przemienników:`n{0}"
        msg_cannot_close_during_update = "Nie można zamknąć programu podczas aktualizacji."
        msg_confirm_close = "Czy na pewno chcesz zamknąć program?"
        filter_channels_open = "Kanały/Przemienniki (*.csv;*.json;*.icf)|*.csv;*.json;*.icf|CSV (*.csv)|*.csv|JSON (*.json)|*.json|ICOM ICF (*.icf)|*.icf|Wszystkie pliki (*.*)|*.*"
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
        log_channels_write_begin = "Start trybu: import i zapis kanałów do radia."
        log_channels_write_canceled = "Anulowano zapis kanałów do radia."
        log_channels_verify_ok_fmt = "Weryfikacja CPS OK. Odczytano kanałów: {0}, brakujących nazw: {1}."
        log_cps_action_fmt = "CPS akcja: {0}"
        log_cps_dialog_fmt = "CPS dialog: {0}"
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
        btn_channels_write = "Channels + &write"
        btn_close = "&Close"
        btn_ok = "OK"
        btn_cancel = "Cancel"
        hints = "Shortcuts: Alt+S Check, Alt+T Start, Alt+A Update app, Alt+K Channels from file, Alt+U Repeaters from internet, Alt+W Channels and write, Alt+L Log, Alt+D Timeout, Alt+F4 Close, F1 Help."
        status_group = "Status"
        status_label = "Current status:"
        status_ready = "Ready"
        last_message_label = "Last message:"
        status_no_messages = "No messages."
        log_group = "Log"
        status_busy = "Operation in progress..."
        status_channels_write = "Importing and writing channels to radio..."
        status_channels_write_error = "Channel write error"
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
        cap_channels_write = "Channels import and write"
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
        msg_help = "Controls:`n- Alt+S: Check version`n- Alt+T: Start update`n- Alt+A: Update app`n- Alt+K: Channels from file`n- Alt+U: Repeaters from internet`n- Alt+W: Channels and write`n- Alt+L: Log focus`n- Alt+D: Timeout focus`n- Alt+F4: Close"
        msg_channels_url_prompt = "Paste URL to CSV, JSON or ICF (ICOM) file with channels/repeaters:"
        msg_channels_no_data = "No channel/repeater entries found."
        msg_channels_saved_fmt = "Saved {0} entries to: {1}"
        msg_channels_saved_with_cps_fmt = "Saved {0} entries to: {1}`nCreated OpenGD77 CPS bundle: {2}`nIn CPS use File -> CSV -> Import CSV."
        msg_channels_write_confirm = "The app will import channels and write them to radio using OpenGD77 CPS.`nMake sure the radio is connected and powered on.`n`nContinue?"
        msg_channels_write_done_fmt = "Channel write completed.`nChannels: {0}`nZones: {1}`nExport verification: {2} channels.`nWorking folder: {3}"
        msg_channels_write_error_fmt = "Automatic channel write error:`n{0}"
        msg_cps_missing_fmt = "OpenGD77 CPS not found. Install CPS and try again.`nChecked locations:`n{0}"
        msg_cps_window_not_ready = "OpenGD77 CPS started but main window is not ready."
        msg_channels_verify_failed_fmt = "Verification after write failed. Expected {0} channels, got {1}."
        msg_channels_verify_missing_names_fmt = "Missing channel names after verification: {0}."
        msg_channels_error_fmt = "Channel/repeater import error:`n{0}"
        msg_cannot_close_during_update = "You cannot close the app during update."
        msg_confirm_close = "Are you sure you want to close the app?"
        filter_channels_open = "Channels/Repeaters (*.csv;*.json;*.icf)|*.csv;*.json;*.icf|CSV (*.csv)|*.csv|JSON (*.json)|*.json|ICOM ICF (*.icf)|*.icf|All files (*.*)|*.*"
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
        log_channels_write_begin = "Mode started: import and write channels to radio."
        log_channels_write_canceled = "Channel write to radio canceled."
        log_channels_verify_ok_fmt = "CPS verification OK. Exported channels: {0}, missing names: {1}."
        log_cps_action_fmt = "CPS action: {0}"
        log_cps_dialog_fmt = "CPS dialog: {0}"
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
    if ($script:btnChannelsWrite) { $script:btnChannelsWrite.Enabled = -not $Busy }
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

function Get-IcomIcfByteArray {
    param([string]$Text)

    $chunks = New-Object System.Collections.Generic.List[object]
    foreach ($rawLine in ($Text -split "`r?`n")) {
        $line = ([string]$rawLine).Trim().ToUpperInvariant()
        if ($line -notmatch "^[0-9A-F]{6}[0-9A-F]+$") { continue }
        if ($line.Length -lt 6) { continue }

        $offset = [Convert]::ToInt32($line.Substring(0, 4), 16)
        $declaredLength = [Convert]::ToInt32($line.Substring(4, 2), 16)
        $hexData = $line.Substring(6)
        if ($hexData.Length -ne ($declaredLength * 2)) { continue }

        $chunks.Add([pscustomobject]@{
            Offset = $offset
            Length = $declaredLength
            HexData = $hexData
        })
    }

    if ($chunks.Count -eq 0) {
        throw "Nie znaleziono blokow danych ICF."
    }

    $maxSize = 0
    foreach ($chunk in $chunks) {
        $end = [int]$chunk.Offset + [int]$chunk.Length
        if ($end -gt $maxSize) { $maxSize = $end }
    }
    if ($maxSize -le 0) {
        throw "Nieprawidlowy rozmiar danych ICF."
    }

    $buffer = New-Object byte[] $maxSize
    foreach ($chunk in $chunks) {
        $start = [int]$chunk.Offset
        $hex = [string]$chunk.HexData
        for ($i = 0; $i -lt $hex.Length; $i += 2) {
            $buffer[$start + ($i / 2)] = [Convert]::ToByte($hex.Substring($i, 2), 16)
        }
    }

    return $buffer
}

function Get-UInt16BE {
    param(
        [byte[]]$Buffer,
        [int]$Offset
    )
    return ([int]$Buffer[$Offset] * 256) + [int]$Buffer[$Offset + 1]
}

function Get-UInt24BE {
    param(
        [byte[]]$Buffer,
        [int]$Offset
    )
    return ([int]$Buffer[$Offset] * 65536) + ([int]$Buffer[$Offset + 1] * 256) + [int]$Buffer[$Offset + 2]
}

function Get-Icom2730RowsFromIcfText {
    param(
        [string]$Text,
        [string]$SourceLabel
    )

    if ($Text -notmatch "(?im)^#Comment=IC-2730") {
        throw "Ten konwerter ICF obsluguje aktualnie tylko modele ICOM IC-2730."
    }

    $buffer = Get-IcomIcfByteArray -Text $Text
    if ($buffer.Length -lt 17088) {
        throw "Plik ICF jest za krotki albo nieobslugiwany."
    }

    $tones = @(
        67.0, 69.3, 71.9, 74.4, 77.0, 79.7, 82.5, 85.4, 88.5, 91.5,
        94.8, 97.4, 100.0, 103.5, 107.2, 110.9, 114.8, 118.8, 123.0, 127.3,
        131.8, 136.5, 141.3, 146.2, 151.4, 156.7, 159.8, 162.2, 165.5, 167.9,
        171.3, 173.8, 177.3, 179.9, 183.5, 186.2, 189.9, 192.8, 196.6, 199.5,
        203.5, 206.5, 210.7, 218.1, 225.7, 229.1, 233.6, 241.8, 250.3, 254.1
    )
    $dtcsCodes = @(
        23, 25, 26, 31, 32, 36, 43, 47, 51, 53, 54, 65, 71, 72, 73, 74,
        114, 115, 116, 122, 125, 131, 132, 134, 143, 145, 152, 155, 156, 162,
        165, 172, 174, 205, 212, 223, 225, 226, 243, 244, 245, 246, 251, 252, 255,
        261, 263, 265, 266, 271, 274, 306, 311, 315, 325, 331, 332, 343, 346, 351,
        356, 364, 365, 371, 411, 412, 413, 423, 431, 432, 445, 446, 452, 454, 455,
        462, 464, 465, 466, 503, 506, 516, 523, 526, 532, 546, 565, 606, 612, 624,
        627, 631, 632, 645, 654, 662, 664, 703, 712, 723, 731, 732, 734, 743, 754
    )
    $duplexNames = @("", "-", "+", "split", "off")
    $tmodeNames = @("", "Tone", "TSQL", "DTCS", "TSQL-R", "DTCS-R")
    $dtcsPolarityNames = @("NN", "NR", "RN", "RR")
    $modeNames = @("FM", "NFM")
    $tuneSteps = @("5.0", "6.25", "10.0", "12.5", "15.0", "20.0", "25.0", "30.0", "50.0")

    $usedFlagsOffset = 0x42c0
    $usedFlagsSize = 125
    if (($usedFlagsOffset + $usedFlagsSize) -gt $buffer.Length) {
        throw "Plik ICF nie zawiera mapy zajetych kanalow."
    }

    $rows = New-Object System.Collections.Generic.List[object]
    for ($channelNumber = 0; $channelNumber -lt 1002; $channelNumber++) {
        $recordOffset = $channelNumber * 17
        if (($recordOffset + 17) -gt $buffer.Length) { break }

        if ($channelNumber -lt 1000) {
            $usedByte = $buffer[$usedFlagsOffset + [int]($channelNumber / 8)]
            $usedBit = (1 -shl ($channelNumber % 8))
            if (($usedByte -band $usedBit) -ne 0) { continue }
        }

        $freqPacked = Get-UInt24BE -Buffer $buffer -Offset $recordOffset
        $freqFlags = [int]($freqPacked -shr 18)
        $freqRaw = [int]($freqPacked -band 0x3FFFF)
        if ($freqRaw -le 0) { continue }

        $freqMultiplierHz = if (($freqFlags -band 0x08) -ne 0) { 6250.0 } else { 5000.0 }
        $offsetMultiplierHz = if (($freqFlags -band 0x01) -ne 0) { 6250.0 } else { 5000.0 }
        $rxHz = [double]$freqRaw * $freqMultiplierHz
        $offsetRaw = [double](Get-UInt16BE -Buffer $buffer -Offset ($recordOffset + 3))
        $offsetHz = $offsetRaw * $offsetMultiplierHz

        $tuneStepMode = [int]$buffer[$recordOffset + 5]
        $modeIndex = [int](($tuneStepMode -shr 4) -band 0x01)
        $stepIndex = [int]($tuneStepMode -band 0x0F)

        $rtoneIndex = [int]$buffer[$recordOffset + 6]
        $ctoneIndex = [int]$buffer[$recordOffset + 7]
        $dtcsIndex = [int]$buffer[$recordOffset + 9]
        $tmodeDuplexPol = [int]$buffer[$recordOffset + 10]
        $tmodeIndex = [int](($tmodeDuplexPol -shr 5) -band 0x07)
        $duplexIndex = [int](($tmodeDuplexPol -shr 2) -band 0x07)
        $dtcsPolIndex = [int]($tmodeDuplexPol -band 0x03)

        $nameBytes = $buffer[($recordOffset + 11)..($recordOffset + 16)]
        $name = [System.Text.Encoding]::ASCII.GetString($nameBytes).Trim([char]0, [char]32)
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = ("CH-{0:000}" -f ($channelNumber + 1))
        }

        $duplex = if ($duplexIndex -lt $duplexNames.Count) { $duplexNames[$duplexIndex] } else { "" }
        $txHz = $null
        if ($duplex -eq "-") {
            $txHz = $rxHz - $offsetHz
        } elseif ($duplex -eq "+") {
            $txHz = $rxHz + $offsetHz
        } elseif ($duplex -eq "split") {
            $txHz = $offsetHz
        } elseif ($duplex -eq "off") {
            $txHz = $null
        } else {
            $txHz = $rxHz
        }

        $toneText = ""
        $tmode = if ($tmodeIndex -lt $tmodeNames.Count) { $tmodeNames[$tmodeIndex] } else { "" }
        if ($tmode -eq "Tone") {
            if ($rtoneIndex -ge 0 -and $rtoneIndex -lt $tones.Count) {
                $toneText = [string]$tones[$rtoneIndex]
            }
        } elseif ($tmode -eq "TSQL" -or $tmode -eq "TSQL-R") {
            if ($ctoneIndex -ge 0 -and $ctoneIndex -lt $tones.Count) {
                $toneText = [string]$tones[$ctoneIndex]
            }
        } elseif ($tmode -eq "DTCS" -or $tmode -eq "DTCS-R") {
            $dtcsCode = if ($dtcsIndex -ge 0 -and $dtcsIndex -lt $dtcsCodes.Count) { [string]$dtcsCodes[$dtcsIndex] } else { [string]$dtcsIndex }
            $dtcsPol = if ($dtcsPolIndex -lt $dtcsPolarityNames.Count) { $dtcsPolarityNames[$dtcsPolIndex] } else { "NN" }
            $toneText = ("DCS " + $dtcsCode + " " + $dtcsPol)
        }

        $mode = if ($modeIndex -lt $modeNames.Count) { $modeNames[$modeIndex] } else { "FM" }
        $step = if ($stepIndex -lt $tuneSteps.Count) { $tuneSteps[$stepIndex] } else { "" }
        $commentParts = New-Object System.Collections.Generic.List[string]
        if (-not [string]::IsNullOrWhiteSpace($tmode)) { $commentParts.Add("tmode=" + $tmode) }
        if (-not [string]::IsNullOrWhiteSpace($duplex)) { $commentParts.Add("duplex=" + $duplex) }
        if (-not [string]::IsNullOrWhiteSpace($step)) { $commentParts.Add("step=" + $step + "k") }
        $comment = [string]::Join("; ", $commentParts.ToArray())

        $rxMHz = ($rxHz / 1000000.0).ToString("0.00000", [System.Globalization.CultureInfo]::InvariantCulture)
        $txMHz = ""
        if ($null -ne $txHz -and $txHz -gt 0) {
            $txMHz = ($txHz / 1000000.0).ToString("0.00000", [System.Globalization.CultureInfo]::InvariantCulture)
        }
        $offsetMHz = ""
        if ($null -ne $txHz -and $txHz -gt 0) {
            $offsetMHz = (($txHz - $rxHz) / 1000000.0).ToString("0.00000", [System.Globalization.CultureInfo]::InvariantCulture)
        }

        $rows.Add([pscustomobject]@{
            name = [string]$name
            rxmhz = [string]$rxMHz
            txmhz = [string]$txMHz
            offsetmhz = [string]$offsetMHz
            tone = [string]$toneText
            mode = [string]$mode
            comment = [string]$comment
            location = ""
            source = [string]$SourceLabel
        })
    }

    return $rows.ToArray()
}

function Get-RepeaterRowsFromText {
    param(
        [string]$Text,
        [string]$SourceHint
    )

    $trimmed = $Text.TrimStart()
    $isJson = $false
    $isIcf = $false
    if (-not [string]::IsNullOrWhiteSpace($SourceHint) -and $SourceHint.ToLowerInvariant().EndsWith(".icf")) { $isIcf = $true }
    if (-not $isIcf -and $trimmed -match "^[0-9A-F]{8}\s*$" -and $Text -match "(?m)^#COMMENT=IC-") { $isIcf = $true }
    if (-not [string]::IsNullOrWhiteSpace($SourceHint) -and $SourceHint.ToLowerInvariant().EndsWith(".json")) { $isJson = $true }
    if ($trimmed.StartsWith("{") -or $trimmed.StartsWith("[")) { $isJson = $true }

    if ($isIcf) {
        return Get-Icom2730RowsFromIcfText -Text $Text -SourceLabel $SourceHint
    }

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

function Save-NormalizedRepeaterRowsAuto {
    param([object[]]$Rows)

    $targetDir = Join-Path $script:BaseDir "downloads\\channel_lists"
    if (-not (Test-Path $targetDir)) { [void](New-Item -ItemType Directory -Path $targetDir) }

    $path = Join-Path $targetDir ("OpenGD77_channels_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
    $Rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    return [string]$path
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
    return Save-OpenGd77CpsBundleCompatible -Rows $Rows
}

function Get-OpenGd77CsvFormat {
    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    $separator = [string]$culture.TextInfo.ListSeparator
    if ([string]::IsNullOrWhiteSpace($separator)) { $separator = ";" }
    $decimal = [string]$culture.NumberFormat.NumberDecimalSeparator
    if ([string]::IsNullOrWhiteSpace($decimal)) { $decimal = "." }

    return [pscustomobject]@{
        Separator = $separator
        DecimalSeparator = $decimal
    }
}

function Convert-ToOpenGd77CpsFrequencyField {
    param(
        [string]$Value,
        [string]$DecimalSeparator = "."
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    $num = 0.0
    if ([double]::TryParse($Value, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
        $normalized = $num.ToString("0.00000", [System.Globalization.CultureInfo]::InvariantCulture)
    } else {
        $normalized = [string]$Value
    }

    if ($DecimalSeparator -eq ",") {
        $normalized = $normalized.Replace(".", ",")
    } else {
        $normalized = $normalized.Replace(",", ".")
    }
    return ("`t" + $normalized)
}

function Convert-ToOpenGd77CsvTonePair {
    param(
        [string]$RxToneRaw,
        [string]$TxToneRaw
    )

    $normalize = {
        param([string]$raw)
        if ([string]::IsNullOrWhiteSpace($raw)) { return "" }
        $t = $raw.Trim()
        if ([string]::IsNullOrWhiteSpace($t)) { return "" }
        if ($t -match "^(?i:none|off|disabled|carrier|csq|n/a|-)$") { return "" }
        $t = $t -replace "(?i)\s*hz", ""
        $t = $t.Replace(",", ".")

        # Accept DCS in forms like: DCS 114 NN, D114N, D114I
        $mDcsSplit = [regex]::Match($t, "(?i)^(?:DCS|D)\s*([0-9]{2,3})\s*([NRI]{1,2})?$")
        if ($mDcsSplit.Success) {
            $code = $mDcsSplit.Groups[1].Value.PadLeft(3, "0")
            $pol = $mDcsSplit.Groups[2].Value.ToUpperInvariant()
            if ([string]::IsNullOrWhiteSpace($pol)) { $pol = "N" }
            if ($pol.StartsWith("R") -or $pol.StartsWith("I")) {
                return ("D" + $code + "I")
            }
            return ("D" + $code + "N")
        }

        $mCtcss = [regex]::Match($t, "^([0-9]{2,3}(?:\\.[0-9]{1,2})?)$")
        if ($mCtcss.Success) {
            return $mCtcss.Groups[1].Value
        }
        return ""
    }

    $rxTone = & $normalize $RxToneRaw
    $txTone = & $normalize $TxToneRaw

    # Preserve split DCS polarity if source has NN/NR/RN/RR form.
    $pairSource = ""
    if (-not [string]::IsNullOrWhiteSpace($RxToneRaw)) {
        $pairSource = [string]$RxToneRaw
    } elseif (-not [string]::IsNullOrWhiteSpace($TxToneRaw)) {
        $pairSource = [string]$TxToneRaw
    }
    if (-not [string]::IsNullOrWhiteSpace($pairSource)) {
        $mPair = [regex]::Match($pairSource.Trim(), "(?i)^(?:DCS|D)\\s*([0-9]{2,3})\\s*([NR]{2})$")
        if ($mPair.Success) {
            $code = $mPair.Groups[1].Value.PadLeft(3, "0")
            $pair = $mPair.Groups[2].Value.ToUpperInvariant()
            $rxTone = "D" + $code + ($(if ($pair[0] -eq "R") { "I" } else { "N" }))
            $txTone = "D" + $code + ($(if ($pair[1] -eq "R") { "I" } else { "N" }))
        }
    }

    return [pscustomobject]@{
        RxTone = $rxTone
        TxTone = $txTone
    }
}

function Get-OpenGd77UniqueChannelName {
    param(
        [string]$BaseName,
        [hashtable]$UsedNames
    )

    $name = if ([string]::IsNullOrWhiteSpace($BaseName)) { "CH" } else { $BaseName.Trim() }
    if ($name.Length -gt 16) { $name = $name.Substring(0, 16) }
    $candidate = $name
    $index = 2
    while ($UsedNames.ContainsKey($candidate)) {
        $suffix = "-" + $index
        $prefixLength = 16 - $suffix.Length
        if ($prefixLength -lt 1) { $prefixLength = 1 }
        $baseCut = $name
        if ($baseCut.Length -gt $prefixLength) {
            $baseCut = $baseCut.Substring(0, $prefixLength)
        }
        $candidate = $baseCut + $suffix
        $index++
    }
    $UsedNames[$candidate] = $true
    return $candidate
}

function Save-OpenGd77CpsBundleCompatible {
    param([object[]]$Rows)

    $channels = Convert-NormalizedToOpenGd77Channels -Rows $Rows
    if ($channels.Count -le 0) { return $null }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $bundleDir = Join-Path $script:BaseDir ("downloads\\channel_lists\\OpenGD77_CPS_AUTO_" + $timestamp)
    if (-not (Test-Path $bundleDir)) { [void](New-Item -ItemType Directory -Path $bundleDir) }

    $channelsPath = Join-Path $bundleDir "Channels.csv"
    $zonesPath = Join-Path $bundleDir "Zones.csv"
    $verifyDir = Join-Path $bundleDir "_verify_export"
    if (-not (Test-Path $verifyDir)) { [void](New-Item -ItemType Directory -Path $verifyDir) }

    $csvFmt = Get-OpenGd77CsvFormat
    $sep = [string]$csvFmt.Separator
    $decimalSep = [string]$csvFmt.DecimalSeparator

    # Keep full CPS CSV schema (28 columns), matching export/import expectations.
    $header = @(
        "Channel Number", "Channel Name", "Channel Type", "Rx Frequency", "Tx Frequency", "Bandwidth (kHz)",
        "Colour Code", "Timeslot", "Contact", "TG List", "DMR ID", "TS1_TA_Tx ID", "TS2_TA_Tx ID",
        "RX Tone", "TX Tone", "Squelch", "Power", "Rx Only", "Zone Skip", "All Skip", "TOT", "VOX",
        "No Beep", "No Eco", "APRS", "Latitude", "Longitude", "Use Location"
    )
    $channelLines = New-Object System.Collections.Generic.List[string]
    $channelLines.Add(($header -join $sep)) | Out-Null

    $usedNames = @{}
    $finalNames = New-Object System.Collections.Generic.List[string]
    $channelNumber = 1
    foreach ($row in $channels) {
        if ($null -eq $row) { continue }

        $name = Get-OpenGd77UniqueChannelName -BaseName ([string]$row."Channel Name") -UsedNames $usedNames
        $finalNames.Add($name) | Out-Null

        $isDigital = (([string]$row."Channel Type") -match "Digital")
        $rxRaw = [string]$row."Rx Frequency"
        $txRaw = [string]$row."Tx Frequency"
        if ([string]::IsNullOrWhiteSpace($txRaw)) { $txRaw = $rxRaw }

        $channelType = if ($isDigital) { "Digital" } else { "Analogue" }
        $bandwidth = if ($isDigital) { "" } else { [string]$row."Bandwidth (kHz)" }
        if ((-not $isDigital) -and [string]::IsNullOrWhiteSpace($bandwidth)) { $bandwidth = "12.5" }
        if ($decimalSep -eq ",") {
            $bandwidth = ($bandwidth -replace "\.", ",")
        } else {
            $bandwidth = ($bandwidth -replace ",", ".")
        }

        $colorCode = if ($isDigital) { [string]$row."Colour Code" } else { "" }
        if ($isDigital -and [string]::IsNullOrWhiteSpace($colorCode)) { $colorCode = "1" }
        $timeslot = if ($isDigital) { [string]$row."Timeslot" } else { "" }
        if ($isDigital -and [string]::IsNullOrWhiteSpace($timeslot)) { $timeslot = "1" }

        $contact = "None"
        $tgList = if ($isDigital) { "Brandmeister" } else { "None" }
        $dmrId = "None"
        $ts1TaTx = "Off"
        $ts2TaTx = "Off"

        $tonePair = Convert-ToOpenGd77CsvTonePair -RxToneRaw ([string]$row."RX Tone") -TxToneRaw ([string]$row."TX Tone")
        $rxTone = if ($isDigital) { "None" } else { [string]$tonePair.RxTone }
        $txTone = if ($isDigital) { "None" } else { [string]$tonePair.TxTone }
        if ([string]::IsNullOrWhiteSpace($rxTone)) { $rxTone = "None" }
        if ([string]::IsNullOrWhiteSpace($txTone)) { $txTone = "None" }
        $squelch = "Disabled"

        $line = @(
            [string]$channelNumber,
            $name,
            $channelType,
            (Convert-ToOpenGd77CpsFrequencyField -Value $rxRaw -DecimalSeparator $decimalSep),
            (Convert-ToOpenGd77CpsFrequencyField -Value $txRaw -DecimalSeparator $decimalSep),
            $bandwidth,
            $colorCode,
            $timeslot,
            $contact,
            $tgList,
            $dmrId,
            $ts1TaTx,
            $ts2TaTx,
            $rxTone,
            $txTone,
            $squelch,
            "Master",
            "No",
            "No",
            "No",
            "0",
            "No",
            "No",
            "No",
            "No",
            "",
            "",
            "No"
        ) -join $sep
        $channelLines.Add($line) | Out-Null
        $channelNumber++
    }

    [System.IO.File]::WriteAllLines($channelsPath, $channelLines, (New-Object System.Text.UTF8Encoding($false)))

    $zoneHeaders = @("Zone Name") + (1..80 | ForEach-Object { "Channel$_" })
    $zoneLines = New-Object System.Collections.Generic.List[string]
    $zoneLines.Add(($zoneHeaders -join $sep)) | Out-Null

    $zoneCount = [int][Math]::Ceiling($finalNames.Count / 80.0)
    for ($z = 0; $z -lt $zoneCount; $z++) {
        $vals = New-Object System.Collections.Generic.List[string]
        $vals.Add(("Imported {0:000}" -f ($z + 1))) | Out-Null
        for ($k = 0; $k -lt 80; $k++) {
            $index = ($z * 80) + $k
            if ($index -lt $finalNames.Count) {
                $vals.Add([string]$finalNames[$index]) | Out-Null
            } else {
                $vals.Add("") | Out-Null
            }
        }
        $zoneLines.Add(($vals -join $sep)) | Out-Null
    }

    [System.IO.File]::WriteAllLines($zonesPath, $zoneLines, (New-Object System.Text.UTF8Encoding($false)))

    return [pscustomobject]@{
        BundleDir = $bundleDir
        ChannelsPath = $channelsPath
        ZonesPath = $zonesPath
        VerifyDir = $verifyDir
        Count = $finalNames.Count
        ZoneCount = $zoneCount
        ChannelNames = $finalNames.ToArray()
    }
}

function Ensure-OpenGd77CpsWinApi {
    if ("OpenGd77CpsWinApi" -as [type]) { return }

    Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class OpenGd77CpsWinApi
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    public delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr hwndCtl);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern bool SetWindowText(IntPtr hWnd, string lpString);
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;
    public const uint BM_CLICK = 0x00F5;
    public const uint WM_COMMAND = 0x0111;
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP = 0x0101;
    public const uint CB_GETCOUNT = 0x0146;
    public const uint CB_GETCURSEL = 0x0147;
    public const uint CB_GETLBTEXT = 0x0148;
    public const uint CB_GETLBTEXTLEN = 0x0149;
    public const uint CB_SETCURSEL = 0x014E;
    public const int VK_RETURN = 0x0D;
}
"@
}

function Get-OpenGd77CpsExecutablePath {
    foreach ($candidate in $script:CpsExecutableCandidates) {
        if (Test-Path $candidate) { return $candidate }
    }
    return $null
}

function Get-OpenGd77CpsSession {
    $proc = Get-Process OpenGD77CPS -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $proc) {
        $exe = Get-OpenGd77CpsExecutablePath
        if ([string]::IsNullOrWhiteSpace($exe)) {
            throw (TF "msg_cps_missing_fmt" ([string]::Join("`n", $script:CpsExecutableCandidates)))
        }
        Start-Process -FilePath $exe | Out-Null
        Start-Sleep -Milliseconds 1200
    }

    $deadline = (Get-Date).AddSeconds(18)
    while ((Get-Date) -lt $deadline) {
        $proc = Get-Process OpenGD77CPS -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($proc) {
            $proc.Refresh()
            if ($proc.MainWindowHandle -ne 0) {
                return [pscustomobject]@{
                    Process = $proc
                    ProcessId = [uint32]$proc.Id
                    MainWindowHandle = [IntPtr]$proc.MainWindowHandle
                }
            }
        }
        Start-Sleep -Milliseconds 250
        [System.Windows.Forms.Application]::DoEvents()
    }

    throw (T "msg_cps_window_not_ready")
}

function Set-OpenGd77CpsForeground {
    param([IntPtr]$MainWindowHandle)
    [void][OpenGd77CpsWinApi]::SetForegroundWindow($MainWindowHandle)
    Start-Sleep -Milliseconds 220
    [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-OpenGd77CpsAbsoluteClick {
    param(
        [int]$X,
        [int]$Y
    )
    [void][OpenGd77CpsWinApi]::SetCursorPos($X, $Y)
    Start-Sleep -Milliseconds 60
    [OpenGd77CpsWinApi]::mouse_event([OpenGd77CpsWinApi]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 25
    [OpenGd77CpsWinApi]::mouse_event([OpenGd77CpsWinApi]::MOUSEEVENTF_LEFTUP, 0, 0, 0, [UIntPtr]::Zero)
}

function Invoke-OpenGd77CpsMainToolbarClick {
    param(
        [IntPtr]$MainWindowHandle,
        [int]$OffsetX,
        [int]$OffsetY
    )

    $rect = New-Object OpenGd77CpsWinApi+RECT
    if (-not [OpenGd77CpsWinApi]::GetWindowRect($MainWindowHandle, [ref]$rect)) {
        return $false
    }

    $targetX = [int]($rect.Left + $OffsetX)
    $targetY = [int]($rect.Top + $OffsetY)
    Invoke-OpenGd77CpsAbsoluteClick -X $targetX -Y $targetY
    return $true
}

function Get-OpenGd77CpsDialogs {
    param([uint32]$ProcessId)

    $list = New-Object System.Collections.Generic.List[object]
    $callback = [OpenGd77CpsWinApi+EnumWindowsProc]{
        param($hwnd, $lParam)
        [uint32]$currentPid = 0
        [void][OpenGd77CpsWinApi]::GetWindowThreadProcessId($hwnd, [ref]$currentPid)
        if ($currentPid -ne $ProcessId) { return $true }

        $cls = New-Object System.Text.StringBuilder 128
        [void][OpenGd77CpsWinApi]::GetClassName($hwnd, $cls, 128)
        $length = [OpenGd77CpsWinApi]::GetWindowTextLength($hwnd)
        $title = New-Object System.Text.StringBuilder ([Math]::Max(1, $length + 1))
        [void][OpenGd77CpsWinApi]::GetWindowText($hwnd, $title, $title.Capacity)
        $list.Add([pscustomobject]@{
            Hwnd = [IntPtr]$hwnd
            Class = $cls.ToString()
            Title = $title.ToString()
        }) | Out-Null
        return $true
    }
    [void][OpenGd77CpsWinApi]::EnumWindows($callback, [IntPtr]::Zero)
    return $list.ToArray()
}

function Get-OpenGd77CpsDialogRows {
    param([IntPtr]$DialogHandle)

    $rows = New-Object System.Collections.Generic.List[object]
    $callback = [OpenGd77CpsWinApi+EnumChildProc]{
        param($childHandle, $lParam)
        $className = New-Object System.Text.StringBuilder 64
        [void][OpenGd77CpsWinApi]::GetClassName($childHandle, $className, 64)
        $length = [OpenGd77CpsWinApi]::GetWindowTextLength($childHandle)
        $text = New-Object System.Text.StringBuilder ([Math]::Max(1, $length + 1))
        [void][OpenGd77CpsWinApi]::GetWindowText($childHandle, $text, $text.Capacity)
        $rows.Add([pscustomobject]@{
            Hwnd = [IntPtr]$childHandle
            Class = $className.ToString()
            Id = [OpenGd77CpsWinApi]::GetDlgCtrlID($childHandle)
            Text = $text.ToString()
        }) | Out-Null
        return $true
    }
    [void][OpenGd77CpsWinApi]::EnumChildWindows($DialogHandle, $callback, [IntPtr]::Zero)
    return $rows.ToArray()
}

function Invoke-OpenGd77CpsDialogButton {
    param(
        [object[]]$Rows,
        [int[]]$Ids,
        [string]$PreferredText = "OK"
    )

    $button = $Rows | Where-Object {
        ([string]$_.Class -match "(?i)button") -and ([string]$_.Text -ieq $PreferredText) -and ($Ids -contains [int]$_.Id)
    } | Select-Object -First 1
    if (-not $button) {
        $button = $Rows | Where-Object {
            ([string]$_.Class -match "(?i)button") -and ($Ids -contains [int]$_.Id)
        } | Select-Object -First 1
    }
    if (-not $button) { return $false }

    [void][OpenGd77CpsWinApi]::SendMessage($button.Hwnd, [uint32][OpenGd77CpsWinApi]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
    return $true
}

function Invoke-OpenGd77CpsBrowseFolderDialog {
    param(
        [object[]]$Rows,
        [string]$FolderPath
    )

    $edit = $Rows | Where-Object { $_.Class -eq "Edit" -and [int]$_.Id -eq 14148 } | Select-Object -First 1
    if (-not $edit) { return $false }

    [void][OpenGd77CpsWinApi]::SetWindowText($edit.Hwnd, $FolderPath)
    [void][OpenGd77CpsWinApi]::SendMessage($edit.Hwnd, [uint32][OpenGd77CpsWinApi]::WM_KEYDOWN, [IntPtr][OpenGd77CpsWinApi]::VK_RETURN, [IntPtr]::Zero)
    [void][OpenGd77CpsWinApi]::SendMessage($edit.Hwnd, [uint32][OpenGd77CpsWinApi]::WM_KEYUP, [IntPtr][OpenGd77CpsWinApi]::VK_RETURN, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 150
    return (Invoke-OpenGd77CpsDialogButton -Rows $Rows -Ids @(1) -PreferredText "OK")
}

function Invoke-OpenGd77CpsSelectPortDialog {
    param(
        [object]$Dialog,
        [object[]]$Rows
    )

    $portCount = 0
    $selectionSet = $false
    $selectedIndex = -1
    $portNames = New-Object System.Collections.Generic.List[string]

    $combo = $Rows | Where-Object { ([string]$_.Class -match "(?i)combo") } | Select-Object -First 1
    if ($combo) {
        try {
            $countPtr = [OpenGd77CpsWinApi]::SendMessage(
                $combo.Hwnd,
                [uint32][OpenGd77CpsWinApi]::CB_GETCOUNT,
                [IntPtr]::Zero,
                [IntPtr]::Zero
            )
            $portCount = [int]$countPtr.ToInt64()
        } catch {
            $portCount = 0
        }

        if ($portCount -gt 0) {
            for ($idx = 0; $idx -lt $portCount; $idx++) {
                try {
                    $lenPtr = [OpenGd77CpsWinApi]::SendMessage(
                        $combo.Hwnd,
                        [uint32][OpenGd77CpsWinApi]::CB_GETLBTEXTLEN,
                        [IntPtr]$idx,
                        [IntPtr]::Zero
                    )
                    $len = [int]$lenPtr.ToInt64()
                    if ($len -lt 0) { continue }

                    $sb = New-Object System.Text.StringBuilder ([Math]::Max(1, $len + 1))
                    [void][OpenGd77CpsWinApi]::SendMessage(
                        $combo.Hwnd,
                        [uint32][OpenGd77CpsWinApi]::CB_GETLBTEXT,
                        [IntPtr]$idx,
                        $sb
                    )
                    $entry = $sb.ToString().Trim()
                    if (-not [string]::IsNullOrWhiteSpace($entry)) {
                        $portNames.Add($entry) | Out-Null
                    }
                } catch {
                    # continue best-effort extraction
                }
            }

            $selectedIndex = 0
            if ($portNames.Count -gt 0) {
                $preferred = -1

                for ($i = 0; $i -lt $portNames.Count; $i++) {
                    if ($portNames[$i] -match "(?i)OpenGD77|MD-UV|UV390|UV380|GD-?77") {
                        $preferred = $i
                        break
                    }
                }
                if ($preferred -lt 0) {
                    for ($i = 0; $i -lt $portNames.Count; $i++) {
                        if (($portNames[$i] -match "(?i)USB|Serial|CP210|CH340|FTDI|Silicon|STM|Virtual") -and ($portNames[$i] -notmatch "(?i)\(COM1\)$")) {
                            $preferred = $i
                            break
                        }
                    }
                }
                if ($preferred -lt 0 -and $portNames.Count -gt 1) {
                    for ($i = 0; $i -lt $portNames.Count; $i++) {
                        if ($portNames[$i] -notmatch "(?i)\(COM1\)$") {
                            $preferred = $i
                            break
                        }
                    }
                }
                if ($preferred -ge 0) {
                    $selectedIndex = $preferred
                }
            }

            [void][OpenGd77CpsWinApi]::SendMessage(
                $combo.Hwnd,
                [uint32][OpenGd77CpsWinApi]::CB_SETCURSEL,
                [IntPtr]$selectedIndex,
                [IntPtr]::Zero
            )
            $selectionSet = $true
        }
    }

    [void][OpenGd77CpsWinApi]::SetForegroundWindow($dialog.Hwnd)
    Start-Sleep -Milliseconds 120

    if (-not $selectionSet) {
        # Keyboard fallback for owner-drawn combos where child handles are limited.
        [System.Windows.Forms.SendKeys]::SendWait("{HOME}{DOWN}")
        Start-Sleep -Milliseconds 90
    }

    $clicked = $false
    $selectButton = $Rows | Where-Object {
        ([string]$_.Class -match "(?i)button") -and ([string]$_.Text -match "(?i)select\s*port|ok")
    } | Select-Object -First 1
    $cancelButton = $Rows | Where-Object {
        ([string]$_.Class -match "(?i)button") -and ([string]$_.Text -match "(?i)^cancel$|anuluj")
    } | Select-Object -First 1

    $noPorts = ($portCount -le 0)
    if ($noPorts) {
        if ($cancelButton) {
            [void][OpenGd77CpsWinApi]::SendMessage(
                $cancelButton.Hwnd,
                [uint32][OpenGd77CpsWinApi]::BM_CLICK,
                [IntPtr]::Zero,
                [IntPtr]::Zero
            )
        } else {
            [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
        }
        $clicked = $true
    } elseif ($selectButton) {
        [void][OpenGd77CpsWinApi]::SendMessage(
            $selectButton.Hwnd,
            [uint32][OpenGd77CpsWinApi]::BM_CLICK,
            [IntPtr]::Zero,
            [IntPtr]::Zero
        )
        $clicked = $true
    } else {
        if (-not $noPorts) {
            if (-not (Invoke-OpenGd77CpsDialogButton -Rows $Rows -Ids @(1, 2, 6) -PreferredText "Select port")) {
                if (-not (Invoke-OpenGd77CpsDialogButton -Rows $Rows -Ids @(1, 2, 6) -PreferredText "OK")) {
                    [System.Windows.Forms.SendKeys]::SendWait("{TAB}{ENTER}")
                }
            }
        }
        $clicked = $true
    }

    return [pscustomobject]@{
        PortCount = $portCount
        PortNames = $portNames.ToArray()
        SelectedIndex = $selectedIndex
        SelectionSet = $selectionSet
        NoPorts = $noPorts
        Clicked = $clicked
    }
}

function Handle-OpenGd77CpsDialogs {
    param(
        [uint32]$ProcessId,
        [int]$Seconds,
        [string]$BrowseFolder,
        [string]$ActionName = "generic"
    )

    $captured = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $deadline = (Get-Date).AddSeconds($Seconds)
    $sawSupportWindow = $false
    $supportVisiblePrevLoop = $false
    $supportCycleCompleted = $false
    $sawReadComplete = $false
    $sawWriteComplete = $false
    $selectorHandledCount = 0
    $selectorSnapshots = New-Object System.Collections.Generic.List[string]

    $errorRegex = "(?i)line\s+\d+\s+is\s+not\s+valid|error|błąd|blad|failed|no\s+com\s+port|comm\s+port\s+not\s+available|radio\s+not\s+detected|failed\s+to\s+open\s+comm\s+port|please\s+connect.*radio"
    $readCompleteRegex = "(?i)read[\s_]+codeplug[\s_]+complete|read\s+complete|wczytano"
    $writeCompleteRegex = "(?i)write[\s_]+codeplug[\s_]+complete|write\s+complete|upload\s+complete|zapisano"

    while ((Get-Date) -lt $deadline) {
        $dialogs = Get-OpenGd77CpsDialogs -ProcessId $ProcessId
        $supportVisibleThisLoop = $false

        if ($dialogs.Count -eq 0) {
            if ($sawSupportWindow -and $supportVisiblePrevLoop) {
                $supportCycleCompleted = $true
            }
            $supportVisiblePrevLoop = $false
            Start-Sleep -Milliseconds 260
            [System.Windows.Forms.Application]::DoEvents()
            continue
        }

        foreach ($dialog in $dialogs) {
            $dialogTitle = [string]$dialog.Title
            $dialogClass = [string]$dialog.Class
            if ($dialogTitle -match "(?i)^OpenGD77 CPS") { continue }
            if ($dialogClass -in @(
                    ".NET-BroadcastEventWindow.4.0.0.0.34f5582.0",
                    "GDI+ Hook Window Class",
                    "MSCTFIME UI",
                    "IME",
                    "Internet Explorer_Hidden"
                )) {
                continue
            }

            $rows = Get-OpenGd77CpsDialogRows -DialogHandle $dialog.Hwnd
            $message = (($rows | Where-Object { $_.Class -eq "Static" -and -not [string]::IsNullOrWhiteSpace([string]$_.Text) } | ForEach-Object { $_.Text }) -join "`n").Trim()
            $messageWithTitle = (($dialogTitle + " " + $message).Trim())

            $key = ([string]$dialog.Hwnd) + "|" + $dialogClass + "|" + $dialogTitle + "|" + $message
            if (-not $seen.Contains($key)) {
                $seen.Add($key) | Out-Null
                if (-not [string]::IsNullOrWhiteSpace($message)) {
                    Append-Log (TF "log_cps_dialog_fmt" ($message -replace "`r?`n", " | "))
                } elseif (-not [string]::IsNullOrWhiteSpace($dialogTitle)) {
                    Append-Log (TF "log_cps_dialog_fmt" ("[" + $dialogClass + "] " + $dialogTitle))
                }
            }

            if ($dialogTitle -match "(?i)^OpenGD77 Support$") {
                $sawSupportWindow = $true
                $supportVisibleThisLoop = $true
                if ($messageWithTitle -match $errorRegex) {
                    $errors.Add($messageWithTitle) | Out-Null
                }
                # Never auto-click arbitrary buttons in the OpenGD77 support form.
                continue
            }

            if ($dialogTitle -match "(?i)select\s+opengd77\s+com\s+port|select\s+.*com\s+port") {
                $selectorHandledCount++
                $selectorResult = Invoke-OpenGd77CpsSelectPortDialog -Dialog $dialog -Rows $rows
                $portsPreview = ""
                if ($selectorResult.PortNames) {
                    $portsPreview = (@($selectorResult.PortNames | Select-Object -First 6) -join " || ")
                }
                if (-not [string]::IsNullOrWhiteSpace($portsPreview)) {
                    $selectorSnapshots.Add($portsPreview) | Out-Null
                }
                Append-Log (TF "log_cps_action_fmt" ("COM selector handled (ports={0}, selected_index={1}, selection_set={2}, no_ports={3}, clicked={4}, names={5})" -f $selectorResult.PortCount, $selectorResult.SelectedIndex, $selectorResult.SelectionSet, $selectorResult.NoPorts, $selectorResult.Clicked, $portsPreview))
                if ($selectorResult.NoPorts) {
                    $errors.Add("No com port detected in OpenGD77 selector.") | Out-Null
                    if ($ActionName -in @("write", "read")) {
                        $deadline = (Get-Date).AddMilliseconds(-1)
                    }
                }
                Start-Sleep -Milliseconds 250
                continue
            }

            $isBrowseDialog = $rows | Where-Object { $_.Class -eq "Edit" -and [int]$_.Id -eq 14148 } | Select-Object -First 1
            if ($isBrowseDialog) {
                [void](Invoke-OpenGd77CpsBrowseFolderDialog -Rows $rows -FolderPath $BrowseFolder)
                Start-Sleep -Milliseconds 250
                continue
            }

            if ($messageWithTitle -match "(?i)are you sure|replace|all current data will be replaced|zastąpi") {
                if (-not (Invoke-OpenGd77CpsDialogButton -Rows $rows -Ids @(6) -PreferredText "Yes")) {
                    [void](Invoke-OpenGd77CpsDialogButton -Rows $rows -Ids @(1) -PreferredText "OK")
                }
                Start-Sleep -Milliseconds 220
                continue
            }

            if ($messageWithTitle -match "(?i)appended|imported|codeplug now contains|codeplug contains|csvs created|read[\s_]+codeplug[\s_]+complete|write[\s_]+codeplug[\s_]+complete|write complete|read complete|upload complete|zapisano|wczytano") {
                $captured.Add($message) | Out-Null
                if ($messageWithTitle -match $readCompleteRegex) {
                    $sawReadComplete = $true
                }
                if ($messageWithTitle -match $writeCompleteRegex) {
                    $sawWriteComplete = $true
                }
                if ($dialogClass -eq "#32770") {
                    if (-not (Invoke-OpenGd77CpsDialogButton -Rows $rows -Ids @(2) -PreferredText "OK")) {
                        [void](Invoke-OpenGd77CpsDialogButton -Rows $rows -Ids @(1) -PreferredText "OK")
                    }
                }
                Start-Sleep -Milliseconds 220
                continue
            }

            if ($messageWithTitle -match $errorRegex) {
                $captured.Add($message) | Out-Null
                $errors.Add($messageWithTitle) | Out-Null
                if ($dialogClass -eq "#32770") {
                    [void](Invoke-OpenGd77CpsDialogButton -Rows $rows -Ids @(1, 2) -PreferredText "OK")
                }
                Start-Sleep -Milliseconds 220
                continue
            }

            # Fallback click only for standard dialogs, never for full application windows.
            if ($dialogClass -eq "#32770") {
                [void](Invoke-OpenGd77CpsDialogButton -Rows $rows -Ids @(1, 2, 6, 7) -PreferredText "OK")
            }
            Start-Sleep -Milliseconds 220
        }

        if ($sawSupportWindow -and $supportVisiblePrevLoop -and (-not $supportVisibleThisLoop)) {
            $supportCycleCompleted = $true
        }
        $supportVisiblePrevLoop = $supportVisibleThisLoop

        Start-Sleep -Milliseconds 220
        [System.Windows.Forms.Application]::DoEvents()
    }

    Append-Log (TF "log_cps_action_fmt" ("Dialog handler '{0}': messages={1}, errors={2}, support_seen={3}, support_cycle={4}, read_ok={5}, write_ok={6}, selectors={7}" -f $ActionName, $captured.Count, $errors.Count, $sawSupportWindow, $supportCycleCompleted, $sawReadComplete, $sawWriteComplete, $selectorHandledCount))
    return [pscustomobject]@{
        Messages = $captured.ToArray()
        Errors = $errors.ToArray()
        SawSupportWindow = $sawSupportWindow
        SupportCycleCompleted = $supportCycleCompleted
        SawReadComplete = $sawReadComplete
        SawWriteComplete = $sawWriteComplete
        SelectorHandledCount = $selectorHandledCount
        SelectorSnapshots = $selectorSnapshots.ToArray()
    }
}

function Import-CsvWithSmartDelimiter {
    param([string]$Path)

    $firstLine = ""
    try {
        $firstLine = (Get-Content -Path $Path -TotalCount 1 -Encoding UTF8)
    } catch {
        $firstLine = ""
    }
    $delimiter = ";"
    if (($firstLine -notmatch ";") -and ($firstLine -match ",")) {
        $delimiter = ","
    }
    return Import-Csv -Path $Path -Delimiter $delimiter
}

function Sync-OpenGd77BundleToDesktopImportFiles {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Bundle
    )

    $desktop = [Environment]::GetFolderPath("Desktop")
    $desktopChannels = Join-Path $desktop "Channels.csv"
    $desktopZones = Join-Path $desktop "Zones.csv"

    if (-not (Test-Path $Bundle.ChannelsPath)) {
        throw "Brak Channels.csv w pakiecie do importu."
    }
    if (-not (Test-Path $Bundle.ZonesPath)) {
        throw "Brak Zones.csv w pakiecie do importu."
    }

    Copy-Item -Path $Bundle.ChannelsPath -Destination $desktopChannels -Force
    Copy-Item -Path $Bundle.ZonesPath -Destination $desktopZones -Force

    return [pscustomobject]@{
        DesktopFolder = $desktop
        DesktopChannels = $desktopChannels
        DesktopZones = $desktopZones
    }
}

function Get-OpenGd77ChannelCountFromMessages {
    param([string[]]$Messages)

    foreach ($m in @($Messages)) {
        if ([string]::IsNullOrWhiteSpace($m)) { continue }
        $match = [regex]::Match([string]$m, "(?i)\b(\d+)\s+Channels\b")
        if ($match.Success) {
            return [int]$match.Groups[1].Value
        }
    }
    return $null
}

function Merge-OpenGd77DialogResults {
    param(
        [Parameter(Mandatory = $true)]
        [object]$First,
        [Parameter(Mandatory = $true)]
        [object]$Second
    )

    return [pscustomobject]@{
        Messages = @($First.Messages + $Second.Messages)
        Errors = @($First.Errors + $Second.Errors)
        SawSupportWindow = ([bool]$First.SawSupportWindow -or [bool]$Second.SawSupportWindow)
        SupportCycleCompleted = ([bool]$First.SupportCycleCompleted -or [bool]$Second.SupportCycleCompleted)
        SawReadComplete = ([bool]$First.SawReadComplete -or [bool]$Second.SawReadComplete)
        SawWriteComplete = ([bool]$First.SawWriteComplete -or [bool]$Second.SawWriteComplete)
        SelectorHandledCount = ([int]$First.SelectorHandledCount + [int]$Second.SelectorHandledCount)
        SelectorSnapshots = @($First.SelectorSnapshots + $Second.SelectorSnapshots)
    }
}

function Invoke-OpenGd77CpsImportWriteVerify {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Bundle
    )

    if ($null -eq $Bundle) { throw "Bundle is null." }
    if ([string]::IsNullOrWhiteSpace([string]$Bundle.BundleDir)) { throw "BundleDir is missing." }
    if ([string]::IsNullOrWhiteSpace([string]$Bundle.VerifyDir)) { throw "VerifyDir is missing." }

    Ensure-OpenGd77CpsWinApi
    $session = Get-OpenGd77CpsSession
    $processId = [uint32]$session.ProcessId
    $mainWindowHandle = [IntPtr]$session.MainWindowHandle
    $expectedCount = [int]$Bundle.Count

    if (-not (Test-Path $Bundle.VerifyDir)) { [void](New-Item -ItemType Directory -Path $Bundle.VerifyDir) }
    $desktopSync = Sync-OpenGd77BundleToDesktopImportFiles -Bundle $Bundle
    Append-Log (TF "log_cps_action_fmt" ("Import files synced to Desktop: " + $desktopSync.DesktopFolder))
    $importBrowseFolder = [string]$desktopSync.DesktopFolder
    $exportBrowseFolder = [string]$desktopSync.DesktopFolder

    Append-Log (TF "log_cps_action_fmt" "Import CSV (replace)")
    Set-OpenGd77CpsForeground -MainWindowHandle $mainWindowHandle
    [System.Windows.Forms.SendKeys]::SendWait("%f{DOWN}{DOWN}{DOWN}{RIGHT}{DOWN}{ENTER}")
    Start-Sleep -Milliseconds 500
    $importResult = Handle-OpenGd77CpsDialogs -ProcessId $processId -Seconds 45 -BrowseFolder $importBrowseFolder -ActionName "import"
    $importSucceeded = @($importResult.Messages | Where-Object { $_ -match "(?i)appended|imported|codeplug now contains" }).Count -gt 0
    if (-not $importSucceeded) {
        Append-Log (TF "log_cps_action_fmt" "Import trigger fallback: retry menu sequence")
        Set-OpenGd77CpsForeground -MainWindowHandle $mainWindowHandle
        [System.Windows.Forms.SendKeys]::SendWait("%f{DOWN}{DOWN}{DOWN}{RIGHT}{DOWN}{ENTER}")
        Start-Sleep -Milliseconds 500
        $importRetryResult = Handle-OpenGd77CpsDialogs -ProcessId $processId -Seconds 55 -BrowseFolder $importBrowseFolder -ActionName "import-retry"
        $importResult = Merge-OpenGd77DialogResults -First $importResult -Second $importRetryResult
        $importSucceeded = @($importResult.Messages | Where-Object { $_ -match "(?i)appended|imported|codeplug now contains" }).Count -gt 0
    }
    if (-not $importSucceeded) {
        throw "Brak potwierdzenia importu CSV do OpenGD77 CPS."
    }
    $importedCount = Get-OpenGd77ChannelCountFromMessages -Messages @($importResult.Messages)
    if ($null -ne $importedCount -and $importedCount -ne $expectedCount) {
        throw ("Import CSV do CPS zakończył się niepoprawną liczbą kanałów (zaimportowano {0}, oczekiwano {1})." -f $importedCount, $expectedCount)
    }

    Append-Log (TF "log_cps_action_fmt" "Write to radio")
    Set-OpenGd77CpsForeground -MainWindowHandle $mainWindowHandle
    $writeTriggered = Invoke-OpenGd77CpsMainToolbarClick -MainWindowHandle $mainWindowHandle -OffsetX 118 -OffsetY 58
    if (-not $writeTriggered) {
        [System.Windows.Forms.SendKeys]::SendWait("^w")
    }
    Start-Sleep -Milliseconds 300
    $writeResult = Handle-OpenGd77CpsDialogs -ProcessId $processId -Seconds 120 -BrowseFolder $importBrowseFolder -ActionName "write"
    $writeCompletedSignal = ([bool]$writeResult.SawWriteComplete) -or (([bool]$writeResult.SawSupportWindow) -and ([bool]$writeResult.SupportCycleCompleted))
    if (-not $writeCompletedSignal) {
        Append-Log (TF "log_cps_action_fmt" "Write trigger fallback: Ctrl+W")
        Set-OpenGd77CpsForeground -MainWindowHandle $mainWindowHandle
        [System.Windows.Forms.SendKeys]::SendWait("^w")
        Start-Sleep -Milliseconds 280
        $writeRetryResult = Handle-OpenGd77CpsDialogs -ProcessId $processId -Seconds 90 -BrowseFolder $importBrowseFolder -ActionName "write-retry"
        $writeResult = Merge-OpenGd77DialogResults -First $writeResult -Second $writeRetryResult
    }

    Append-Log (TF "log_cps_action_fmt" "Read from radio")
    Set-OpenGd77CpsForeground -MainWindowHandle $mainWindowHandle
    $readTriggered = Invoke-OpenGd77CpsMainToolbarClick -MainWindowHandle $mainWindowHandle -OffsetX 95 -OffsetY 58
    if (-not $readTriggered) {
        [System.Windows.Forms.SendKeys]::SendWait("^r")
    }
    Start-Sleep -Milliseconds 300
    $readResult = Handle-OpenGd77CpsDialogs -ProcessId $processId -Seconds 90 -BrowseFolder $importBrowseFolder -ActionName "read"
    if (-not ([bool]$readResult.SawReadComplete)) {
        Append-Log (TF "log_cps_action_fmt" "Read trigger fallback: Ctrl+R")
        Set-OpenGd77CpsForeground -MainWindowHandle $mainWindowHandle
        [System.Windows.Forms.SendKeys]::SendWait("^r")
        Start-Sleep -Milliseconds 280
        $readRetryResult = Handle-OpenGd77CpsDialogs -ProcessId $processId -Seconds 90 -BrowseFolder $importBrowseFolder -ActionName "read-retry"
        $readResult = Merge-OpenGd77DialogResults -First $readResult -Second $readRetryResult
    }

    Append-Log (TF "log_cps_action_fmt" "Export CSV for verify")
    Set-OpenGd77CpsForeground -MainWindowHandle $mainWindowHandle
    $exportStartedAt = Get-Date
    [System.Windows.Forms.SendKeys]::SendWait("%f{DOWN}{DOWN}{DOWN}{RIGHT}{ENTER}")
    $exportResult = Handle-OpenGd77CpsDialogs -ProcessId $processId -Seconds 40 -BrowseFolder $exportBrowseFolder -ActionName "export"
    if (@($exportResult.Messages).Count -eq 0) {
        Append-Log (TF "log_cps_action_fmt" "Export trigger fallback: retry menu sequence")
        Set-OpenGd77CpsForeground -MainWindowHandle $mainWindowHandle
        Start-Sleep -Milliseconds 320
        $exportStartedAt = Get-Date
        [System.Windows.Forms.SendKeys]::SendWait("%f{DOWN}{DOWN}{DOWN}{RIGHT}{ENTER}")
        $exportRetryResult = Handle-OpenGd77CpsDialogs -ProcessId $processId -Seconds 45 -BrowseFolder $exportBrowseFolder -ActionName "export-retry"
        $exportResult = Merge-OpenGd77DialogResults -First $exportResult -Second $exportRetryResult
    }

    $importMessages = @($importResult.Messages)
    $writeMessages = @($writeResult.Messages)
    $readMessages = @($readResult.Messages)
    $exportMessages = @($exportResult.Messages)

    $allDialogMessages = @($importMessages + $writeMessages + $readMessages + $exportMessages | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $allDialogErrors = @($importResult.Errors + $writeResult.Errors + $readResult.Errors + $exportResult.Errors | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $firstError = $allDialogErrors | Select-Object -First 1
    if (-not $firstError) {
        $firstError = $allDialogMessages | Where-Object { $_ -match "(?i)line\s+\d+\s+is\s+not\s+valid|error|błąd|blad|failed|no\s+com\s+port|radio\s+not\s+detected|comm\s+port\s+not\s+available" } | Select-Object -First 1
    }
    if ($firstError) {
        throw ([string]$firstError)
    }

    $selectorPortTokens = New-Object System.Collections.Generic.List[string]
    foreach ($snapshot in @($writeResult.SelectorSnapshots + $readResult.SelectorSnapshots)) {
        if ([string]::IsNullOrWhiteSpace([string]$snapshot)) { continue }
        foreach ($piece in ([string]$snapshot -split "\s+\|\|\s+")) {
            $clean = ([string]$piece).Trim()
            if ([string]::IsNullOrWhiteSpace($clean)) { continue }
            $alreadyThere = $selectorPortTokens | Where-Object { ([string]$_).Trim().ToUpperInvariant() -eq $clean.ToUpperInvariant() } | Select-Object -First 1
            if (-not $alreadyThere) {
                $selectorPortTokens.Add($clean) | Out-Null
            }
        }
    }
    $selectorPortsJoined = (@($selectorPortTokens | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) -join " | ")
    if (-not [string]::IsNullOrWhiteSpace($selectorPortsJoined)) {
        Append-Log (TF "log_cps_action_fmt" ("Selector ports observed: " + $selectorPortsJoined))
    }
    $selectorLooksWrong = $false
    if ($selectorPortTokens.Count -gt 0) {
        $joinedUpper = $selectorPortsJoined.ToUpperInvariant()
        $hasOnlyLegacyCom = ($joinedUpper -match "\(COM1\)") -and ($joinedUpper -notmatch "OPENGD77|CH340|CP210|FTDI|USB")
        if ($hasOnlyLegacyCom) {
            $selectorLooksWrong = $true
        }
    }

    $writeSucceeded = $writeResult.SawWriteComplete -or ($writeResult.SawSupportWindow -and $writeResult.SupportCycleCompleted)
    if (-not $writeSucceeded) {
        if ($selectorLooksWrong) {
            throw ("Brak potwierdzenia zapisu do radia. CPS widzi tylko port: {0}. Podłącz radio w trybie normalnym (nie DFU), sprawdź sterownik COM dla radia i spróbuj ponownie." -f $selectorPortsJoined)
        }
        throw "Brak potwierdzenia zakończenia zapisu codepluga do radia (Write)."
    }
    if (-not $readResult.SawReadComplete) {
        if ($selectorLooksWrong) {
            throw ("Brak potwierdzenia odczytu z radia. CPS widzi tylko port: {0}. Podłącz radio w trybie normalnym (nie DFU), sprawdź sterownik COM dla radia i spróbuj ponownie." -f $selectorPortsJoined)
        }
        throw "Brak potwierdzenia odczytu codepluga z radia (Read Codeplug complete)."
    }

    $exportChannelsPath = Join-Path $Bundle.VerifyDir "Channels.csv"
    if (-not (Test-Path $exportChannelsPath)) {
        $candidatePaths = @(
            (Join-Path $Bundle.BundleDir "Channels.csv"),
            (Join-Path ([Environment]::GetFolderPath("Desktop")) "Channels.csv")
        )
        foreach ($candidate in $candidatePaths) {
            if (-not (Test-Path $candidate)) { continue }
            $candidateInfo = Get-Item -Path $candidate
            if ($candidateInfo.LastWriteTime -ge $exportStartedAt.AddSeconds(-1)) {
                $exportChannelsPath = $candidate
                Append-Log (TF "log_cps_action_fmt" ("Export fallback path: " + $candidate))
                break
            }
        }
    }
    if (-not (Test-Path $exportChannelsPath)) {
        throw "Brak pliku weryfikacyjnego Channels.csv po eksporcie z CPS."
    }

    $exportRows = Import-CsvWithSmartDelimiter -Path $exportChannelsPath
    $exportCount = @($exportRows).Count

    $exportNames = @($exportRows | ForEach-Object { [string]$_."Channel Name" })
    $missingNames = @()
    if ($Bundle.ChannelNames) {
        $missingNames = @($Bundle.ChannelNames | Where-Object { $_ -and ($_ -notin $exportNames) })
    }

    if ($exportCount -ne $expectedCount) {
        throw (TF "msg_channels_verify_failed_fmt" $expectedCount $exportCount)
    }
    if ($missingNames.Count -gt 0) {
        throw (TF "msg_channels_verify_missing_names_fmt" $missingNames.Count)
    }

    return [pscustomobject]@{
        ExpectedChannels = $expectedCount
        ExportedChannels = $exportCount
        MissingNames = $missingNames.Count
        VerifyChannelsPath = $exportChannelsPath
        ImportMessages = $importMessages
        WriteMessages = $writeMessages
        ReadMessages = $readMessages
        ExportMessages = $exportMessages
    }
}

function Start-ChannelsImport {
    param(
        [switch]$FromUrl,
        [switch]$SkipSummaryDialog,
        [switch]$AutoSaveOutput
    )

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

        if ($AutoSaveOutput) {
            $savedPath = Save-NormalizedRepeaterRowsAuto -Rows $normalized
        } else {
            $savedPath = Save-NormalizedRepeaterRows -Rows $normalized
        }
        if ([string]::IsNullOrWhiteSpace($savedPath)) { return }

        Append-Log (TF "log_channels_saved_fmt" $savedPath)
        $cpsBundle = $null
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

        if (-not $SkipSummaryDialog) {
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
        }

        return [pscustomobject]@{
            Source = $sourceLabel
            Count = $normalized.Count
            SavedPath = $savedPath
            CpsBundle = $cpsBundle
            CpsBundlePath = $cpsBundlePath
            NormalizedRows = $normalized
        }
    } catch {
        Write-DebugLog -Level "ERROR" -Message ("CHANNELS_IMPORT failed: " + $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            (TF "msg_channels_error_fmt" $_.Exception.Message),
            (T "cap_error"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $null
    }
}

function Start-ChannelsImportAndWriteRadio {
    param([switch]$FromUrl)

    if ($script:ProcessRunning) { return }

    try {
        Append-Log (T "log_channels_write_begin")
        $importResult = Start-ChannelsImport -FromUrl:$FromUrl -SkipSummaryDialog -AutoSaveOutput
        if (-not $importResult) {
            Append-Log (T "log_channels_write_canceled")
            return
        }

        $bundle = Save-OpenGd77CpsBundleCompatible -Rows $importResult.NormalizedRows
        if (-not $bundle) {
            throw "Nie udalo sie przygotowac kompatybilnego pakietu CPS."
        }
        Append-Log (TF "log_channels_cps_saved_fmt" $bundle.BundleDir)

        Set-Busy $true (T "status_channels_write")
        try {
            $verify = Invoke-OpenGd77CpsImportWriteVerify -Bundle $bundle
            Append-Log (TF "log_channels_verify_ok_fmt" $verify.ExportedChannels $verify.MissingNames)
            Set-Busy $false (T "status_success")
            Play-RadioSuccessChirp

            [System.Windows.Forms.MessageBox]::Show(
                (TF "msg_channels_write_done_fmt" $bundle.Count $bundle.ZoneCount $verify.ExportedChannels $bundle.BundleDir),
                (T "cap_channels_write"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } catch {
            Write-DebugLog -Level "ERROR" -Message ("CHANNELS_WRITE failed: " + $_.Exception.Message)
            Set-Busy $false (T "status_channels_write_error")
            [System.Windows.Forms.MessageBox]::Show(
                (TF "msg_channels_write_error_fmt" $_.Exception.Message),
                (T "cap_error"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    } catch {
        Write-DebugLog -Level "ERROR" -Message ("CHANNELS_WRITE precheck failed: " + $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            (TF "msg_channels_write_error_fmt" $_.Exception.Message),
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
$btnCheck.Size = New-Object System.Drawing.Size(130, 32)
$btnCheck.TabIndex = 4
$form.Controls.Add($btnCheck)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = (T "btn_start")
$btnStart.Location = New-Object System.Drawing.Point(146, 162)
$btnStart.Size = New-Object System.Drawing.Size(130, 32)
$btnStart.TabIndex = 5
$form.Controls.Add($btnStart)

$btnProgramUpdate = New-Object System.Windows.Forms.Button
$btnProgramUpdate.Text = (T "btn_program_update")
$btnProgramUpdate.Location = New-Object System.Drawing.Point(280, 162)
$btnProgramUpdate.Size = New-Object System.Drawing.Size(130, 32)
$btnProgramUpdate.TabIndex = 6
$form.Controls.Add($btnProgramUpdate)

$btnChannelsFile = New-Object System.Windows.Forms.Button
$btnChannelsFile.Text = (T "btn_channels_file")
$btnChannelsFile.Location = New-Object System.Drawing.Point(414, 162)
$btnChannelsFile.Size = New-Object System.Drawing.Size(130, 32)
$btnChannelsFile.TabIndex = 7
$form.Controls.Add($btnChannelsFile)

$btnChannelsUrl = New-Object System.Windows.Forms.Button
$btnChannelsUrl.Text = (T "btn_channels_url")
$btnChannelsUrl.Location = New-Object System.Drawing.Point(548, 162)
$btnChannelsUrl.Size = New-Object System.Drawing.Size(130, 32)
$btnChannelsUrl.TabIndex = 8
$form.Controls.Add($btnChannelsUrl)

$btnChannelsWrite = New-Object System.Windows.Forms.Button
$btnChannelsWrite.Text = (T "btn_channels_write")
$btnChannelsWrite.Location = New-Object System.Drawing.Point(682, 162)
$btnChannelsWrite.Size = New-Object System.Drawing.Size(130, 32)
$btnChannelsWrite.TabIndex = 9
$form.Controls.Add($btnChannelsWrite)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = (T "btn_close")
$btnClose.Location = New-Object System.Drawing.Point(816, 162)
$btnClose.Size = New-Object System.Drawing.Size(96, 32)
$btnClose.TabIndex = 10
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
$txtStatus.TabIndex = 11
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
$txtLastMessage.TabIndex = 12
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
$txtLog.TabIndex = 13
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
$script:btnChannelsWrite = $btnChannelsWrite
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
    $script:btnChannelsWrite.Text = (T "btn_channels_write")
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

$btnChannelsWrite.Add_Click({
    if ($script:ProcessRunning) { return }
    Write-DebugLog "UI_CLICK: Import i zapis kanałów do radia"
    Start-ChannelsImportAndWriteRadio
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
    } elseif ($e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::W) {
        $btnChannelsWrite.PerformClick()
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
