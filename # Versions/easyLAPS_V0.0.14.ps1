# [1.1 | Script Version Requirement]
# ENGLISH - Ensures PowerShell version 5.1 or higher is used
# GERMAN - Stellt sicher, dass mindestens PowerShell Version 5.1 verwendet wird
#requires -Version 5.1
[CmdletBinding()]
param()

<#
  easyLAPS – PS5, für das **neue** integrierte LAPS
   - Nutzt Cmdlets: Get-LapsADPassword, Reset-LapsPassword, etc. (Parameter -Identity)
   - Zeigt Computer in einer reinen AD-Computerliste
   - GUI:
     * Header: Links AppName, Mitte "ComputerName" + TextBox + "LAPS auslesen", Rechts Logo
     * Links: GroupBox "Local Administrator Password Solution (Windows LAPS 2.0)"
     * Rechts: AD-Computerliste, Refresh-Button
     * Footer: Platzhaltertext (Version, Datum, Autor) + optionaler Footer
#>

# Globale Variablen für das Skript
$Global:DebugEnabled = $false
$ScriptDir = Split-Path -Parent $PSCommandPath

#region [2.1 | Read INI File - INI-Datei lesen]
# ENGLISH - Reads and parses the specified INI file, ignoring blank or commented lines
# GERMAN - Liest und analysiert die angegebene INI-Datei und ignoriert leere oder auskommentierte Zeilen
function Read-INIFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        Throw "INI-Datei nicht gefunden: $Path"
    }
    try {
        $iniContent = Get-Content -Path $Path | Where-Object {
            $_.Trim() -ne "" -and $_ -notmatch '^\s*[;#]'
        }
    }
    catch {
        Throw "Fehler beim Lesen der INI-Datei: $($_.Exception.Message)"
    }

    $result  = New-Object 'System.Collections.Specialized.OrderedDictionary'
    $section = $null
    foreach ($line in $iniContent) {
        if ($line -match '^\[(.+)\]$') {
            $section = $matches[1].Trim()
            if (-not $result.Contains($section)) {
                $result[$section] = New-Object System.Collections.Specialized.OrderedDictionary
            }
        }
        elseif ($line -match '^(.*?)=(.*)$') {
            $key   = $matches[1].Trim()
            $value = $matches[2].Trim()
            if ($section -and $result[$section]) {
                $result[$section][$key] = $value
            }
        }
    }
    return $result
}
#endregion

#region [3.1 | Logging Functions - Protokollierungsfunktionen]
# [3.1.1 | Primary Logging Function - Primäre Protokollierungsfunktion]
# ENGLISH - Primary logging wrapper for consistent message formatting
# GERMAN - Primäre Protokollierungsfunktion für konsistente Nachrichtenformatierung
function Write-LogMessage {
    param(
        [string]$message,
        [string]$logLevel = "INFO"
    )
    Write-Log -message $message -logLevel $logLevel
}

# [3.1.2 | Debug Logging Function - Debug-Protokollierungsfunktion]
# ENGLISH - Logs debug messages when DebugMode is enabled
# GERMAN - Protokolliert Debug-Nachrichten, wenn DebugMode aktiviert ist
function Write-DebugMessage {
    param(
        [string]$Message
    )
    # Angepasst um Global:DebugEnabled zu verwenden
    if ($Global:DebugEnabled) {
        # Use [void] to ensure no output is sent to the pipeline
        [void](Write-Log -Message "$Message" -LogLevel "DEBUG")
    }
}
Write-Host "Logging-Funktion..."
# Log-Level (Standard ist "INFO"): "WARN", "ERROR", "DEBUG".
#endregion

#region [3.2 | Log File Writer - Protokoll-Datei-Schreiber]
# [3.2.1 | Low-level File Logging Implementation - Implementierung der Dateiprotokollierung]
# ENGLISH - Writes log messages to file system with error handling
# GERMAN - Schreibt Lognachrichten ins Dateisystem mit Fehlerbehandlung
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Message,
        [string]$LogLevel = "INFO"
    )
    process {
        # Fallback für Log-Verzeichnis
        $logFilePath = Join-Path $ScriptDir "Logs"
        
        # Log-Verzeichnis sicherstellen
        if (-not (Test-Path $logFilePath)) {
            try {
                New-Item -ItemType Directory -Path $logFilePath -Force -ErrorAction Stop | Out-Null
            } catch {
                Write-Warning "Fehler beim Erstellen des Log-Verzeichnisses: $($_.Exception.Message)"
                return
            }
        }

        $logFile = Join-Path -Path $logFilePath -ChildPath "easyLAPS.log"
        $errorLogFile = Join-Path -Path $logFilePath -ChildPath "easyLAPS_error.log"
        $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timeStamp [$LogLevel] $Message"

        try {
            Add-Content -Path $logFile -Value $logEntry -ErrorAction Stop
        } catch {
            # Bei Schreibfehler in der Logdatei, zusätzlichen Fehler loggen
            Add-Content -Path $errorLogFile -Value "$timeStamp [ERROR] Fehler beim Schreiben in Logdatei: $($_.Exception.Message)" -ErrorAction SilentlyContinue
            Write-Warning "Fehler beim Schreiben in Logdatei: $($_.Exception.Message)"
        }

        if ($Global:DebugEnabled) {
            Write-Output "Debug: $logEntry"
        }
    }
}
#endregion

#region [2.2 | Host Log Function - Protokoll an Host ausgeben]
# ENGLISH - Outputs log messages to the host
# GERMAN - Gibt Logmeldungen an den Host aus
function Write-Log {
    param([string]$Message)
    Write-Host "[LOG] $Message"
}
#endregion

#region [2.3 | LAPS Functions - LAPS-Funktionen]
# ENGLISH - Contains functions that interact with Windows LAPS 2.0
# GERMAN - Enthält Funktionen, die mit Windows LAPS 2.0 interagieren

# [2.3.1 | function Load-Computers]
function Load-Computers {
    [CmdletBinding()]
    param()
    
    if ($Global:DebugEnabled) {
        Write-DebugMessage "Attempting to load AD computers..."
    }
    try {
        if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
            Throw "Active Directory module is not installed or not available."
        }
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        Write-Warning "Fehler beim Laden des Active Directory Moduls: $($_.Exception.Message)"
        return @()
    }

    try {
        $comps = Get-ADComputer -Filter * -Property Name
        if (-not $comps) {
            Write-Warning "Keine Computerobjekte gefunden."
            return @()
        }
        if ($Global:DebugEnabled) {
            Write-DebugMessage "Loaded $($comps.Count) AD computer objects."
        }
        return $comps
    }
    catch {
        Write-Warning "Fehler beim Laden der Computerobjekte: $($_.Exception.Message)"
        return @()
    }
}

# [2.3.2 | function Get-LapsPW]
function Get-LapsPW {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    
    try {
        # LAPS Modul prüfen
        if (-not (Get-Command -Name Get-LapsADPassword -ErrorAction SilentlyContinue)) {
            Throw "LAPS PowerShell Modul ist nicht installiert."
        }
        
        # LAPS-Passwort abrufen
        $lapsInfo = Get-LapsADPassword -Identity $ComputerName -ErrorAction Stop
        
        # SecureString in Klartext umwandeln
        $plainTextPassword = $null
        if ($lapsInfo.Password -is [System.Security.SecureString]) {
            $plainTextPassword = Convert-SecureStringToPlainText -SecureString $lapsInfo.Password
        }
        else {
            $plainTextPassword = $lapsInfo.Password
        }
        
        # Erstelle ein Objekt mit den relevanten Informationen
        $result = [PSCustomObject]@{
            Password = $plainTextPassword
            ExpirationTimestamp = $lapsInfo.ExpirationTimestamp
        }
        
        return $result
    }
    catch {
        Write-Warning "Fehler beim Abrufen des LAPS-Passworts für $ComputerName`: $($_.Exception.Message)"
        throw $_
    }
}

# [2.3.3 | function Reset-LapsPWNow]
function Reset-LapsPWNow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    
    try {
        # LAPS Modul prüfen
        if (-not (Get-Command -Name Reset-LapsPassword -ErrorAction SilentlyContinue)) {
            Throw "LAPS PowerShell Modul ist nicht installiert."
        }
        
        # LAPS-Passwort zurücksetzen
        Reset-LapsPassword -Identity $ComputerName -ErrorAction Stop
        
        Write-DebugMessage "LAPS-Passwort für $ComputerName wurde zurückgesetzt."
        return $true
    }
    catch {
        Write-Warning "Fehler beim Zurücksetzen des LAPS-Passworts für $ComputerName`: $($_.Exception.Message)"
        throw $_
    }
}
#endregion

#region [2.4 | Helper Functions: Clipboard and URL - Hilfsfunktionen: Zwischenablage und URL]
# [2.4.1 | Set Clipboard - Zwischenablage setzen]
# ENGLISH - Copies the specified text to the clipboard
# GERMAN - Kopiert den angegebenen Text in die Zwischenablage
function Set-Clipboard {
    param([string]$Text)
    try {
        Add-Type -AssemblyName PresentationCore -ErrorAction Stop
        [System.Windows.Clipboard]::SetText($Text)
    }
    catch {
        Write-Warning "Fehler beim Kopieren in die Zwischenablage: $($_.Exception.Message)"
    }
}

# [2.4.2 | Open URL in Browser - URL im Browser öffnen]
# ENGLISH - Opens the specified URL in the default browser
# GERMAN - Öffnet die angegebene URL im Standardbrowser
function Open-URLInBrowser {
    param([string]$URL)
    try {
        if (-not [string]::IsNullOrWhiteSpace($URL)) {
            # Überprüfe, ob URL ein gültiges Format hat
            if ($URL -match '^https?://') {
                [System.Diagnostics.Process]::Start($URL) | Out-Null
            }
            else {
                Write-Warning "Die URL '$URL' scheint ungültig zu sein (muss mit http:// oder https:// beginnen)."
            }
        }
    }
    catch {
        Write-Warning "URL konnte nicht geöffnet werden: $URL - $($_.Exception.Message)"
    }
}

# [2.4.3 | Convert SecureString to Plain Text - SecureString in Klartext umwandeln]
# ENGLISH - Converts a SecureString to plain text
# GERMAN - Konvertiert einen SecureString in Klartext
function Convert-SecureStringToPlainText {
    param(
        [Parameter(Mandatory=$true)]
        [System.Security.SecureString]$SecureString
    )
    
    try {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        $plainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        return $plainText
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}
#endregion

#region [2.5 | Main GUI Function - Haupt-GUI-Funktion]
# ENGLISH - Builds and shows the main LAPS GUI with all interface elements in English
# GERMAN - Erstellt und zeigt die Haupt-GUI für LAPS, alle GUI-Elemente in Englisch
function Show-LAPSForm {
    param(
        [Parameter(Mandatory=$true)][hashtable]$INI
    )

    if (-not $INI) {
        Throw "INI-Parameter darf nicht leer sein."
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # (1) Werte aus INI
    $appName        = $INI.'Branding-GUI'.APPName
    if (-not $appName) { $appName = "easyLAPS" }
    $themeColorStr  = $INI.'Branding-GUI'.ThemeColor
    if (-not $themeColorStr) { $themeColorStr = "DarkGray" }
    $boxColorStr    = $INI.'Branding-GUI'.BoxColor
    if (-not $boxColorStr) { $boxColorStr = "LightGray" }
    $rahmenColorStr = $INI.'Branding-GUI'.RahmenColor
    if (-not $rahmenColorStr) { $rahmenColorStr = "DarkBlue" }
    $fontFamily     = $INI.'Branding-GUI'.FontFamily
    if (-not $fontFamily) { $fontFamily = "Arial" }
    $fontSize       = $INI.'Branding-GUI'.FontSize
    if (-not $fontSize) { $fontSize = 9 }
    $headerLogoPath = $INI.'Branding-GUI'.HeaderLogo
    $clickURL       = $INI.'Branding-GUI'.HeaderLogoURL
    $footerWebseite = $INI.'Branding-GUI'.FooterWebseite
    $footerTemplate = $INI.'Branding-GUI'.GUI_Header
    # GUI-Sektion könnte fehlen, daher Prüfung hinzufügen
    $btnForeColor  = if ($INI.GUI -and $INI.GUI.ButtonForeColor) { $INI.GUI.ButtonForeColor } else { "#000000" }
    $btnFont       = if ($INI.GUI -and $INI.GUI.ButtonFont) { $INI.GUI.ButtonFont } else { "Segoe UI" }
    $btnFontSize   = if ($INI.GUI -and $INI.GUI.ButtonFontSize) { $INI.GUI.ButtonFontSize } else { 10 }
    
    $btnBC1 = if ($INI.GUI -and $INI.GUI.ButtonBackColor1) { $INI.GUI.ButtonBackColor1 } else { "#F0F0F0" }
    $btnBC2 = if ($INI.GUI -and $INI.GUI.ButtonBackColor2) { $INI.GUI.ButtonBackColor2 } else { "#F0F0F0" }
    $btnBC3 = if ($INI.GUI -and $INI.GUI.ButtonBackColor3) { $INI.GUI.ButtonBackColor3 } else { "#F0F0F0" }
    $btnBC4 = if ($INI.GUI -and $INI.GUI.ButtonBackColor4) { $INI.GUI.ButtonBackColor4 } else { "#F0F0F0" }
    
    # placeholders: {ScriptVersion}, {LastUpdate}, {Author}
    $scriptVersion = $INI.ScriptInfo.ScriptVersion
    $lastUpdate    = $INI.ScriptInfo.LastUpdate
    $author        = $INI.ScriptInfo.Author
    if ($footerTemplate) {
        if ($scriptVersion) { $footerTemplate = $footerTemplate -replace '\{ScriptVersion\}', $scriptVersion }
        if ($lastUpdate)    { $footerTemplate = $footerTemplate -replace '\{LastUpdate\}',    $lastUpdate }
        if ($author)        { $footerTemplate = $footerTemplate -replace '\{Author\}',        $author }
    }
    else {
        $footerTemplate = "easyLAPS"
    }

    # (2) Hauptfenster
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $appName
    $form.StartPosition = "CenterScreen"
    $form.Size = [System.Drawing.Size]::new(1000,640)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    # HEADER (70 px)
    $panelHeader = [System.Windows.Forms.Panel]::new()
    $panelHeader.Dock = 'Top'
    $panelHeader.Height = 70
    try {
        $panelHeader.BackColor = [System.Drawing.ColorTranslator]::FromHtml($themeColorStr)
    }
    catch {
        $panelHeader.BackColor = [System.Drawing.Color]::DarkGray
    }
    $form.Controls.Add($panelHeader)

    # Label=Left
    $lblAppName = [System.Windows.Forms.Label]::new()
    $lblAppName.Dock = 'Left'
    $lblAppName.Width = 220
    try {
        # Korrektur: headerFontName durch fontFamily ersetzt
        $lblAppName.Font = [System.Drawing.Font]::new($fontFamily,[float]$fontSize * 1.4,[System.Drawing.FontStyle]::Bold)
    }
    catch {
        $lblAppName.Font = [System.Drawing.Font]::new("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
    }
    $lblAppName.Text = $appName
    $lblAppName.TextAlign = 'MiddleLeft'
    $lblAppName.Padding = '10,0,0,0'
    $panelHeader.Controls.Add($lblAppName)

    # "ComputerName:" + Input + Button
    $lblTopComp = [System.Windows.Forms.Label]::new()
    $lblTopComp.Text = "ComputerName:"  # ENGLISH: ComputerName label remains in English
    $lblTopComp.AutoSize = $true
    $lblTopComp.Location = [System.Drawing.Point]::new(230,25)
    $panelHeader.Controls.Add($lblTopComp)

    $txtTopComp = [System.Windows.Forms.TextBox]::new()
    $txtTopComp.Location = [System.Drawing.Point]::new(330,22)
    $txtTopComp.Size = [System.Drawing.Size]::new(300,25)
    $panelHeader.Controls.Add($txtTopComp)

    $btnTopLoad = [System.Windows.Forms.Button]::new()
    $btnTopLoad.Text = "Load LAPS Password"  # ENGLISH - Button label in English
    $btnTopLoad.Location = [System.Drawing.Point]::new(640,20)
    $btnTopLoad.Size = [System.Drawing.Size]::new(120,28)
    $panelHeader.Controls.Add($btnTopLoad)

    # Logo=Right
    if ($headerLogoPath -and (Test-Path $headerLogoPath)) {
        $picLogo = [System.Windows.Forms.PictureBox]::new()
        $picLogo.Dock = 'Right'
        $picLogo.SizeMode = 'StretchImage'
        $picLogo.Width = 200
        $picLogo.Height = 70
        $picLogo.ImageLocation = $headerLogoPath
        $picLogo.Cursor = [System.Windows.Forms.Cursors]::Hand
        $picLogo.Add_Click({ if ($clickURL) { Open-URLInBrowser $clickURL } })
        $panelHeader.Controls.Add($picLogo)
    }

    # FOOTER (30 px)
    $panelFooter = [System.Windows.Forms.Panel]::new()
    $panelFooter.Dock = 'Bottom'
    $panelFooter.Height = 30
    try {
        $panelFooter.BackColor = [System.Drawing.ColorTranslator]::FromHtml($themeColorStr)
    }
    catch {
        $panelFooter.BackColor = [System.Drawing.Color]::DarkGray
    }
    $form.Controls.Add($panelFooter)

    $lblFooter = [System.Windows.Forms.Label]::new()
    $lblFooter.AutoSize = $true
    $lblFooter.Text = $footerTemplate
    $lblFooter.Font = [System.Drawing.Font]::new($fontFamily,8)
    $lblFooter.Location = [System.Drawing.Point]::new(10,5)
    $lblFooter.Cursor = [System.Windows.Forms.Cursors]::Hand
    $lblFooter.Add_Click({ if ($clickURL) { Open-URLInBrowser $clickURL } })
    $panelFooter.Controls.Add($lblFooter)

    if ($footerWebseite) {
        $lblFooter2 = [System.Windows.Forms.Label]::new()
        $lblFooter2.AutoSize = $true
        $lblFooter2.Text = $footerWebseite
        $lblFooter2.Font = [System.Drawing.Font]::new($fontFamily,8)
        $lblFooter2.Location = [System.Drawing.Point]::new(300,5)
        $lblFooter2.Cursor = [System.Windows.Forms.Cursors]::Hand
        $lblFooter2.Add_Click({ if ($clickURL) { Open-URLInBrowser $clickURL } })
        $panelFooter.Controls.Add($lblFooter2)
    }

    # MAIN
    $panelMain = [System.Windows.Forms.Panel]::new()
    $panelMain.Dock = 'Fill'
    $form.Controls.Add($panelMain)

    # Left: LAPS-Funktionen
    $panelLeft = [System.Windows.Forms.Panel]::new()
    $panelLeft.Parent = $panelMain
    $panelLeft.Location = [System.Drawing.Point]::new(10,80)
    $panelLeft.Size = [System.Drawing.Size]::new(430,480)

    # Right: Computer-Liste
    $panelRight = [System.Windows.Forms.Panel]::new()
    $panelRight.Parent = $panelMain
    $panelRight.Location = [System.Drawing.Point]::new(460,80)
    $panelRight.Size = [System.Drawing.Size]::new(570,480)
    $panelRight.BackColor = [System.Drawing.Color]::WhiteSmoke

    # GroupBox
    $grpLAPS = [System.Windows.Forms.GroupBox]::new()
    $grpLAPS.Text = "Local Administrator Password Solution"
    $grpLAPS.Font = [System.Drawing.Font]::new($fontFamily,9)
    $grpLAPS.Location = [System.Drawing.Point]::new(0,0)
    $grpLAPS.Size = [System.Drawing.Size]::new(430,480)
    try {
        $grpLAPS.BackColor = [System.Drawing.ColorTranslator]::FromHtml($boxColorStr)
    }
    catch {
        $grpLAPS.BackColor = [System.Drawing.Color]::LightGray
    }
    try {
        $grpLAPS.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($rahmenColorStr)
    }
    catch {
        $grpLAPS.ForeColor = [System.Drawing.Color]::DarkBlue
    }
    $panelLeft.Controls.Add($grpLAPS)

    $yPos = 30
    # Label "Ablauf LAPS-Kennwort"
    $lblAblaufAktuell = [System.Windows.Forms.Label]::new()
    $lblAblaufAktuell.Text = "Expiration of the current LAPS password:"
    $lblAblaufAktuell.AutoSize = $true
    $lblAblaufAktuell.Location = [System.Drawing.Point]::new(10,$yPos)
    $grpLAPS.Controls.Add($lblAblaufAktuell)
    $yPos += 25

    $txtCurrentExpire = [System.Windows.Forms.TextBox]::new()
    $txtCurrentExpire.ReadOnly = $true
    $txtCurrentExpire.Location = [System.Drawing.Point]::new(10,$yPos)
    $txtCurrentExpire.Size = [System.Drawing.Size]::new(400,25)
    $grpLAPS.Controls.Add($txtCurrentExpire)
    $yPos += 40

    # Button "Jetzt ablaufen lassen"
    $btnExpireNow = [System.Windows.Forms.Button]::new()
    $btnExpireNow.Text = "Force immediate password reset (Reset-LapsPassword)"
    $btnExpireNow.Location = [System.Drawing.Point]::new(10,$yPos)
    $btnExpireNow.Size = [System.Drawing.Size]::new(400,25)
    $grpLAPS.Controls.Add($btnExpireNow)
    $yPos += 45

    # LAPS-Admin-Konto
    $lblAdminAcct = [System.Windows.Forms.Label]::new()
    $lblAdminAcct.Text = "Local LAPS Administrator Account:"
    $lblAdminAcct.AutoSize = $true
    $lblAdminAcct.Location = [System.Drawing.Point]::new(10,$yPos)
    $grpLAPS.Controls.Add($lblAdminAcct)
    $yPos += 25

    $txtAdminAcct = [System.Windows.Forms.TextBox]::new()
    $txtAdminAcct.ReadOnly = $true
    $txtAdminAcct.Location = [System.Drawing.Point]::new(10,$yPos)
    $txtAdminAcct.Size = [System.Drawing.Size]::new(400,25)
    $grpLAPS.Controls.Add($txtAdminAcct)
    $yPos += 45

    # LAPS-Passwort
    $lblAdminPwd = [System.Windows.Forms.Label]::new()
    $lblAdminPwd.Text = "LAPS Password:"
    $lblAdminPwd.AutoSize = $true
    $lblAdminPwd.Location = [System.Drawing.Point]::new(10,$yPos)
    $grpLAPS.Controls.Add($lblAdminPwd)
    $yPos += 25

    $txtAdminPwd = [System.Windows.Forms.TextBox]::new()
    $txtAdminPwd.ReadOnly = $true
    $txtAdminPwd.Location = [System.Drawing.Point]::new(10,$yPos)
    $txtAdminPwd.Size = [System.Drawing.Size]::new(400,25)
    $txtAdminPwd.UseSystemPasswordChar = $true
    $grpLAPS.Controls.Add($txtAdminPwd)
    $yPos += 40

    $btnCopyPwd = [System.Windows.Forms.Button]::new()
    $btnCopyPwd.Text = "Copy Password"
    $btnCopyPwd.Location = [System.Drawing.Point]::new(10,$yPos)
    $btnCopyPwd.Size = [System.Drawing.Size]::new(190,30)
    $grpLAPS.Controls.Add($btnCopyPwd)

    $btnShowPwd = [System.Windows.Forms.Button]::new()
    $btnShowPwd.Text = "Show Password"
    $btnShowPwd.Location = [System.Drawing.Point]::new(220,$yPos)
    $btnShowPwd.Size = [System.Drawing.Size]::new(190,30)
    $grpLAPS.Controls.Add($btnShowPwd)
    $yPos += 50

    # Right => Computer-Liste
    $lblList = [System.Windows.Forms.Label]::new()
    $lblList.Text = "Computers (Windows LAPS 2.0):"
    $lblList.Font = [System.Drawing.Font]::new("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)
    $lblList.AutoSize = $true
    $lblList.Location = [System.Drawing.Point]::new(10,10)
    $panelRight.Controls.Add($lblList)

    $lstComputers = [System.Windows.Forms.ListBox]::new()
    $lstComputers.Location = [System.Drawing.Point]::new(10,40)
    $lstComputers.Size = [System.Drawing.Size]::new(500,370)
    $panelRight.Controls.Add($lstComputers)

    $btnRefresh = [System.Windows.Forms.Button]::new()
    $btnRefresh.Text = "Refresh List"
    $btnRefresh.Location = [System.Drawing.Point]::new(10,420)
    $btnRefresh.Size = [System.Drawing.Size]::new(120,30)
    $panelRight.Controls.Add($btnRefresh)

    # (C) Button-Style
    function Apply-ButtonStyle {
        param([System.Windows.Forms.Button]$Button,[string]$bgColor)
        if (-not $Button) {
            Throw "Button-Parameter darf nicht leer sein."
        }
        try {
            $Button.BackColor = [System.Drawing.ColorTranslator]::FromHtml($bgColor)
        }
        catch {
            $Button.BackColor = [System.Drawing.Color]::LightGray
        }
        try {
            $Button.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($btnForeColor)
        }
        catch {
            $Button.ForeColor = [System.Drawing.Color]::Black
        }
        try {
            # Korrektur: Verwende $btnFont und $btnFontSize anstelle von $fontFamily und $fontSize
            $Button.Font = [System.Drawing.Font]::new($btnFont,[float]$btnFontSize,[System.Drawing.FontStyle]::Regular)
        }
        catch {
            $Button.Font = [System.Drawing.Font]::new("Segoe UI",10,[System.Drawing.FontStyle]::Regular)
        }
    }
    Apply-ButtonStyle $btnRefresh   $btnBC1
    Apply-ButtonStyle $btnShowPwd   $btnBC2
    Apply-ButtonStyle $btnCopyPwd   $btnBC3
    Apply-ButtonStyle $btnExpireNow $btnBC4
    Apply-ButtonStyle $btnTopLoad   $btnBC1

    # ------------------------------------------------
    # (D) Event-Logik
    # ------------------------------------------------
    function Refresh-ComputerList {
        if ($Global:DebugEnabled) {
            Write-DebugMessage "Refreshing computer list in the GUI..."
        }
        $lstComputers.Items.Clear()
        $comps = Load-Computers
        if ($comps.Count -eq 0) {
            Write-Warning "Keine Computerobjekte gefunden."
        }
        foreach ($c in $comps) {
            [void]$lstComputers.Items.Add($c)
        }
        if ($Global:DebugEnabled) {
            Write-Log "Computer list refreshed."
        }
    }
    Refresh-ComputerList

    $btnRefresh.Add_Click({
        Refresh-ComputerList
    })

    function Show-LapsNewData {
        param([string]$ComputerName)
        if ([string]::IsNullOrWhiteSpace($ComputerName)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify/select a computer.")
            return
        }
        try {
            $pwInfo = Get-LapsPW -ComputerName $ComputerName
            # => .Password, .ExpirationTimestamp
            $txtAdminPwd.Text  = $pwInfo.Password
            $txtCurrentExpire.Text = if ($pwInfo.ExpirationTimestamp) {
                ($pwInfo.ExpirationTimestamp).ToString("yyyy-MM-dd HH:mm")
            } else { "" }

            # Admin-Konto => not directly stored, 
            # du koenntest z.B. "Get-LapsADComputer -Identity $ComputerName" => .ManagedLocalAccounts, etc.
            $txtAdminAcct.Text = "Local Admin (Windows LAPS 2.0)"

        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error retrieving LAPS: $($_.Exception.Message)",
                "Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Klick in ListBox => LAPS-Daten
    $lstComputers.Add_SelectedIndexChanged({
        $val = $lstComputers.SelectedItem
        if ($val) {
            Show-LapsNewData $val
            $txtTopComp.Text = $val
        }
    })

    # Button "LAPS auslesen"
    $btnTopLoad.Add_Click({
        $comp = $txtTopComp.Text.Trim()
        Show-LapsNewData $comp
    })

    # Kennwort kopieren
    $btnCopyPwd.Add_Click({
        if ($txtAdminPwd.Text) {
            Set-Clipboard $txtAdminPwd.Text
            [System.Windows.Forms.MessageBox]::Show("Password copied to clipboard.")
        }
    })

    # Kennwort anzeigen
    $btnShowPwd.Add_Click({
        if ($txtAdminPwd.UseSystemPasswordChar -eq $true) {
            $txtAdminPwd.UseSystemPasswordChar = $false
            $btnShowPwd.Text = "Hide Password"
        }
        else {
            $txtAdminPwd.UseSystemPasswordChar = $true
            $btnShowPwd.Text = "Show Password"
        }
    })

    # "Jetzt ablaufen lassen" => Reset-LapsPassword
    $btnExpireNow.Add_Click({
        $compName = $txtTopComp.Text.Trim()
        if (-not $compName) {
            if ($lstComputers.SelectedItem) {
                $compName = $lstComputers.SelectedItem
            }
        }
        if (-not $compName) {
            [System.Windows.Forms.MessageBox]::Show("Please specify or select a computer.")
            return
        }
        try {
            Reset-LapsPWNow -ComputerName $compName
            [System.Windows.Forms.MessageBox]::Show("A new LAPS password is set immediately (Reset-LapsPassword).",
                "Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            Show-LapsNewData $compName
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error with Reset-LapsPassword: $($_.Exception.Message)",
                "Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # ------------------------------------------------
    # Start the GUI
    # ------------------------------------------------
    [void]$form.ShowDialog()
}
#endregion

#region [2.6 | Script Start - Skriptstart]
# ENGLISH - Reads the INI file and launches the main GUI form
# GERMAN - Liest die INI-Datei und startet das Haupt-GUI-Formular
try {
    $scriptPath = Split-Path -Parent $PSCommandPath
    $iniFile    = Join-Path $scriptPath "easyLAPS.ini"
    if (-not (Test-Path $iniFile)) {
        Throw "INI nicht gefunden unter: $iniFile"
    }
    $INI = Read-INIFile $iniFile
    
    # Stellen Sie sicher, dass GUI-Sektion existiert, wenn nicht, erstellen
    if (-not $INI.GUI) {
        $INI.GUI = New-Object System.Collections.Specialized.OrderedDictionary
    }
}
catch {
    Write-Error "Fehler beim Laden der INI: $($_.Exception.Message)"
    return
}

Show-LAPSForm -INI $INI
#endregion

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAFDsA7XqehyRAU
# UxGXSRXittntk3YWptTi/MW/uphrc6CCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCBoLoNtE7vXsbcaCxQCr4BaAdQwakUsElwkkY2hZPEowzANBgkqhkiG
# 9w0BAQEFAASCAQBoIs3pxa5w85JBlvjNuYz6EFqFn+v+DvMewCD+LxPcQWMU4XrS
# j7yA8iS/bynOprZTcWQ757aI/EoW4rp9guP/FmdJXs8IMIRJxnc4QxjLLMeL8Kbv
# inTDC+A/BlIW8BaPxFW4x9lmqgrMg4KfUyviqzSpG81EZC98vjnIF9dJYTVod9k7
# 5ED0ZigItaxtmOAxiAdGWvo+7ooxeWpSUuvNjroNSfcvkBNN2B8MEkfQ7X9k7cis
# bloGSeqv/kcQYDYyL2yuI7XVG1FA2nNym6AHJfoHhE2jZmIUAznHBXVQzmBO2zw4
# YFF5Cjn4k4ZNJPQXMJ/hpinwRTCC3E7eQvpaoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1MFowLwYJKoZIhvcNAQkEMSIEIMLXeiqJMkqItnMJycviEF93j5yy
# GEYHt8G9/iUza/wMMA0GCSqGSIb3DQEBAQUABIICAARPHl03r16wEAk0J9ZW1+Tc
# jMDGrqDafJQlVh1NrcJ9igfZti1mLjO+wFBRM/d7hjwll4uHtGnqrkRrxHgX6mFL
# UaGuKU6ek/hTmwk6BbvR1GZGyPJqARJvtbubHAdIqTn/fM27+ynrAAGHYReCEigD
# B1L+8Ti6P85jGr/zWjnOgcrA7Ot9mEb/dv3zy8Zy6OWLUnQHa9tz6uHJuPT8601B
# 4TGECq9GVj4OPs4tMiNe+0+QXXI1ZyIXum7DFsbd41lwoouYwdcYqxlsbCgsDDS9
# zaJ3cnQZX/mm/uqk2wOYyOgUzt0mSEhjkdD86buCPMpsKCuDYZdrJ4umVEUG9+iE
# Cpudm3rzkqFTwVjeqfieeAB2oj8i1JYsUMejSR4qtvFgX0Rppt7e07bqobMvXGnF
# p86mqmVyCthdW8v2MLLn7AV/AnlcY3GuqboztePs0xr9j8mjFI7ebxhDGwciLysO
# ZSCazl6QcuS625oh5g+0Kz2yjDAXE4U/nPhkvGMG1llhpL/L0AcLDZ9J3T0Tb2on
# 85oojk0VN8FyEJUhoDi5ol8VUZCulkleBo6wtd66a5SWH4xDKcthyGN9oLM6XctT
# yGRzrOIpRfjTjS3A1bbL/xY/m0vneccURzV8IwpkrgzNIHQZM8E2eh+QLPYUkMpX
# AdXxdT84MdfpmRb/g0s+
# SIG # End signature block
