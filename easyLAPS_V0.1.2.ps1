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

# ----------------------------------------------------
# 1) INI-Datei einlesen
# ----------------------------------------------------
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
        Throw "Fehler beim Lesen der INI-Datei: $_"
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

# ----------------------------------------------------
# 2) Neue LAPS-Funktionen
# ----------------------------------------------------
function Load-Computers {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $comps = Get-ADComputer -Filter * | Sort-Object Name
        if (-not $comps) {
            Write-Warning "Keine Computerobjekte gefunden."
            return @()
        }
        return $comps | Select-Object -ExpandProperty Name
    }
    catch {
        Write-Warning "Fehler beim Laden der Computerobjekte: $($_.Exception.Message)"
        return @()
    }
}

function Get-LapsPW {
    param([string]$ComputerName)
    if (-not $ComputerName) {
        Throw "ComputerName darf nicht leer sein."
    }
    try {
        Import-Module LAPS -ErrorAction Stop
        $pwInfo = Get-LapsADPassword -Identity $ComputerName -AsPlainText -ErrorAction Stop
        if (-not $pwInfo) {
            Throw "Keine neuen LAPS-Daten für '$ComputerName'."
        }
        return $pwInfo
    }
    catch {
        Throw "Fehler beim Lesen des LAPS-Kennworts: $($_.Exception.Message)"
    }
}

function Reset-LapsPWNow {
    param([string]$ComputerName)
    if (-not $ComputerName) {
        Throw "ComputerName darf nicht leer sein."
    }
    try {
        Import-Module LAPS -ErrorAction Stop
        Reset-LapsPassword -Identity $ComputerName -ErrorAction Stop
    }
    catch {
        Throw "Fehler beim 'sofort ablaufen lassen': $($_.Exception.Message)"
    }
}

# ----------------------------------------------------
# 3) Hilfsfunktionen: Clipboard / URL
# ----------------------------------------------------
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

function Open-URLInBrowser {
    param([string]$URL)
    try {
        [System.Diagnostics.Process]::Start($URL) | Out-Null
    }
    catch {
        Write-Warning "URL konnte nicht geöffnet werden: $URL - $($_.Exception.Message)"
    }
}

# ----------------------------------------------------
# 4) GUI-Hauptfunktion
# ----------------------------------------------------
function Show-LAPSForm {
    param([hashtable]$INI)

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
    $lblTopComp.Text = "ComputerName:"
    $lblTopComp.AutoSize = $true
    $lblTopComp.Location = [System.Drawing.Point]::new(230,25)
    $panelHeader.Controls.Add($lblTopComp)

    $txtTopComp = [System.Windows.Forms.TextBox]::new()
    $txtTopComp.Location = [System.Drawing.Point]::new(330,22)
    $txtTopComp.Size = [System.Drawing.Size]::new(300,25)
    $panelHeader.Controls.Add($txtTopComp)

    $btnTopLoad = [System.Windows.Forms.Button]::new()
    $btnTopLoad.Text = "LAPS auslesen"
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
    $lblAblaufAktuell.Text = "Ablauf des aktuellen LAPS-Kennworts:"
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
    $btnExpireNow.Text = "Jetzt ablaufen lassen (Reset-LapsPassword)"
    $btnExpireNow.Location = [System.Drawing.Point]::new(10,$yPos)
    $btnExpireNow.Size = [System.Drawing.Size]::new(400,25)
    $grpLAPS.Controls.Add($btnExpireNow)
    $yPos += 45

    # LAPS-Admin-Konto
    $lblAdminAcct = [System.Windows.Forms.Label]::new()
    $lblAdminAcct.Text = "Lokales LAPS-Administratorkonto:"
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
    $lblAdminPwd.Text = "LAPS-Passwort:"
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
    $btnCopyPwd.Text = "Kennwort kopieren"
    $btnCopyPwd.Location = [System.Drawing.Point]::new(10,$yPos)
    $btnCopyPwd.Size = [System.Drawing.Size]::new(190,30)
    $grpLAPS.Controls.Add($btnCopyPwd)

    $btnShowPwd = [System.Windows.Forms.Button]::new()
    $btnShowPwd.Text = "Kennwort anzeigen"
    $btnShowPwd.Location = [System.Drawing.Point]::new(220,$yPos)
    $btnShowPwd.Size = [System.Drawing.Size]::new(190,30)
    $grpLAPS.Controls.Add($btnShowPwd)
    $yPos += 50

    # Right => Computer-Liste
    $lblList = [System.Windows.Forms.Label]::new()
    $lblList.Text = "Computer (Windows LAPS 2.0):"
    $lblList.Font = [System.Drawing.Font]::new("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)
    $lblList.AutoSize = $true
    $lblList.Location = [System.Drawing.Point]::new(10,10)
    $panelRight.Controls.Add($lblList)

    $lstComputers = [System.Windows.Forms.ListBox]::new()
    $lstComputers.Location = [System.Drawing.Point]::new(10,40)
    $lstComputers.Size = [System.Drawing.Size]::new(500,370)
    $panelRight.Controls.Add($lstComputers)

    $btnRefresh = [System.Windows.Forms.Button]::new()
    $btnRefresh.Text = "Liste aktualisieren"
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
        $lstComputers.Items.Clear()
        $comps = Load-Computers
        if ($comps.Count -eq 0) {
            Write-Warning "Keine Computerobjekte gefunden."
        }
        foreach ($c in $comps) {
            [void]$lstComputers.Items.Add($c)
        }
    }
    Refresh-ComputerList

    $btnRefresh.Add_Click({
        Refresh-ComputerList
    })

    function Show-LapsNewData {
        param([string]$ComputerName)
        if ([string]::IsNullOrWhiteSpace($ComputerName)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte einen Computer angeben/auswählen.")
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
            $txtAdminAcct.Text = "Lokaler Admin (Windows LAPS 2.0)"

        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim LAPS-Abruf: $($_.Exception.Message)",
                "Fehler",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
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
            [System.Windows.Forms.MessageBox]::Show("Kennwort in die Zwischenablage kopiert.")
        }
    })

    # Kennwort anzeigen
    $btnShowPwd.Add_Click({
        if ($txtAdminPwd.UseSystemPasswordChar -eq $true) {
            $txtAdminPwd.UseSystemPasswordChar = $false
            $btnShowPwd.Text = "Kennwort verstecken"
        }
        else {
            $txtAdminPwd.UseSystemPasswordChar = $true
            $btnShowPwd.Text = "Kennwort anzeigen"
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
            [System.Windows.Forms.MessageBox]::Show("Bitte einen Computer angeben oder auswählen.")
            return
        }
        try {
            Reset-LapsPWNow -ComputerName $compName
            [System.Windows.Forms.MessageBox]::Show("Neues LAPS-Passwort wird sofort gesetzt (Reset-LapsPassword).",
                "Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            Show-LapsNewData $compName
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler bei Reset-LapsPassword: $($_.Exception.Message)",
                "Fehler",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # ------------------------------------------------
    # Start the GUI
    # ------------------------------------------------
    [void]$form.ShowDialog()
}

# ----------------------------------------------------
# 5) Skriptstart
# ----------------------------------------------------
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
    Write-Error "Fehler beim Laden der INI: $_"
    return
}

Show-LAPSForm -INI $INI
