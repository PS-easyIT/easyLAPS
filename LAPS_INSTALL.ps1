<# 
Skript: NewWindowsLAPS_GUI_AutoOU.ps1
Beschreibung:
  - Dieses Skript nutzt das neue Windows LAPS (in aktuellen Windows‑Versionen integriert).
  - Funktionen:
      • LAPS installieren: Führt Update-LapsADSchema aus, um das AD‑Schema (falls nötig) zu erweitern.
      • LAPS konfigurieren: Liest automatisch den OU‑Teil des DistinguishedName des lokalen Computers aus und delegiert mittels Set-LapsADComputerSelfPermission die Rechte.
      • LAPS testen: Liest das LAPS‑Passwort für den Computer ab, dessen Namen im Eingabefeld angegeben wird.
      • Beenden: Schließt die GUI.
      
Voraussetzungen:
  - Windows LAPS (ab Windows Server 2019 / Windows 10 ab April 2023 Update)
  - PowerShell 5.1 oder höher
  - RSAT‑AD‑PowerShell (für AD‑Cmdlets)
  - Das Windows LAPS PowerShell-Modul (Name: LAPS) ist verfügbar
  - Der Computer ist in Active Directory registriert
#>

# .NET Assemblies laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Definition der Button-Styles
$buttonBackColor = [System.Drawing.ColorTranslator]::FromHtml("#0055AA")
$buttonForeColor = [System.Drawing.Color]::White
$buttonFont = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)

# Funktion zum automatischen Auslesen der OU des lokalen Computers
function Get-LocalComputerOU {
    try {
        # Importiere das AD-Modul (RSAT) – falls nicht schon geladen
        Import-Module ActiveDirectory -ErrorAction Stop
        $localAD = Get-ADComputer -Identity $env:COMPUTERNAME -Properties DistinguishedName
        if ($localAD -and $localAD.DistinguishedName) {
            # Entferne den CN-Teil, um die OU zu erhalten
            $ou = $localAD.DistinguishedName -replace '^CN=[^,]+,',''
            return $ou
        }
        else {
            throw "Kein AD-Eintrag für $env:COMPUTERNAME gefunden."
        }
    }
    catch {
        throw "Fehler beim Auslesen der OU: $_"
    }
}

# Erstellen des Hauptformulars
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows LAPS – Installation, Konfiguration & Test"
$form.Size = New-Object System.Drawing.Size(420,320)
$form.StartPosition = "CenterScreen"

# Installations-Button (Schema-Aktualisierung)
$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Location = New-Object System.Drawing.Point(20,20)
$btnInstall.Size = New-Object System.Drawing.Size(360,40)
$btnInstall.Text = "LAPS installieren (Schema aktualisieren)"
$btnInstall.BackColor = $buttonBackColor
$btnInstall.ForeColor = $buttonForeColor
$btnInstall.Font = $buttonFont
$btnInstall.Add_Click({
    Install-LAPS
})

# Konfigurations-Button (automatisch OU auslesen)
$btnConfigure = New-Object System.Windows.Forms.Button
$btnConfigure.Location = New-Object System.Drawing.Point(20,80)
$btnConfigure.Size = New-Object System.Drawing.Size(360,40)
$btnConfigure.Text = "LAPS konfigurieren (autom. OU auslesen)"
$btnConfigure.BackColor = $buttonBackColor
$btnConfigure.ForeColor = $buttonForeColor
$btnConfigure.Font = $buttonFont
$btnConfigure.Add_Click({
    try {
        $ou = Get-LocalComputerOU
        Configure-LAPS $ou
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Auslesen der OU:`n$_","Fehler",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Label und TextBox für Computername (für den Test)
$lblComputer = New-Object System.Windows.Forms.Label
$lblComputer.Location = New-Object System.Drawing.Point(20,140)
$lblComputer.Size = New-Object System.Drawing.Size(120,25)
$lblComputer.Text = "Computername:"
$lblComputer.Font = $buttonFont

$txtComputer = New-Object System.Windows.Forms.TextBox
$txtComputer.Location = New-Object System.Drawing.Point(150,140)
$txtComputer.Size = New-Object System.Drawing.Size(230,25)
$txtComputer.Text = $env:COMPUTERNAME
$txtComputer.Font = $buttonFont

# Test-Button
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Location = New-Object System.Drawing.Point(20,180)
$btnTest.Size = New-Object System.Drawing.Size(360,40)
$btnTest.Text = "LAPS testen (autom. auslesen)"
$btnTest.BackColor = $buttonBackColor
$btnTest.ForeColor = $buttonForeColor
$btnTest.Font = $buttonFont
$btnTest.Add_Click({
    Test-LAPS $txtComputer.Text
})

# Beenden-Button
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = New-Object System.Drawing.Point(20,240)
$btnExit.Size = New-Object System.Drawing.Size(360,40)
$btnExit.Text = "Beenden"
$btnExit.BackColor = $buttonBackColor
$btnExit.ForeColor = $buttonForeColor
$btnExit.Font = $buttonFont
$btnExit.Add_Click({
    $form.Close()
})

# Steuerelemente zum Formular hinzufügen
$form.Controls.Add($btnInstall)
$form.Controls.Add($btnConfigure)
$form.Controls.Add($lblComputer)
$form.Controls.Add($txtComputer)
$form.Controls.Add($btnTest)
$form.Controls.Add($btnExit)

# Funktion: LAPS installieren (Schema-Aktualisierung)
function Install-LAPS {
    Write-Output "Starte Schema-Aktualisierung für Windows LAPS..."
    try {
        Import-Module LAPS -ErrorAction Stop
        Update-LapsADSchema -Verbose -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("Windows LAPS Schema-Aktualisierung erfolgreich.","Erfolg",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Schema-Aktualisierung:`n$_","Fehler",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Funktion: LAPS konfigurieren (Delegierung der Rechte)
function Configure-LAPS {
    param(
        [string]$OU
    )
    Write-Output "Starte Konfiguration von Windows LAPS für OU: $OU ..."
    try {
        Import-Module LAPS -ErrorAction Stop
        # Setzt die Rechte für die ermittelte OU
        Set-LapsADComputerSelfPermission -Identity $OU -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("Windows LAPS wurde erfolgreich konfiguriert für die OU:`n$OU","Erfolg",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der LAPS-Konfiguration:`n$_","Fehler",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Funktion: LAPS testen (Passwort aus AD auslesen)
function Test-LAPS {
    param(
        [string]$ComputerName
    )
    Write-Output "Führe Windows LAPS Test für $ComputerName durch..."
    $results = @()
    try {
        Import-Module LAPS -ErrorAction Stop
        # Abruf des LAPS-Passworts mittels -Identity (bei Windows LAPS)
        $lapsResult = Get-LapsADPassword -Identity $ComputerName -AsPlainText -ErrorAction Stop
        if ($lapsResult -and $lapsResult.Password) {
            $results += "LAPS Passwort für $ComputerName gefunden:"
            $results += "Account: $($lapsResult.Account)"
            $results += "Passwort: $($lapsResult.Password)"
            $results += "Aktualisiert: $($lapsResult.PasswordUpdateTime)"
            $results += "Expires: $($lapsResult.ExpirationTimestamp)"
        }
        else {
            $results += "Kein LAPS Passwort für $ComputerName gefunden."
        }
    }
    catch {
        $results += "Fehler beim LAPS-Test: $_"
    }
    [System.Windows.Forms.MessageBox]::Show(($results -join "`n"),"Testergebnisse",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
}

# Formular anzeigen
$form.Add_Shown({ $form.Activate() })
[void] $form.ShowDialog()
