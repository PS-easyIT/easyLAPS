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

# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBtWcmuAKbbLg/9
# YthaCOHrBOAqkPaje8ho+Uey4EY5haCCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3H
# rFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAw
# MFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU
# 7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR
# +2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwE
# u7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Za
# zch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW3
# 5xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gd
# FpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rq
# BvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vH
# espYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QE
# PHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1
# Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9
# EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk
# 97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2
# UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71
# WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQf
# jXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noD
# js6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxi
# Df06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/
# D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8Ml
# uDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG
# 2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8
# hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIFEDCCBQwCAQEwNDAgMR4wHAYDVQQD
# DBVQaGluSVQtUFNzY3JpcHRzX1NpZ24CEHePOzJf0KCMSL6wELasExMwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgpMUyCcbKA2ezM5CxN7d5RgM7Na1Jra8N8OElvsSyh2Uw
# DQYJKoZIhvcNAQEBBQAEggEAH9fnfXOTwqKdNeCEYZn7nQ/Y50w8/3FbyUl8cqpJ
# Lu4H9SbM3OPic2X87mSGD3agigZk1HTImQscf9JkdAPDRE0HWV4lebztazmn+Nbb
# QkOU1m0Sfb7mOsdoWN+4F3tftbe0/RMFWVc4dAm0sq4tl1fMLyuIDsRL8V92rid4
# P8PXzge4WaQnlzJDNstzkROU/aQOE9xXSDzixUha8W0xu1meP4HeEARiwCfb7KCH
# X0+DXHMCrD5z/mFcV7NYwjzYdbKNnAW+nrcoE3Xa/PiuYlbv3yAy/7+0tVrKnHNS
# GSxxC29xPwJd+KcWgp65geiFsTCW2AA2LcEOXnUQkUaB9KGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAwMTdaMC8GCSqGSIb3DQEJBDEiBCAn1NrhPW+1
# CMSlvdIDReyjKsHr4yCtXkfk/qwklsELiDANBgkqhkiG9w0BAQEFAASCAgBOfqBh
# GoczPsfEM5pVKrdbBHEdGTRLjatfR/w+HWpCos1tpJwC7+iaDVKBEqKoc5WFL/AI
# 44Z/i9C59RWpqGKJKlqBgELbemeZMzjHFGmLwdEuMI1tUfkcDqHib24q1Suh0JDw
# VEIdcWvlzcIMMtiCPbXfJboCFyd2S6VesE5dufiYHtv6fpeQN8wq+g8zN4arvDm5
# 2R4QaNA+MhIMXCM5LBIKnl7JOhiBn9BzmWIJ8xZdifVIHVK3yorCVKl4W6YirecF
# /W9B49OzOXl/l1gBIM63bkOKO64DTBdPrpsxFwVbZyOJDNv11kDAWlcIcil5Smyn
# EwUp6XrVnGIr9N83tuPxVUxV84uxe8Xz4ACPRcv1NUNobRMx8SNV39TTk+K5VEbN
# 7vl2AYtnghxruKOW4PGc1akMNmE0MOFAEqtxaU9VY4EC7DyOAiHfrRtHh+Gjuk+N
# 5Os8m+B3y7nOatacLZ3oRNBp4n/QI9nzPYmToe7dQnxW6PZrrr+WcNfzUVDYbePD
# flF632ZO2Vj7h0qQxuwx52S9B08SkDthh09/0n/nYIsPq4XhHAgLJZxOfkbtWkwz
# KYHnK3xd2ut1iF2IEQd7zf8OOWL76tGjwkmzabfOiuaBQSkYiGlQgvw+nHAuUfZL
# vuE6uyphPM0TgSFVZDc2jUV0uEAayeynpYhgkA==
# SIG # End signature block
