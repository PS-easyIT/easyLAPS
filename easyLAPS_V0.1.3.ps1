#region [1.0 | Variablen und Konfiguration]
# Skriptinformationen
$script:ScriptInfo = @{
    AppName = "easyLAPS"
    ScriptVersion = "0.1.3"
    Author = "easyIT"
    ThemeColor = "#0078D7"
    LastUpdate = (Get-Date -Format "yyyy-MM-dd")
    Company = "easyIT GmbH"
    Website = "https://www.easyit.de"
    Description = "Local Admin Password Solution fuer Windows-Umgebungen"
    Copyright = "© $(Get-Date -Format 'yyyy') easyIT GmbH"
}

# Branding-GUI Einstellungen
$script:BrandingGUI = @{
    APPName = $script:ScriptInfo.AppName
    FontFamily = "Segoe UI"
    HeaderLogo = "$PSScriptRoot\resources\logo.png"
    HeaderLogoURL = $script:ScriptInfo.Website
    FooterWebseite = $script:ScriptInfo.Website
    ThemeColor = $script:ScriptInfo.ThemeColor
}

# Pfade und Einstellungen
$script:logPath = Join-Path $PSScriptRoot "Logs"
$script:logFile = "$($script:ScriptInfo.AppName)_$(Get-Date -Format 'yyyyMMdd').log"
$script:LogFilePath = Join-Path $script:logPath $script:logFile
$script:registryPath = "HKCU:\Software\easyIT\$($script:ScriptInfo.AppName)"

# Debug-Einstellungen
$script:debugMode = $false
$script:logLevel = "Info" # Mögliche Werte: Error, Warning, Info, Debug
$script:logRotationDays = 30 # Logs älter als X Tage werden gelöscht

# Stellen Sie sicher, dass der Log-Ordner existiert
if (-not (Test-Path -Path $script:logPath)) {
    try {
        New-Item -Path $script:logPath -ItemType Directory -Force | Out-Null
        Write-Host "Log-Verzeichnis erstellt: $($script:logPath)" -ForegroundColor Green
    }
    catch {
        Write-Host "Fehler beim Erstellen des Log-Verzeichnisses: $_" -ForegroundColor Red
    }
}

# Initialisiere Registry-Einstellungen
try {
    if (-not (Test-Path -Path $script:registryPath)) {
        New-Item -Path $script:registryPath -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "Debug" -Value 0 -PropertyType DWORD -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "AppName" -Value $script:ScriptInfo.AppName -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "Version" -Value $script:ScriptInfo.ScriptVersion -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "ThemeColor" -Value $script:ScriptInfo.ThemeColor -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "LogPath" -Value $script:logPath -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $script:registryPath -Name "LogRotationDays" -Value $script:logRotationDays -PropertyType DWORD -Force | Out-Null
    }
    
    # Lade Debug-Einstellungen aus Registry
    $script:debugMode = (Get-ItemProperty -Path $script:registryPath -Name "Debug" -ErrorAction SilentlyContinue).Debug -eq 1
}
catch {
    Write-Host "Fehler beim Initialisieren der Registry-Einstellungen: $_" -ForegroundColor Red
}
#endregion


#region [2.0 | Debug und Logging]
function Write-DebugMessage {
    param (
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$false)][ValidateSet("Info", "Warning", "Error", "Success")][string]$Level = "Info",
        [Parameter(Mandatory=$false)][switch]$NoNewLine
    )
    
    # Wenn Debug-Modus nicht aktiviert ist, nichts tun
    if (-not $script:DebugMode) { return }
    
    # Zeitstempel für die Ausgabe
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Farben je nach Meldungstyp
    $foregroundColor = switch ($Level) {
        "Error"   { "Red" }
        "Warning" { "Yellow" }
        "Success" { "Green" }
        "Info"    { "Cyan" }
        default   { "White" }
    }
    
    # Ausgabe formatieren
    $formattedMessage = "[$timestamp] [$Level] $Message"
    
    # Ausgabe in Konsole
    if ($NoNewLine) {
        Write-Host $formattedMessage -ForegroundColor $foregroundColor -NoNewline
    } else {
        Write-Host $formattedMessage -ForegroundColor $foregroundColor
    }
}
#endregion

#region [2.1 | Logging]
function Write-Log {
    param (
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$false)][string]$Level = "Info",
        [Parameter(Mandatory=$false)][string]$LogFilePath = $script:LogFilePath
    )
    
    try {
        # Log-Rotation durchführen
        Invoke-LogRotation

        # Sicherstellen, dass der Log-Ordner existiert
        $logDir = Split-Path -Path $LogFilePath -Parent
        if (-not (Test-Path -Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        # Zeitstempel für Log-Eintrag
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # Log-Eintrag formatieren
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # In Datei schreiben
        Add-Content -Path $LogFilePath -Value $logEntry -Encoding UTF8
    }
    catch {
        # Fallback wenn Logging fehlschlägt
        $errorMsg = $_.Exception.Message
        Write-Host "FEHLER BEIM LOGGING: $errorMsg" -ForegroundColor Red
        
        # Versuche in temporäre Datei zu schreiben
        try {
            $tempLogPath = Join-Path -Path $env:TEMP -ChildPath "easyLAPS_error.log"
            Add-Content -Path $tempLogPath -Value "[$timestamp] [ERROR] Logging-Fehler: $errorMsg" -Encoding UTF8
            Add-Content -Path $tempLogPath -Value "[$timestamp] [INFO] Originale Nachricht: $Message" -Encoding UTF8
        }
        catch {
            # Wenn auch das fehlschlägt, können wir nichts mehr tun
        }
    }
}

function Invoke-LogRotation {
    try {
        # Prüfe, ob Log-Rotation aktiv ist
        $rotationDays = Get-RegistryValue -Path $script:registryPath -Name "LogRotationDays" -DefaultValue $script:logRotationDays
        
        if ($rotationDays -le 0) { return }
        
        # Berechne Cutoff-Datum
        $cutoffDate = (Get-Date).AddDays(-$rotationDays)
        
        # Finde alle Log-Dateien älter als der Cutoff
        $oldLogs = Get-ChildItem -Path $script:logPath -Filter "*.log" | 
                   Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        # Lösche alte Logs
        foreach ($log in $oldLogs) {
            Remove-Item -Path $log.FullName -Force
        }
    }
    catch {
        # Fehler beim Log-Rotation - nur in Konsole ausgeben, nicht ins Log schreiben (Vermeidung von Rekursion)
        Write-Host "Fehler bei Log-Rotation: $($_.Exception.Message)" -ForegroundColor Red
    }
}
#endregion

#region [2.2 | Fehlerbehandlung]
function Get-FormattedError {
    param (
        [Parameter(Mandatory=$true)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [Parameter(Mandatory=$false)][string]$DefaultText = "Ein unbekannter Fehler ist aufgetreten."
    )
    
    if ($null -eq $ErrorRecord) { return $DefaultText }
    
    try {
        $errorMessage = $ErrorRecord.Exception.Message
        $errorPosition = $ErrorRecord.InvocationInfo.PositionMessage
        $errorType = $ErrorRecord.Exception.GetType().Name
        
        return "[$errorType] $errorMessage`n$errorPosition"
    }
    catch {
        return $DefaultText
    }
}

function Show-ErrorMessage {
    param (
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$false)][string]$Title = "Fehler",
        [Parameter(Mandatory=$false)][System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    
    $detailedMessage = $Message
    
    if ($ErrorRecord) {
        $errorDetails = Get-FormattedError -ErrorRecord $ErrorRecord
        $detailedMessage += "`n`nDetails: $errorDetails"
    }
    
    # Fehler loggen
    Write-Log -Message $detailedMessage -Type "Error"
    
    # MessageBox anzeigen
    [System.Windows.MessageBox]::Show(
        $detailedMessage,
        $Title,
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
}
#endregion

#region [2.3 | Hilfsfunktionen]
function Update-GuiText {
    param (
        [Parameter(Mandatory=$true)][System.Windows.Controls.TextBlock]$TextElement,
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$false)][System.Windows.Media.Brush]$Color = $null,
        [Parameter(Mandatory=$false)][int]$MaxLength = 10000
    )
    
    try {
        if ($null -eq $TextElement) { return }
        
        # Text kürzen, wenn zu lang
        if ($Message.Length -gt $MaxLength) {
            $Message = $Message.Substring(0, $MaxLength) + "... (gekürzt)"
        }
        
        # Dispatcher verwenden, um Thread-Sicherheit zu gewährleisten
        $TextElement.Dispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Normal,
            [System.Action]{
                $TextElement.Text = $Message
                if ($null -ne $Color) {
                    $TextElement.Foreground = $Color
                }
            }
        )
    }
    catch {
        Write-Log -Message "Fehler bei Update-GuiText: $($_.Exception.Message)" -Type "Error"
    }
}

function Show-MessageBox {
    param (
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$false)][string]$Title = "Information",
        [Parameter(Mandatory=$false)][System.Windows.MessageBoxButton]$Button = [System.Windows.MessageBoxButton]::OK,
        [Parameter(Mandatory=$false)][System.Windows.MessageBoxImage]$Icon = [System.Windows.MessageBoxImage]::Information
    )
    
    try {
        return [System.Windows.MessageBox]::Show($Message, $Title, $Button, $Icon)
    }
    catch {
        Write-Log -Message "Fehler beim Anzeigen der MessageBox: $($_.Exception.Message)" -Type "Error"
        return [System.Windows.MessageBoxResult]::None
    }
}
#endregion

#region [2.4 | Registry-Funktionen]
function Get-RegistryValue {
    param (
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$false)]$DefaultValue = $null
    )
    
    try {
        if (Test-Path -Path $Path) {
            $value = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $value) {
                return $value.$Name
            }
        }
        return $DefaultValue
    }
    catch {
        Write-Log -Message "Fehler beim Lesen des Registry-Werts '$Name': $($_.Exception.Message)" -Type "Error"
        return $DefaultValue
    }
}

function Set-RegistryValue {
    param (
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)]$Value,
        [Parameter(Mandatory=$false)][string]$Type = "String"
    )
    
    try {
        # Sicherstellen, dass der Registry-Pfad existiert
        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        
        # Wert setzen
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
        return $true
    }
    catch {
        Write-Log -Message "Fehler beim Setzen des Registry-Werts '$Name': $($_.Exception.Message)" -Type "Error"
        return $false
    }
}
#endregion


#region [2.5 | Main GUI Function - Haupt-GUI-Funktion]
function Show-LAPSForm {
    param(
        $ModuleStatus
    )

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms # Für MessageBox
    Add-Type -AssemblyName System.Drawing # Für Farben in MessageBox (optional)

    # (1) Werte aus globalen Hashtables (statt INI)
    $appName        = $script:BrandingGUI.APPName
    if (-not $appName) { $appName = "easyLAPS" }
    $fontFamily     = $script:BrandingGUI.FontFamily
    if (-not $fontFamily) { $fontFamily = "Segoe UI" }
    $headerLogoPath = $script:BrandingGUI.HeaderLogo
    $clickURL       = $script:BrandingGUI.HeaderLogoURL

    $scriptVersion = $script:ScriptInfo.ScriptVersion
    $lastUpdate    = $script:ScriptInfo.LastUpdate
    $author = $script:ScriptInfo.Author
    $footerText = "$appName v$scriptVersion ($lastUpdate) by $author"
    if ($script:BrandingGUI.FooterWebseite) { 
        $webseiteInfo = $script:BrandingGUI.FooterWebseite
        $footerText += " | $webseiteInfo" 
    }

    # (2) XAML Definition für die GUI
    [xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$appName" Height="900" Width="1200" MinHeight="600" MinWidth="800"
        WindowStartupLocation="CenterScreen" Background="#F0F0F0" FontFamily="$fontFamily">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <!-- Header -->
            <RowDefinition Height="*"/>
            <!-- Content -->
            <RowDefinition Height="Auto"/>
            <!-- Footer -->
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#E0E0E0" Padding="10" CornerRadius="5,5,0,0">
            <DockPanel>
                <TextBlock Text="$appName" FontSize="20" FontWeight="Bold" VerticalAlignment="Center" DockPanel.Dock="Left"/>
                <TextBlock Name="lblVersion" Text="v$scriptVersion" FontSize="12" Margin="5,0,0,0" VerticalAlignment="Center" DockPanel.Dock="Left" Foreground="#505050"/>
                <TextBlock Name="lblModuleStatus" Text="" FontSize="12" Margin="5,0,0,0" VerticalAlignment="Center" DockPanel.Dock="Left" Foreground="#505050"/>
                <Image Name="imgLogo" Source="$headerLogoPath" MaxHeight="40" MaxWidth="150" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="10,0,0,0" Cursor="Hand" DockPanel.Dock="Right"/>
            </DockPanel>
        </Border>

        <!-- Content -->
        <Grid Grid.Row="1" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>
                <ColumnDefinition Width="3*"/>
            </Grid.ColumnDefinitions>

            <!-- AD Computer and Actions -->
            <Border Grid.Column="0" Background="White" Padding="10" CornerRadius="5" BorderBrush="#CCCCCC" BorderThickness="1" Grid.ColumnSpan="2" Margin="0,0,505,0">
                <DockPanel>
                    <StackPanel DockPanel.Dock="Top" Orientation="Vertical" Margin="0,0,0,10">
                        <DockPanel Margin="0,0,0,5">
                            <TextBlock Text="Computer name:" VerticalAlignment="Center" Margin="0,0,5,0"/>
                            <TextBox Name="txtComputerName" Width="200" Height="24" VerticalContentAlignment="Center"/>
                            <Button Name="btnRefreshAdList" Content="Load AD List" Margin="10,0,0,0" Padding="5" Height="26" DockPanel.Dock="Right"/>
                        </DockPanel>
                    </StackPanel>

                    <ListBox Name="lstAdComputers" Margin="0,0,0,10" DockPanel.Dock="Top" Height="425" SelectionMode="Single"/>

                    <TextBlock Text="LAPS Password Details:" FontWeight="Bold" Margin="0,10,0,5" DockPanel.Dock="Top"/>
                    <Grid DockPanel.Dock="Top">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Username:" Margin="0,0,5,5" VerticalAlignment="Center"/>
                        <TextBox Name="txtLapsUser" Grid.Row="0" Grid.Column="1" IsReadOnly="True" Height="24" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                        <TextBlock Grid.Row="1" Grid.Column="0" Text="Password:" Margin="0,0,5,5" VerticalAlignment="Center"/>
                        <TextBox Name="txtLapsPasswordDisplay" Grid.Row="1" Grid.Column="1" IsReadOnly="True" Height="24" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Expiration date:" Margin="0,0,5,5" VerticalAlignment="Center"/>
                        <TextBox Name="txtPasswordExpiry" Grid.Row="2" Grid.Column="1" IsReadOnly="True" Height="24" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                        <TextBlock Grid.Row="3" Grid.Column="0" Text="Last change:" Margin="0,0,5,5" VerticalAlignment="Center"/>
                        <TextBox Name="txtPasswordLastSet" Grid.Row="3" Grid.Column="1" IsReadOnly="True" Height="24" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                        <TextBlock Grid.Row="4" Grid.Column="0" Text="LAPS Version:" Margin="0,0,5,5" VerticalAlignment="Center"/>
                        <TextBox Name="txtLapsVersion" Grid.Row="4" Grid.Column="1" IsReadOnly="True" Height="24" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                    </Grid>

                    <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="0,10,0,0" Width="457" Height="42">
                        <Button Name="btnGetLapsPassword" Content="Show Password" Width="120" Height="26" Margin="0,0,10,0" ToolTip="Shows the LAPS password (Ctrl+G)"/>
                        <Button Name="btnResetLapsPassword" Content="Reset Password" Width="140" Height="26" Margin="0,0,10,0" ToolTip="Resets the LAPS password (Ctrl+R)"/>
                        <Button Name="btnCopyPassword" Content="Copy to Clipboard" Width="120" Height="26" ToolTip="Copies the password to clipboard (Ctrl+C)"/>
                    </StackPanel>
                </DockPanel>
            </Border>

            <!-- LAPS Details and Status -->
            <Border Grid.Column="1" Background="White" Padding="10" CornerRadius="5" Margin="208,0,0,0" BorderBrush="#CCCCCC" BorderThickness="1">
                <DockPanel>
                    <TextBlock Text="LAPS Client Status and Policies" DockPanel.Dock="Top" FontSize="14" FontWeight="Bold" Margin="0,0,0,10"></TextBlock>
                    <Button Name="btnCheckLapsStatus" Content="Check Status" Width="120" Height="26" DockPanel.Dock="Top" HorizontalAlignment="Left" Margin="0,0,0,10" ToolTip="Checks the LAPS status (Ctrl+P)"/>

                    <Grid DockPanel.Dock="Top">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Grid.Row="0" Grid.Column="0" Text="LAPS installed:" VerticalAlignment="Center" Margin="0,0,5,5"/>
                        <Ellipse Name="ellipseLapsInstalled" Grid.Row="0" Grid.Column="1" Width="16" Height="16" Fill="Gray" Stroke="Black" StrokeThickness="0.5" VerticalAlignment="Center" Margin="0,0,5,5"/>
                        <TextBlock Name="lblLapsInstalledStatus" Grid.Row="0" Grid.Column="2" Text="Not checked" VerticalAlignment="Center" Margin="0,0,0,5"/>

                        <TextBlock Grid.Row="1" Grid.Column="0" Text="LAPS GPO active:" VerticalAlignment="Center" Margin="0,0,5,5"/>
                        <Ellipse Name="ellipseGpoStatus" Grid.Row="1" Grid.Column="1" Width="16" Height="16" Fill="Gray" Stroke="Black" StrokeThickness="0.5" VerticalAlignment="Center" Margin="0,0,5,5"/>
                        <TextBlock Name="lblGpoStatus" Grid.Row="1" Grid.Column="2" Text="Not checked" VerticalAlignment="Center" Margin="0,0,0,5"/>

                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Policy details:" VerticalAlignment="Top" Margin="0,5,5,0"/>
                        <TextBox Name="txtPolicyDetails" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" IsReadOnly="True" Height="200" Margin="0,5,0,0" FontFamily="Consolas" FontSize="11"/>
                    </Grid>
                    <TextBlock Name="txtStatus" Text="Ready." DockPanel.Dock="Bottom" Margin="0,10,0,0" FontStyle="Italic" FontSize="10"/>
                </DockPanel>
            </Border>
        </Grid>

        <!-- Footer -->
        <Border Grid.Row="2" Background="#E0E0E0" Padding="5" CornerRadius="0,0,5,5" Margin="0,10,0,0">
            <TextBlock Name="lblFooter" Text="$footerText" FontSize="10" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
    </Grid>
</Window>
"@

    # (3) XAML laden und GUI-Elemente referenzieren
    try {
        $reader = [System.IO.StringReader]::new($XAML)
        $xmlReader = [System.Xml.XmlReader]::Create($reader)
        $window = [Windows.Markup.XamlReader]::Load($xmlReader)
    }
    catch {
        $errorMsg = "Fehler beim Laden der Benutzeroberfläche: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            $errorMsg += "`nInner Exception: $($_.Exception.InnerException.Message)"
        }
        [System.Windows.MessageBox]::Show($errorMsg, "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $null
    }

    # GUI-Elemente (Variablen für einfacheren Zugriff)
    $controls = @{}
    @("txtComputerName", "btnRefreshAdList", "lstAdComputers", "txtFilter", "txtLapsUser", "txtLapsPasswordDisplay", "txtPasswordExpiry",
      "txtPasswordLastSet", "txtLapsVersion", "btnGetLapsPassword", "btnResetLapsPassword", "btnCopyPassword", "btnCheckLapsStatus", 
      "ellipseLapsInstalled", "lblLapsInstalledStatus", "ellipseGpoStatus", "lblGpoStatus", "txtPolicyDetails", 
      "txtStatus", "imgLogo", "lblFooter", "lblVersion", "lblModuleStatus") | ForEach-Object {
        $controls[$_] = $window.FindName($_)
    }

    # (4) Event Handler und Logik

    # Tastenkombinationen definieren
    $window.Add_KeyDown({
        param($sender, $e)
        
        # Strg+P für Status & Richtlinien prüfen
        if ($e.Key -eq 'P' -and $e.KeyboardDevice.Modifiers -eq 'Control') {
            $controls.btnCheckLapsStatus.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            $e.Handled = $true
        }        
        # Strg+F für Fokus auf Filter
        if ($e.Key -eq 'F' -and $e.KeyboardDevice.Modifiers -eq 'Control') {
            $controls.txtFilter.Focus()
            $controls.txtFilter.SelectAll()
            $e.Handled = $true
        }
        
        # F5 für AD-Liste aktualisieren
        if ($e.Key -eq 'F5') {
            $controls.btnRefreshAdList.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            $e.Handled = $true
        }
        
        # Enter im Computer-Textfeld löst Passwortanzeige aus
        if ($e.Key -eq 'Return' -and $e.Source -eq $controls.txtComputerName) {
            $controls.btnGetLapsPassword.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            $e.Handled = $true
        }
        
        # Strg+G für Passwort anzeigen
        if ($e.Key -eq 'G' -and $e.KeyboardDevice.Modifiers -eq 'Control') {
            $controls.btnGetLapsPassword.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            $e.Handled = $true
        }
        
        # Strg+R für Passwort zurücksetzen
        if ($e.Key -eq 'R' -and $e.KeyboardDevice.Modifiers -eq 'Control') {
            $controls.btnResetLapsPassword.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            $e.Handled = $true
        }
        
        # Strg+C für Passwort kopieren (nur wenn kein Text markiert ist)
        if ($e.Key -eq 'C' -and $e.KeyboardDevice.Modifiers -eq 'Control') {
            $controls.btnCopyPassword.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            $e.Handled = $true
        }
    })
    
    # Logo Klick (falls URL in INI definiert)
    if ($controls.imgLogo -and $clickURL) {
        $controls.imgLogo.Add_MouseUp({
            param($sender, $e)
            if ($e.ChangedButton -eq 'Left') { Open-URLInBrowser $clickURL }
        })
    }
    if (-not (Test-Path $headerLogoPath)) {
        $controls.imgLogo.Visibility = 'Collapsed'
    }

    #region [4.0 | Statusanzeige und Hilfsfunktionen]

# Globale Funktion für Status-Updates (Session-Historie)
$global:UpdateStatus = {
    param(
        [string]$Message,
        [string]$Level = "Info",
        [switch]$IsLapsCheck
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $hostname = $env:COMPUTERNAME
    $prefix = "[$timestamp]"
    if ($IsLapsCheck) {
        $prefix += " [LAPS-Check $hostname]"
    }
    $entry = "$prefix [$Level] $Message"
    if ($controls -and $controls.txtStatus) {
        if ([string]::IsNullOrWhiteSpace($controls.txtStatus.Text)) {
            $controls.txtStatus.Text = $entry
        } else {
            $controls.txtStatus.Text += "`r`n$entry"
        }
        # Farblogik
        switch ($Level.ToLower()) {
            "pruefung" { $controls.txtStatus.Foreground = [System.Windows.Media.Brushes]::DarkBlue }
            "check"    { $controls.txtStatus.Foreground = [System.Windows.Media.Brushes]::DarkBlue }
            "ok"       { $controls.txtStatus.Foreground = [System.Windows.Media.Brushes]::DarkGreen }
            "success"  { $controls.txtStatus.Foreground = [System.Windows.Media.Brushes]::DarkGreen }
            "debug"    { $controls.txtStatus.Foreground = [System.Windows.Media.Brushes]::LightBlue }
            "fehler"   { $controls.txtStatus.Foreground = [System.Windows.Media.Brushes]::DarkRed }
            "error"    { $controls.txtStatus.Foreground = [System.Windows.Media.Brushes]::DarkRed }
            default     { $controls.txtStatus.Foreground = [System.Windows.Media.Brushes]::Black }
        }
    }
}
#endregion

    # Global variables for computer filtering
    $script:allComputers = @()
    
    # Function to filter computers based on search text
    $FilterComputers = {
        param($filterText)
        
        if ([string]::IsNullOrWhiteSpace($filterText)) {
            return $script:allComputers
        }
        
        return $script:allComputers | Where-Object { $_ -like "*$filterText*" }
    }
    
    # AD Computerliste laden
    $controls.btnRefreshAdList.Add_Click({
        & $global:UpdateStatus -IsLapsCheck "Lade AD Computerliste..."
        try {
            $script:allComputers = Get-ADComputer -Filter * -Properties Name | Select-Object -ExpandProperty Name | Sort-Object
            $controls.lstAdComputers.ItemsSource = $script:allComputers
            if ($script:allComputers) {
                & $global:UpdateStatus -IsLapsCheck "AD Computerliste geladen ($($script:allComputers.Count) Eintraege)."
            } else {
                & $global:UpdateStatus -IsLapsCheck "Keine Computer in AD gefunden."
            }
        }
        catch {
            & $global:UpdateStatus -IsLapsCheck "Fehler beim Laden der AD Computerliste: $($_.Exception.Message)" "Error"
        }
        
        $filterText = $controls.txtFilter.Text
        # Fix syntax error by explicitly passing the parameter
        $filteredComputers = Invoke-Command -ScriptBlock $FilterComputers -ArgumentList $filterText
        $controls.lstAdComputers.ItemsSource = $filteredComputers
        
        if ([string]::IsNullOrWhiteSpace($filterText)) {
            & $global:UpdateStatus -IsLapsCheck "Filter zurueckgesetzt. Zeige alle $($script:allComputers.Count) Computer."
        } else {
            & $global:UpdateStatus -IsLapsCheck "Filter angewendet - '$filterText'. $($filteredComputers.Count) von $($script:allComputers.Count) Computern werden angezeigt."
        }
    })

    # Register event handler for list selection
    $controls.lstAdComputers.Add_SelectionChanged({
        param($sender, $e)
        
        $selectedComputer = $controls.lstAdComputers.SelectedItem
        if (-not $selectedComputer) { return }
        
        $controls.txtComputerName.Text = $selectedComputer
        
        # Clear previous values
        $controls.txtLapsUser.Text = "Lade..."
        $controls.txtLapsPasswordDisplay.Text = "Lade..."
        $controls.txtPasswordExpiry.Text = "Lade..."
        $controls.txtPasswordLastSet.Text = "Lade..."
        $controls.txtLapsVersion.Text = "Lade..."
        
        try {
            # Check if LAPS is installed
            $lapsInstalled = $false
            $lapsType = ""
            $winLaps = Get-WmiObject -Class Win32_OptionalFeature -Filter "Name='LAPS.Windows' AND InstallState=1" -ErrorAction SilentlyContinue
            if ($winLaps) {
                $lapsInstalled = $true
                $lapsType = "Windows LAPS"
            } elseif (Test-Path "C:\Program Files\LAPS\CSE\Admpwd.dll") {
                $lapsInstalled = $true
                $lapsType = "Legacy LAPS (Admpwd.dll)"
            }
            
            if (-not $lapsInstalled) {
                $controls.txtLapsUser.Text = "N/A"
                $controls.txtLapsPasswordDisplay.Text = "LAPS ist nicht installiert."
                $controls.txtPasswordExpiry.Text = "N/A"
                $controls.txtPasswordLastSet.Text = "N/A"
                $controls.txtLapsVersion.Text = "N/A"
                & $global:UpdateStatus -IsLapsCheck "LAPS ist auf diesem System nicht installiert." "Warning"
                return
            }

            # Check if Get-LapsPassword cmdlet is available
            if (-not (Get-Command -Name Get-LapsPassword -ErrorAction SilentlyContinue)) {
                $controls.txtLapsUser.Text = "N/A"
                $controls.txtLapsPasswordDisplay.Text = "Das Cmdlet 'Get-LapsPassword' ist nicht verfuegbar."
                $controls.txtPasswordExpiry.Text = "N/A"
                $controls.txtPasswordLastSet.Text = "N/A"
                $controls.txtLapsVersion.Text = "N/A"
                & $global:UpdateStatus -IsLapsCheck "Das Cmdlet 'Get-LapsPassword' ist nicht verfuegbar. Pruefe LAPS-Installation/Modul." "Warning"
                return
            }
            
            & $global:UpdateStatus -IsLapsCheck "Lese LAPS Passwort fuer $selectedComputer..."
            
            try {
                # Use the improved function for LAPS password
                $passwordInfo = Get-LapsPassword -ComputerName $selectedComputer
                
                if (-not $passwordInfo.Error) {
                    $plainPassword = Convert-SecureStringToPlainText -SecureString $passwordInfo.Password
                    $controls.txtLapsUser.Text = $passwordInfo.UserName
                    $controls.txtLapsPasswordDisplay.Text = $plainPassword
                    $controls.txtPasswordExpiry.Text = if ($passwordInfo.PasswordExpiresTimestamp) { $passwordInfo.PasswordExpiresTimestamp.ToString("g") } else { "N/A" }
                    $controls.txtPasswordLastSet.Text = if ($passwordInfo.PasswordLastSetTimestamp) { $passwordInfo.PasswordLastSetTimestamp.ToString("g") } else { "N/A" }
                    $controls.txtLapsVersion.Text = $passwordInfo.LapsVersion
                    & $global:UpdateStatus -IsLapsCheck "LAPS Passwort fuer $selectedComputer erfolgreich gelesen."
                } else {
                    $controls.txtLapsUser.Text = "N/A"
                    $controls.txtLapsPasswordDisplay.Text = "Passwort nicht gefunden oder nicht verfuegbar."
                    $controls.txtPasswordExpiry.Text = "N/A"
                    $controls.txtPasswordLastSet.Text = "N/A"
                    $controls.txtLapsVersion.Text = "N/A"
                    & $global:UpdateStatus -IsLapsCheck "Konnte LAPS-Passwort fuer $selectedComputer nicht abrufen." "Warning"
                }
            } catch {
                $controls.txtLapsUser.Text = "N/A"
                $controls.txtLapsPasswordDisplay.Text = "Fehler beim Lesen des LAPS-Passworts."
                $controls.txtPasswordExpiry.Text = "N/A"
                $controls.txtPasswordLastSet.Text = "N/A"
                $controls.txtLapsVersion.Text = "N/A"
                & $global:UpdateStatus -IsLapsCheck "Fehler beim Lesen des LAPS-Passworts: $($_.Exception.Message)" "Error"
            }
        } catch {
            $controls.txtLapsUser.Text = "N/A"
            $controls.txtLapsPasswordDisplay.Text = "LAPS Passwort konnte nicht gelesen werden."
            $controls.txtPasswordExpiry.Text = "N/A"
            $controls.txtPasswordLastSet.Text = "N/A"
            $controls.txtLapsVersion.Text = "N/A"
            & $global:UpdateStatus -IsLapsCheck "Fehler beim Laden der AD-Computer: $($_.Exception.Message)" "Error"
        }
    })

    # LAPS Passwort zuruecksetzen
    $controls.btnResetLapsPassword.Add_Click({
        $computer = $controls.txtComputerName.Text
        if ([string]::IsNullOrWhiteSpace($computer)) {
            & $global:UpdateStatus -IsLapsCheck "Bitte einen Computernamen eingeben oder auswaehlen, um den Status zu pruefen." "Error"
            [System.Windows.MessageBox]::Show("Bitte einen Computernamen eingeben oder auswaehlen.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }
        $confirm = [System.Windows.MessageBox]::Show(
            "Moechten Sie das LAPS Passwort fuer '$computer' wirklich jetzt zuruecksetzen?", 
            "Bestaetigung", 
            [System.Windows.MessageBoxButton]::YesNo, 
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($confirm -eq 'Yes') {
            & $global:UpdateStatus -IsLapsCheck "Setze LAPS Passwort fuer $computer zurueck..."
            try {
                # Use the improved function for resetting LAPS passwords
                $resetResult = Reset-LapsPasswordEx -ComputerName $computer -Force
                
                if ($resetResult.Success) {
                    & $global:UpdateStatus -IsLapsCheck "LAPS Passwort fuer $computer erfolgreich zurueckgesetzt. Es kann einige Zeit dauern, bis es aktiv wird."
                    [System.Windows.MessageBox]::Show(
                        "LAPS Passwort fuer '$computer' erfolgreich zurueckgesetzt. LAPS Version: $($resetResult.LapsVersion)", 
                        "Erfolg", 
                        [System.Windows.MessageBoxButton]::OK, 
                        [System.Windows.MessageBoxImage]::Information
                    )
                    # Clear password fields since the new password cannot be read immediately
                    $controls.txtLapsUser.Text = ""
                    $controls.txtLapsPasswordDisplay.Text = "(Passwort wurde zurueckgesetzt)"
                    $controls.txtPasswordExpiry.Text = ""
                    $controls.txtPasswordLastSet.Text = $(Get-Date -Format "g")
                    $controls.txtLapsVersion.Text = $resetResult.LapsVersion
                } else {
                    & $global:UpdateStatus -IsLapsCheck "Fehler beim Zuruecksetzen des LAPS-Passworts: $($resetResult.ErrorMessage)" "Error"
                    [System.Windows.MessageBox]::Show("Fehler beim Zuruecksetzen des LAPS-Passworts: $($resetResult.ErrorMessage)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            }
            catch {
                & $global:UpdateStatus -IsLapsCheck "Fehler beim Zuruecksetzen des LAPS-Passworts: $($_.Exception.Message)" "Error"
                [System.Windows.MessageBox]::Show("Fehler beim Zuruecksetzen des LAPS-Passworts: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })

    # Passwort kopieren
    $controls.btnCopyPassword.Add_Click({
        $password = $controls.txtLapsPasswordDisplay.Text
        if (-not [string]::IsNullOrWhiteSpace($password) -and $password -notmatch "Fehler|gefunden|zurueckgesetzt") {
            Set-Clipboard -Text $password
            & $global:UpdateStatus -IsLapsCheck "Passwort in die Zwischenablage kopiert."
        } else {
            & $global:UpdateStatus -IsLapsCheck "Kein Passwort zum Kopieren vorhanden." "Warning"
        }
    })

    # LAPS Client Status & Richtlinien pruefen
    $controls.btnCheckLapsStatus.Add_Click({
        $computer = $env:COMPUTERNAME
        & $global:UpdateStatus -IsLapsCheck "Pruefe LAPS-Status fuer $computer..." "Pruefung"
        
        try {
            # Pruefe, ob LAPS installiert ist
            $lapsInstalled = $false
            $lapsType = ""
            $winLaps = Get-WmiObject -Class Win32_OptionalFeature -Filter "Name='LAPS.Windows' AND InstallState=1" -ErrorAction SilentlyContinue
            if ($winLaps) {
                $lapsInstalled = $true
                $lapsType = "Windows LAPS"
            } elseif (Test-Path "C:\Program Files\LAPS\CSE\Admpwd.dll") {
                $lapsInstalled = $true
                $lapsType = "Legacy LAPS (Admpwd.dll)"
            }
            
            # Aktualisiere die UI basierend auf dem LAPS-Status
            if ($lapsInstalled) {
                $controls.lblLapsInstalledStatus.Text = "$lapsType installiert"
                $controls.ellipseLapsInstalled.Fill = [System.Windows.Media.Brushes]::Green
                & $global:UpdateStatus -IsLapsCheck "$lapsType auf diesem System gefunden. Pruefe GPO..." "Pruefung"
                
                # Pruefe GPO-Einstellungen
                $os = Get-CimInstance -ClassName Win32_OperatingSystem
                $isServer = $os.ProductType -ne 1
                $gpoKey = if ($isServer) { "HKLM:\SOFTWARE\Policies\Microsoft Services\AdmPwd" } else { "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdmPwd" }
                $gpoActive = Test-Path $gpoKey
                
                if ($gpoActive) {
                    $controls.lblGpoStatus.Text = "LAPS GPO aktiv"
                    $controls.ellipseGpoStatus.Fill = [System.Windows.Media.Brushes]::Green
                    $controls.txtPolicyDetails.Text = "LAPS ist korrekt konfiguriert.`nGPO-Pfad: $gpoKey"
                    & $global:UpdateStatus -IsLapsCheck "LAPS GPO auf diesem System gefunden und aktiv." "OK"
                } else {
                    $controls.lblGpoStatus.Text = "Keine LAPS GPO gefunden"
                    $controls.ellipseGpoStatus.Fill = [System.Windows.Media.Brushes]::Red
                    $controls.txtPolicyDetails.Text = "Keine passende LAPS GPO gefunden.`nBitte konfigurieren Sie die LAPS-Gruppenrichtlinie."
                    & $global:UpdateStatus -IsLapsCheck "Keine passende LAPS GPO gefunden. Bitte konfigurieren Sie die LAPS-Gruppenrichtlinie." "Warning"
                }
            } else {
                $controls.lblLapsInstalledStatus.Text = "Nicht installiert"
                $controls.ellipseLapsInstalled.Fill = [System.Windows.Media.Brushes]::Red
                $controls.lblGpoStatus.Text = "Nicht verfügbar"
                $controls.ellipseGpoStatus.Fill = [System.Windows.Media.Brushes]::Gray
                $controls.txtPolicyDetails.Text = "LAPS ist auf diesem System nicht installiert.`nBitte installieren Sie LAPS, um fortzufahren."
                & $global:UpdateStatus -IsLapsCheck "LAPS ist auf diesem System nicht installiert." "Warning"
            }
        } catch {
            $controls.lblLapsInstalledStatus.Text = "Fehler"
            $controls.ellipseLapsInstalled.Fill = [System.Windows.Media.Brushes]::Red
            $controls.lblGpoStatus.Text = "Fehler"
            $controls.ellipseGpoStatus.Fill = [System.Windows.Media.Brushes]::Red
            $controls.txtPolicyDetails.Text = "Fehler beim Überprüfen des LAPS-Status: $($_.Exception.Message)"
            & $global:UpdateStatus -IsLapsCheck "Fehler beim Überprüfen des LAPS-Status: $($_.Exception.Message)" "Error"
        }
    })

    # LAPS Status Check Handler
    $controls.btnCheckLapsStatus.Add_Click({
        & $global:UpdateStatus -IsLapsCheck "Pruefe LAPS-Status des lokalen Systems..." "Pruefung"
    })
    
    # Initiale Prüfung beim Programmstart ausführen
    & $global:UpdateStatus -IsLapsCheck "Pruefe LAPS-Status des lokalen Systems..." "Pruefung"
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $isServer = $os.ProductType -ne 1
        $systemType = if ($isServer) { "Server" } else { "Client" }
        $controls.txtPolicyDetails.Text = "Systemtyp erkannt: $systemType"
        
        # Initialisierung der Variablen
        $lapsInstalled = $false
        $lapsType = ""
        $winLaps = Get-WmiObject -Class Win32_OptionalFeature -Filter "Name='LAPS.Windows' AND InstallState=1" -ErrorAction SilentlyContinue
        if ($winLaps) {
            $lapsInstalled = $true
            $lapsType = "Windows LAPS"
        } elseif (Test-Path "C:\Program Files\LAPS\CSE\Admpwd.dll") {
            $lapsInstalled = $true
            $lapsType = "Legacy LAPS (Admpwd.dll)"
        }
        if ($lapsInstalled) {
            $controls.lblLapsInstalledStatus.Text = "$lapsType installiert"
            $controls.ellipseLapsInstalled.Fill = [System.Windows.Media.Brushes]::Green
            & $global:UpdateStatus -IsLapsCheck "$lapsType auf diesem System gefunden. Pruefe GPO..." "Pruefung"
            $gpoKey = if ($isServer) { "HKLM:\SOFTWARE\Policies\Microsoft Services\AdmPwd" } else { "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdmPwd" }
            $gpoActive = Test-Path $gpoKey
            if ($gpoActive) {
                $controls.lblGpoStatus.Text = "LAPS GPO aktiv"
                $controls.ellipseGpoStatus.Fill = [System.Windows.Media.Brushes]::Green
                $controls.txtPolicyDetails.Text += "`nLAPS GPO wurde gefunden und ist aktiv ($gpoKey)."
                & $global:UpdateStatus -IsLapsCheck "LAPS GPO auf diesem System gefunden." "OK"
            } else {
                $controls.lblGpoStatus.Text = "Keine LAPS GPO gefunden"
                $controls.ellipseGpoStatus.Fill = [System.Windows.Media.Brushes]::Red
                $controls.txtPolicyDetails.Text += "`nKeine passende LAPS GPO gefunden. Weitere Infos auf der Microsoft-Webseite."
                & $global:UpdateStatus -IsLapsCheck "Keine passende LAPS GPO gefunden. Oeffne Microsoft-Webseite..." "Fehler"
                Start-Process "https://learn.microsoft.com/de-de/windows-server/identity/laps/laps-overview"
            }
        } else {
            $controls.lblLapsInstalledStatus.Text = "Nicht installiert"
            $controls.ellipseLapsInstalled.Fill = [System.Windows.Media.Brushes]::Red
            $controls.lblGpoStatus.Text = "Unbekannt"
            $controls.ellipseGpoStatus.Fill = [System.Windows.Media.Brushes]::Gray
            $controls.txtPolicyDetails.Text += "`nLAPS ist auf diesem Computer nicht installiert."
            & $global:UpdateStatus -IsLapsCheck "LAPS ist auf diesem System nicht installiert." "Fehler"
        }
    }
    catch {
        $controls.lblLapsInstalledStatus.Text = "Fehler"
        $controls.ellipseLapsInstalled.Fill = [System.Windows.Media.Brushes]::Red
        $controls.lblGpoStatus.Text = "Fehler"
        $controls.ellipseGpoStatus.Fill = [System.Windows.Media.Brushes]::Red
        $controls.txtPolicyDetails.Text = "Fehler"
        & $global:UpdateStatus -IsLapsCheck "Ausnahme beim Pruefen des LAPS-Status: $($_.Exception.Message)" "Fehler"
    }

    # Update module status label
    if ($ModuleStatus) {
        $moduleStatusText = "Module Status: $($ModuleStatus.LegacyModuleAvailable), $($ModuleStatus.NewModuleAvailable)"
        $controls.lblModuleStatus.Text = $moduleStatusText
    }

    # Fenster anzeigen
    $window.ShowDialog() | Out-Null
}

#region Main Script Execution - Hauptskriptausführung
# Set error action preference for the main script block
$ErrorActionPreference = "Stop"

# Initialize Global Log File Path
$global:easyLAPSLogFile = Join-Path -Path $PSScriptRoot -ChildPath "easyLAPS_$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

try {
    # Lade erforderliche Assemblies für WPF-Komponenten
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    
    Write-Log "easyLAPS Script gestartet. Logging in: $global:easyLAPSLogFile" -Level "INFO"
    Write-Log "easyLAPS GUI wird gestartet..." -Level "INFO"

    # Hauptformular anzeigen
    Show-LAPSForm

    Write-Log "easyLAPS GUI wurde vom Benutzer oder einem internen Prozess geschlossen." -Level "INFO"
}
catch {
    $criticalErrorMsg = "Ein kritischer Fehler ist im Hauptskriptblock aufgetreten: $($_.Exception.Message)"
    Write-Log $criticalErrorMsg -Level "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    try { 
        # Versuche, die erforderlichen Assemblies zu laden, falls noch nicht geschehen
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show("$criticalErrorMsg`n`nDetails:`n$($_.ScriptStackTrace)", "Kritischer Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } 
    catch {
        # Fallback auf einfache Konsolenausgabe, wenn MessageBox nicht verfuegbar ist
        Write-Host "KRITISCHER FEHLER: $criticalErrorMsg`n`nDetails:`n$($_.ScriptStackTrace)" -ForegroundColor Red
    }
    Write-Log "easyLAPS Script wegen kritischem Fehler beendet." -Level "ERROR"
    exit 1
}
finally {
    Write-Log "easyLAPS Script-Ausfuehrung beendet." -Level "INFO"
}
