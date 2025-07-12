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
    $XAML = @"
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
        # XAML-Text in XML-Objekt konvertieren
        [xml]$xamlObject = $XAML
        # Namespace-Manager für XPath-Abfragen erstellen
        $namespace = New-Object System.Xml.XmlNamespaceManager($xamlObject.NameTable)
        $namespace.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")
        
        # XAML in Objekt umwandeln
        $reader = New-Object System.Xml.XmlNodeReader($xamlObject)
        $window = [Windows.Markup.XamlReader]::Load($reader)
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

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCACdRvHWf+Pmy2n
# 4Ujx3Sbd96posUxKI81MyRX1cwOySqCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCBBUUPCJbNEUsrKlkZfdwIW9j5WTLdIoPmF+W4UvKhxeTANBgkqhkiG
# 9w0BAQEFAASCAQALRSsCDCuNEO2WlIif6MMoB23I9f1x4V6sGsxAa6RnPT86aJAk
# tGKdxMsYTbsJ+XFPGviKXBrP5tJx4+cNLHnIdII1cyfy6KaiGxIbFTJPBygy0KrL
# EGgZ44fBFGoGxS+qogU1WgyHHZ9LXqBChKWsPeo0ykM6fJyNmpveE3HPtYQ94GzB
# 09Tkcg0R6jmsEck++auXFy5r/lEJv1ssweMy8IjXsaZkzHrtRHLUoUcgvKR9P63/
# gtI4YO8ujaKq0AAb8Pn4NPF6CuCXyz0m05cLHPD2sz/tYcIB2aZpNzSm3YaZtFUe
# +NPxZWfp/MVJqyV40BiRNIU5+E6vxmj/HxqKoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTE1MFowLwYJKoZIhvcNAQkEMSIEIMMV/YhwmaboJUKNdxp7HBqk9C8n
# g74s6WhpOCd3wgK9MA0GCSqGSIb3DQEBAQUABIICAFvc1t/oEKNRdh1dE8nwjKfg
# oF7DtN6p8//B1p0cbpXhcLoEzDZd5AKnP2pGEhzqvLOVY6VHx25FWDCJqkfzhpJD
# O/ZxJg4ZZEPqxORkNbjkv94ejWccW3de857qFZV7Z0gHtZZeF3dgy7E/eUwPMEVK
# FKT2mXFnhN7fTOTxIFqI1gK71PmFS7tLOBgt69rEtbeZcaDbS7tfmswvL+eUA7No
# v1J9niBTCqDF87xeEAtzrtN7/p1NT960h6eKvTxnyIU8g5Hys6qMGsrUKUOQ6p01
# SbjIOsG5Mg1SaBjC4XsPDKRjHqJMxAnzaHXg7gfQQNSy0UPcr4e8czug+H5if+VJ
# dnFzD4qggtfTkdOL72ppkkP/05Sh/fnmyK+swfBi+fhvZCumuT/O6zPNXuHTMpTE
# lIHDH3HUuNscE5cL2jDasX6kmhDD8YGV4zFZM94cJzJCcTPQECsvOSdKU7fOcEmW
# DV9T78dIFuu8Vh7jAup2bcxjuUKkN2uxEyMaGNYDptv1joRDYhCKKUI1ZWoOooTc
# FmBlSg3aYIN9Qpt7Z6U4mGTMYZTGe9OklMaxx0gYQ3vvMx5RaK024+UFm6Wml/hI
# 888D+btLhLxfVsAyHJW10gC9+2GJG0FmhoYzebF09nIErAlKRmQxt9yFNgEHQf5Z
# ThDrq9pwCSIzDItxuQ5q
# SIG # End signature block
