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
        Title="$appName" Height="940" Width="1200" MinHeight="600" MinWidth="800"
        WindowStartupLocation="CenterScreen" Background="#FFFFFF" FontFamily="$fontFamily" FontSize="12">
    <Window.Resources>
        <!-- Modern Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="10,5" />
            <Setter Property="Margin" Value="5" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="MinWidth" Value="80" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border
                            x:Name="border"
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="4">
                            <ContentPresenter
                                Margin="{TemplateBinding Padding}"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#106EBE" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#005A9E" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Card Style -->
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="White" />
            <Setter Property="BorderBrush" Value="#E0E0E0" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="CornerRadius" Value="8" />
            <Setter Property="Padding" Value="15" />
            <Setter Property="Margin" Value="5" />
        </Style>

        <!-- TextBox Style -->
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="BorderBrush" Value="#BFBFBF" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Padding" Value="8,5" />
            <Setter Property="Margin" Value="4" />
            <Setter Property="Background" Value="White" />
            <Setter Property="FontSize" Value="12" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#0078D4" />
                                <Setter Property="BorderThickness" Value="2" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5" />
                                <Setter Property="Background" Value="#F0F0F0" />
                            </Trigger>
                            <Trigger Property="IsReadOnly" Value="True">
                                <Setter Property="Background" Value="#F5F5F5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
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
        <Border Grid.Row="0" Background="#FF1C323C" Padding="10" CornerRadius="5,5,0,0" BorderBrush="#E0E0E0" BorderThickness="0,0,0,1">
            <DockPanel>
                <TextBlock Text="$appName" FontSize="20" FontWeight="SemiBold" VerticalAlignment="Center" DockPanel.Dock="Left" Foreground="#e2e2e2"/>
                <TextBlock Name="lblVersion" Text="v$scriptVersion" FontSize="12" Margin="10,0,0,0" VerticalAlignment="Center" DockPanel.Dock="Left" Foreground="#d0e8ff"/>
                <TextBlock Name="lblModuleStatus" Text="" FontSize="12" Margin="10,0,0,0" VerticalAlignment="Center" DockPanel.Dock="Left" Foreground="#d0e8ff"/>
                <Image Name="imgLogo" Source="$headerLogoPath" MaxHeight="35" MaxWidth="140" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="10,0,0,0" Cursor="Hand" DockPanel.Dock="Right"/>
            </DockPanel>
        </Border>

        <!-- Content -->
        <Grid Grid.Row="1" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>
                <ColumnDefinition Width="3*"/>
            </Grid.ColumnDefinitions>

            <!-- AD Computer and Actions -->
            <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,5,0">
                <DockPanel>
                    <StackPanel DockPanel.Dock="Top" Orientation="Vertical" Margin="0,0,0,10">
                        <DockPanel Margin="0,0,0,5">
                            <TextBlock Text="Computer name:" VerticalAlignment="Center" Margin="0,0,5,0" Foreground="#1C1C1C"/>
                            <TextBox Name="txtComputerName" Width="200" Height="30" VerticalContentAlignment="Center"/>
                            <Button Name="btnRefreshAdList" Content="Load AD List" Margin="10,0,0,0" Padding="10,5" Height="30" DockPanel.Dock="Right" Style="{StaticResource ModernButton}"/>
                        </DockPanel>
                    </StackPanel>

                    <ListBox Name="lstAdComputers" Margin="0,0,0,10" DockPanel.Dock="Top" Height="427" SelectionMode="Single" BorderBrush="#BFBFBF" BorderThickness="1"/>

                    <TextBlock Text="LAPS Password Details:" FontWeight="SemiBold" Margin="0,10,0,5" DockPanel.Dock="Top" Foreground="#1C1C1C"/>
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
                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Username:" Margin="0,0,5,5" VerticalAlignment="Center" Foreground="#505050"/>
                        <TextBox Name="txtLapsUser" Grid.Row="0" Grid.Column="1" IsReadOnly="True" Height="30" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                        <TextBlock Grid.Row="1" Grid.Column="0" Text="Password:" Margin="0,0,5,5" VerticalAlignment="Center" Foreground="#505050"/>
                        <TextBox Name="txtLapsPasswordDisplay" Grid.Row="1" Grid.Column="1" IsReadOnly="True" Height="30" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Expiration date:" Margin="0,0,5,5" VerticalAlignment="Center" Foreground="#505050"/>
                        <TextBox Name="txtPasswordExpiry" Grid.Row="2" Grid.Column="1" IsReadOnly="True" Height="30" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                        <TextBlock Grid.Row="3" Grid.Column="0" Text="Last change:" Margin="0,0,5,5" VerticalAlignment="Center" Foreground="#505050"/>
                        <TextBox Name="txtPasswordLastSet" Grid.Row="3" Grid.Column="1" IsReadOnly="True" Height="30" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                        <TextBlock Grid.Row="4" Grid.Column="0" Text="LAPS Version:" Margin="0,0,5,5" VerticalAlignment="Center" Foreground="#505050"/>
                        <TextBox Name="txtLapsVersion" Grid.Row="4" Grid.Column="1" IsReadOnly="True" Height="30" VerticalContentAlignment="Center" Margin="0,0,0,5"/>
                    </Grid>

                    <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="0,10,0,0" Height="45">
                        <Button Name="btnGetLapsPassword" Content="Show Password" Width="130" Height="30" Margin="0,0,10,0" ToolTip="Shows the LAPS password (Ctrl+G)" Style="{StaticResource ModernButton}"/>
                        <Button Name="btnResetLapsPassword" Content="Reset Password" Width="130" Height="30" Margin="0,0,10,0" ToolTip="Resets the LAPS password (Ctrl+R)" Style="{StaticResource ModernButton}"/>
                        <Button Name="btnCopyPassword" Content="Copy to Clipboard" Width="130" Height="30" ToolTip="Copies the password to clipboard (Ctrl+C)" Style="{StaticResource ModernButton}"/>
                    </StackPanel>
                </DockPanel>
            </Border>

            <!-- LAPS Details and Status -->
            <Border Grid.Column="1" Style="{StaticResource Card}" Margin="5,0,0,0">
                <DockPanel>
                    <TextBlock Text="LAPS Client Status and Policies" DockPanel.Dock="Top" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10" Foreground="#1C1C1C"/>
                    <Button Name="btnCheckLapsStatus" Content="Check Status" Width="130" Height="30" DockPanel.Dock="Top" HorizontalAlignment="Left" Margin="75,0,0,10" ToolTip="Checks the LAPS status (Ctrl+P)" Style="{StaticResource ModernButton}"/>

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

                        <TextBlock Grid.Row="0" Grid.Column="1" Text="LAPS installed:" VerticalAlignment="Top" Margin="1,1,511,0" Foreground="#505050" Grid.ColumnSpan="2"/>
                        <Ellipse Name="ellipseLapsInstalled" Grid.Column="2" Width="16" Height="16" Fill="Gray" Stroke="Black" StrokeThickness="0.5" VerticalAlignment="Center" Margin="84,0,495,0"/>
                        <TextBlock Name="lblLapsInstalledStatus" Grid.Column="2" Text="Not checked" VerticalAlignment="Center" Margin="117,0,307,0" Foreground="#1C1C1C"/>

                        <TextBlock Grid.Column="2" Text="LAPS GPO active:" VerticalAlignment="Center" Margin="297,0,200,0" Foreground="#505050"/>
                        <Ellipse Name="ellipseGpoStatus" Grid.Column="2" Width="16" Height="16" Fill="Gray" Stroke="Black" StrokeThickness="0.5" VerticalAlignment="Center" Margin="395,0,184,0" RenderTransformOrigin="29.087,-0.851"/>
                        <TextBlock Name="lblGpoStatus" Grid.Column="2" Text="Not checked" VerticalAlignment="Center" Margin="429,0,0,0" Foreground="#1C1C1C"/>

                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Policy details:" VerticalAlignment="Top" Margin="0,10,5,0" Foreground="#505050"/>
                        <TextBox Name="txtPolicyDetails" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" IsReadOnly="True" Margin="0,10,0,-236" FontFamily="Consolas" FontSize="11" Style="{StaticResource ModernTextBox}"/>
                    </Grid>
                    <TextBox Name="txtStatus" Text="Ready." DockPanel.Dock="Bottom" Margin="0,240,0,0" FontStyle="Italic" FontSize="10" IsReadOnly="True" Height="390" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" Style="{StaticResource ModernTextBox}" Background="#F8F8F8"/>
                </DockPanel>
            </Border>
        </Grid>

        <!-- Footer -->
        <Border Grid.Row="2" Background="#FF1C323C" Padding="10" CornerRadius="0,0,5,5" Margin="0,10,0,0" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0">
            <TextBlock Name="lblFooter" Text="$footerText" FontSize="11" HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="#e2e2e2"/>
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
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAh73uQVUwTQZlE
# X5yWxXbjYhBY1p7IzAKgljuUlXhGq6CCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# BgkqhkiG9w0BCQQxIgQgdK1G+Pn2yGYIsldu2KairVnWEYe5eEwkqK/B/55DJnEw
# DQYJKoZIhvcNAQEBBQAEggEAHd2HWfYpdegT/GwBTT1YstpWJBwWT69cPjQiLcAo
# jvK2QJAR9TirZDpKl6kuuN6YbmWbwOYpJTeS6BGOztUjfcw8uv0ytD9I9yp5Ms4d
# GiQwPmQ9KsUI73UpD2eAjWDFt5yqUB7XyFfO2yvZx7iQl6iiCjCa5sd8VAzNVf1J
# bIpuYDOLUHGAAXpnLYVHQsFVVPb/YNCvDU1QSU1D145yD5BFroK1R8guemyGG7MW
# qg2f+VV8codnslwiCfnLOaWkOb9kCCRmeCUKXUq2YJaFEtPD/a8cDbGVjnF5LC5w
# c0WlMCOV5p7b3lD6j9wMVtrxx5FGHllDIi2QCqd4K6ADn6GCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAwMjVaMC8GCSqGSIb3DQEJBDEiBCDkc3i75NAl
# HZ3GQAd3e8yx4UgDHmeOxAagHJnVIlAY5TANBgkqhkiG9w0BAQEFAASCAgBQa0XN
# rcShMfNYokJILBrRFiweoyz3AlesxTJ3yOKH4nWxgkMf6tQ/eQ8HuBQcUmdYA6YY
# bLTbdEjU9t/BJOmqAwW55cBebh2+IQidkrsq25sUgZaL5vu2BSNqNd8PYwkahLnA
# XcY3LkaG3Tur2GZ2fL/SDHPg9PITm/Egyo0GCH3OTkcZSgO/+GmD1mEWPYUbxgBY
# tKIn0SDftJXX6gm4LRv1ZA7B8asDYzqd25nDtN1iva4wCN+KHxfeXA3SjZ+XRfXx
# S62OeCAkFw2sxLVUutVUDqshiGb2jSeAh8fwCzvzTjzxgzT9ORdC4nQFj8YI5030
# qgILnsJaQbOt4XiLc+2n3yL0qIEdVComXvXyEOx+UTSsFnkzKE5H+KI2YDdbBrhG
# e8VuGnN3MApgVj7DjrmFz1kBWUu1YZpahp2ArjqXA47HP9ipdewvyY9P0gf+tj96
# mBofPkWg42KYXeNd214CasHQiyVDUvjfRW7lvMgjDx9geFbnNmTh9BX9k2KEn7lV
# WkpE+mN+re3LgMkhNC0YoU4+93lkap08seq4c9bDTpeK/q8iNlKqli+oI5OFyVtK
# dfQR0o/o6bIlGCWVcMphk8Q4Famhuft2dpF0RERKzIRgfLrToO90tWvCNvLOYsYe
# gVR/7wvXe6ShGERHBm86+IY8kxqBZrpYe84IfQ==
# SIG # End signature block
