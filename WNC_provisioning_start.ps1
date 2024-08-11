echo "#####################################"
echo "# Windows Configuration Designer    #"
echo "# Version 0.9                       #"
echo "# Author: Premium                   #"
echo "#####################################"

$LogfilePath = "C:\Windows\Temp\rmm\logs"
$Logfile = "C:\Windows\Temp\rmm\logs\logfile.html"
$start_time = Get-Date

# Log Funktion
function WriteLog {
    Param (
        [string]$LogString,
        [string]$ForegroundColor = "black",
        [bool]$IsBold = $false
    )
    $Stamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp <span style='color: $ForegroundColor;'>$LogString</span><br>"
    Add-Content $Logfile -Value $LogMessage
}
WriteLog "Script start" -ForegroundColor Black

# Wait for network
function func_check_internet {
    do {
        $ping = Test-NetConnection '8.8.8.8' -InformationLevel Quiet
        if (!$ping) {
            Clear-Host
            Write-Host 'Warte auf Netzwerk Verbindung' -ForegroundColor Yellow
            WriteLog "Warte auf Netzwerk Verbindung" -ForegroundColor Yellow
            Start-Sleep -Seconds 5
        } else {
            Write-Host "Netzwerk OK" -ForegroundColor Green
            WriteLog "Netzwerk OK" -ForegroundColor Green
        }
    } while (!$ping)
}

function func_win_updates {
    func_check_internet
    # Installiere Windows Updates
    WriteLog "Vorbereitung Windowsupdates..." -ForegroundColor Cyan
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    WriteLog "Installiere PSWindowsUpdate" -ForegroundColor Cyan
    Install-Module PSWindowsUpdate -Force -ErrorAction Stop
    WriteLog "Importiere PSWindowsUpdate" -ForegroundColor Cyan
    Import-Module PSWindowsUpdate -ErrorAction Stop
    WriteLog "Installieren Updates..." -ForegroundColor Cyan
    Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot
    WriteLog "Fertig Windows Updates" -ForegroundColor Green
}

function func_uninstall_bloatware {
    WriteLog "Start Bloatware deinstallation"
    $app_packages = @(
        "Microsoft.WindowsCamera",
        "Clipchamp.Clipchamp",
        "Microsoft.WindowsAlarms",
        "Microsoft.549981C3F5F10", # Cortana
        "Microsoft.WindowsFeedbackHub",
        "microsoft.windowscommunicationsapps",
        "Microsoft.WindowsMaps",
        "Microsoft.ZuneMusic",
        "Microsoft.BingNews",
        "Microsoft.Todos",
        "Microsoft.ZuneVideo",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.OutlookForWindows",
        "Microsoft.People",
        "Microsoft.PowerAutomateDesktop",
        "MicrosoftCorporationII.QuickAssist",
        "Microsoft.ScreenSketch",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.WindowsSoundRecorder",
        "Microsoft.MicrosoftStickyNotes",
        "Microsoft.BingWeather",
        "Microsoft.Xbox.TCUI",
        "Microsoft.GamingApp",
        "Microsoft.Windows.Ai.Copilot.Provider"
    )
    Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -in $app_packages} | Remove-AppxProvisionedPackage -Online -AllUsers
    WriteLog "Bloatware deeinstalltion fertig"
}

function func_uninstall_Office365 {
    WriteLog "Starte Office365 deinstalltion"
    $officeProducts = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE 'Microsoft Office 365%'" 
    if ($officeProducts) {
        foreach ($product in $officeProducts) {
            WriteLog "Deinstalliere $($product.Name)..." -ForegroundColor Cyan
            $product.Uninstall() | Out-Null
            WriteLog "$($product.Name) wurde deinstalliert." -ForegroundColor Green
        }
    } else {
        WriteLog "Keine Office 365-Produkte gefunden." -ForegroundColor Yellow
    }
}

function func_download_software {
    WriteLog "Starte Software Download"
    # Scope user or machine
    $scope = 'machine'

    $packages = @(
        [PSCustomObject]@{ Name = "Adobe.Acrobat.Reader.64-bit"; Scope = $scope },
        [PSCustomObject]@{ Name = "Google.Chrome"; Scope = $scope }
    )
    $packages | ForEach-Object {
        if ($_.Scope) {
            winget install -e --id $_.Name --scope 'machine' --silent --accept-source-agreements
            WriteLog "Programm $_.Name wurde installiert" -ForegroundColor Green
        } else {
            winget install -e --id $_.Name --silent --accept-source-agreements
            WriteLog "Programm $_.Name wurde installiert Scope not Machine" -ForegroundColor Green
        }
    }
}

WriteLog "Starte Internet Check Funktion" -ForegroundColor Cyan
func_check_internet
WriteLog "Starte Software Download" -ForegroundColor Cyan
func_download_software
WriteLog "Starte Uninstall-Bloatware Funktion" -ForegroundColor Cyan
func_uninstall_bloatware
WriteLog "Starte Uninstall Office 365 Funktion" -ForegroundColor Cyan
func_uninstall_Office365
WriteLog "Starte Windows Updates Funktion" -ForegroundColor Cyan
func_win_updates

WriteLog "Ende" -ForegroundColor Green

pause
