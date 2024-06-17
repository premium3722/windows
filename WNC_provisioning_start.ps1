echo "#####################################"
echo "# Windows Configuration Designer    #"
echo "# Version 0.9                       #"
echo "# Author: Premium                   #"
echo "#####################################"

Write-Host "Winget installiert" -ForegroundColor Green

# Wait for network
function func_check_internet {
    do {
        $ping = Test-NetConnection '8.8.8.8' -InformationLevel Quiet
        if (!$ping) {
            Clear-Host
            Write-Host 'Warte auf Netzwerk Verbindung' -ForegroundColor Yellow
            Start-Sleep -Seconds 5
        } else {
            Write-Host "Netzwerk OK" -ForegroundColor Green
        }
    } while (!$ping)
}

function func_win_updates {
    func_check_internet
    # Installiere Windows Updates
    Write-Host "Vorbereitung Windowsupdates..." -ForegroundColor Cyan
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Write-Host "Installiere PSWindowsUpdate" -ForegroundColor Cyan
    Install-Module PSWindowsUpdate -Force -ErrorAction Stop
    Write-Host "Importiere PSWindowsUpdate" -ForegroundColor Cyan
    Import-Module PSWindowsUpdate -ErrorAction Stop
    Write-Host "Installieren Updates..." -ForegroundColor Cyan
    Get-WindowsUpdate -AcceptAll -Install
    Write-Host "Fertig Windows Updates" -ForegroundColor Green
}

function func_uninstall_bloatware {
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
}

function func_uninstall_Office365 {
    $officeProducts = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE 'Microsoft Office 365%'" 
    if ($officeProducts) {
        foreach ($product in $officeProducts) {
            Write-Host "Deinstalliere $($product.Name)..." -ForegroundColor Cyan
            $product.Uninstall() | Out-Null
            Write-Host "$($product.Name) wurde deinstalliert." -ForegroundColor Green
        }
    } else {
        Write-Host "Keine Office 365-Produkte gefunden." -ForegroundColor Yellow
    }
}

function func_download_software {
    # Scope user or machine
    $scope = 'machine'

    $packages = @(
        [PSCustomObject]@{ Name = "Adobe.Acrobat.Reader.64-bit"; Scope = $scope },
        [PSCustomObject]@{ Name = "Google.Chrome"; Scope = $scope }
    )
    $packages | ForEach-Object {
        if ($_.Scope) {
            winget install -e --id $_.Name --scope 'machine' --silent --accept-source-agreements
            Write-Host "Programm $_.Name wurde installiert" -ForegroundColor Green
        } else {
            winget install -e --id $_.Name --silent --accept-source-agreements
        }
    }
}

Write-Host "Programme installiert" -ForegroundColor Green

Write-Host "Starte Internet Check Funktion" -ForegroundColor Cyan
func_check_internet
Write-Host "Starte Software Download" -ForegroundColor Cyan
func_download_software
Write-Host "Starte Windows Updates Funktion" -ForegroundColor Cyan
func_win_updates
Write-Host "Starte Uninstall-Bloatware Funktion" -ForegroundColor Cyan
func_uninstall_bloatware
Write-Host "Starte Uninstall Office 365 Funktion" -ForegroundColor Cyan
func_uninstall_Office365

Write-Host "Ende" -ForegroundColor Green

pause
