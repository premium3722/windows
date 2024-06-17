echo "#####################################"
echo "# Windows Configuration Designer    #"
echo "# Version 0.9                       #"
echo "# Author: Premium                   #"
echo "#####################################"

Write-Host "Winget installiert" -ForegroundColor Green

# Wait for network
function func_check_internet
{
do{
    $ping = Test-NetConnection '8.8.8.8' -InformationLevel Quiet
    if(!$ping){
        cls
        'Warte auf Netzwerk Verbindung' | Out-Host
        sleep -s 5
    }
    else
    {
    Write-Host "Netzwerk OK" -ForegroundColor Green
    }
} while(!$ping)
}

function func_win_updates
{
	func_check_internet
        #Installiere Windwos updates
        Write-Host "Vorbereitung Windowsupdates..."
        set-psrepository -name PSGallery -installationpolicy trusted
	Write-Host "Installiere PSWindowsUpdate"
        Install-Module PSWindowsUpdate
	Write-Host "Importiere PSWindowsUpdate"
        Import-Module PSWindowsUpdate
        Write-Host "Installieren Updates..."
        Get-WindowsUpdate -AcceptAll -Install
	Write-Host "Fertig Windows Updates"
}

function func_uninstall_bloatware
{
	$app_packages = 
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
	Get-AppxProvisionedPackage -Online | ?{$_.DisplayName -in $app_packages} | Remove-AppxProvisionedPackage -Online -AllUser
}

function func_uninstall-Office365 
{
	$officeProducts = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE 'Microsoft Office 365%'" 
	if ($officeProducts) 
	{
		foreach ($product in $officeProducts) 
		{
			Write-Output "Deinstalliere $($product.Name)..."
			$product.Uninstall() | Out-Null
			Write-Output "$($product.Name) wurde deinstalliert."
		}
	}
	else 
	{
		Write-Output "Keine Office 365-Produkte gefunden."
	}
}

function func_download_software
{
	# Scope user or machine
	$scope = 'machine'

	$packages = 
	[PSCustomObject]@{
    		Name  = "Adobe.Acrobat.Reader.64-bit"
    		Scope = $scope
	},
	[PSCustomObject]@{
    		Name  = "Google.Chrome"
    		Scope = $scope
	}
	$packages | % 
 	{
    		if ($_.Scope) 
      		{
        		winget install -e --id $_.Name --scope 'machine' --silent --accept-source-agreements
			Write-Host "Programme $_.Name wurde installiert" -ForegroundColor Green
    		}
    		else 
     		{
        		winget install -e --id $_.Name --silent --accept-source-agreements
    		}
	}
 }


Write-Host "Programme installiert" -ForegroundColor Green

Write-Host "Starte Internet Check funkction"
func_check_internet
Write-Host "Starte Software Download"
func_download_software
Write-Host "Starte Windows Updates funktion"
func_win_updates
Write-Host "Starte Uninstall-Bloatware funktion"
func_uninstall_bloatware
Write-Host "Starte Uninstall O365 funtion"
func_uninstall-Office365 

Write-Host "Ende" -ForegroundColor Green

pause
