echo "#####################################"
echo "# Windows Configuration Designer    #"
echo "# Version 0.9                       #"
echo "# Author: Premium                   #"
echo "#####################################"


# Update winget
ls "$($env:ProgramData)\provisioning\Microsoft.VCLibs.x64.*", 
"$($env:ProgramData)\provisioning\Microsoft.UI.Xaml.*", 
"$($env:ProgramData)\provisioning\Microsoft.DesktopAppInstaller_*" | %{
    "Updating $($_.Name)" | Out-Host
    Add-AppxPackage $_.FullName -ErrorAction SilentlyContinue
}
Write-Host "Winget installiert" -ForegroundColor Green

# Wait for network
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

$packages | % {
    if ($_.Scope) {
        winget install -e --id $_.Name --scope 'machine' --silent --accept-source-agreements
	Write-Host "Programme $_.Name wurde installiert" -ForegroundColor Green
    }
    else {
        winget install -e --id $_.Name --silent --accept-source-agreements
    }
}


Write-Host "Programme installiert" -ForegroundColor Green

pause
