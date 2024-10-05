# Pfad für 32-Bit Office auf 64-Bit Windows
$OfficeRegPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"

# Pfad für 64-Bit Office auf 64-Bit Windows oder 32-Bit Office auf 32-Bit Windows
if (-not (Test-Path $OfficeRegPath)) {
    $OfficeRegPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
}

# Abfrage der Office-Version
$OfficeVersion = Get-ItemProperty -Path $OfficeRegPath -Name VersionToReport, ProductReleaseIds -ErrorAction SilentlyContinue

if ($OfficeVersion) {
    # Produktname bestimmen
    $ProductName = $OfficeVersion.ProductReleaseIds

    # Office Version bestimmen
    switch -Wildcard ($ProductName) {
        "*O365*" { $OfficeVersionDetected = "Office 365" }
        "*O365ProPlusRetail*" { $OfficeVersionDetected = "Office 365 ProPlus" }
        "*Professional2021*" { $OfficeVersionDetected = "Office 2021" }
        "*HomeBusiness2021Retail*" { $OfficeVersionDetected = "Office 2021" }
        "*Professional2019*" { $OfficeVersionDetected = "Office 2019" }
        "*Professional2016*" { $OfficeVersionDetected = "Office 2016" }
        default { 
            if ($OfficeVersion.VersionToReport -like "16.0.*") {
                $OfficeVersionDetected = "Office 2016/2019/2021 (genaue Version nicht bestimmbar)"
            } else {
                $OfficeVersionDetected = "Unbekannte Version"
            }
        }
    }

    Write-Output "Installierte Office-Version: $OfficeVersionDetected ($($OfficeVersion.VersionToReport))"
    #Ninja-Property-Set officeversion "$OfficeVersionDetected ($($OfficeVersion.VersionToReport)"
} else {
    Write-Output "Keine Office-Version gefunden oder Office ist nicht als Click-to-Run installiert."
    #Ninja-Property-Set officeversion "Kein Office gefunden"
}
