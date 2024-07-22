$webhookUrl = "Webhook_URL Dieses Beispiel ist mit Teams Webhook"

# Administratoren abfragen
$admins = net localgroup administratoren | Where-Object {$_ -AND $_ -notmatch "Der Befehl *"} | Select-Object -Skip 4
$pcName = $env:COMPUTERNAME
$adminsString = $admins -join ", "  # Administratoren mit Komma und Leerzeichen zusammenfügen

# BitLocker-Status abfragen
$bitlockerStatus = (Get-BitLockerVolume -MountPoint "C:").ProtectionStatus
$bitlockerStatusText = if ($bitlockerStatus -eq 1) { "Aktiviert" } else { "Deaktiviert" }

# Überprüfen, ob RDP aktiviert ist
$rdpStatus = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name 'fDenyTSConnections').fDenyTSConnections
$rdpStatusText = if ($rdpStatus -eq 0) { "Aktiviert" } else { "Deaktiviert" }

# Abfragen von Computerinformationen
$computerInfo = Get-ComputerInfo

# Informationen zusammenstellen
$csModel = $computerInfo.CsModel
$csDomain = $computerInfo.CsDomain
$biosReleaseDate = $computerInfo.BiosReleaseDate
$biosSerialNr = $computerInfo.BiosSeralNumber
$windowsProductName = $computerInfo.ProductName
$windowsVersion = $computerInfo.WindowsBuildLabEx

if ($admins.Count -eq 1 -and $admins -contains 'localadmin') {
    Write-Host "'localadmin' ist der einzige Administrator. Kein Webhook wird gesendet."
} else {
    # Definieren der Daten
    $dataJson = @{
        title = "Administrator List"
        description = "Folgender PC hat noch andere Administratoren als der Localadmin."
        pcName = $pcName
        adminsString = $adminsString
        bitlockerStatus = $bitlockerStatusText
        rdpStatus = $rdpStatusText  
        csModel = $csModel
        csDomain = $csDomain
        biosReleaseDate = $biosReleaseDate
        windowsProductName = $windowsProductName
        windowsVersion = $windowsVersion
        ticketUrl = ""
    }

    $AdaptiveCard = @{
        type = "message"
        attachments = @(
            @{
                contentType = "application/vnd.microsoft.card.adaptive"
                contentUrl = $null
                content = @{
                    "$schema" = "http://adaptivecards.io/schemas/adaptive-card.json"
                    type = "AdaptiveCard"
                    version = "1.0"
                    body = @(
                        @{
                            type = "TextBlock"
                            size = "Large"
                            weight = "Bolder"
                            text = $dataJson.title
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.description
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.pcName
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.adminsString  # Verwenden des zusammengefügten Strings
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.bitlockerStatus  # BitLocker-Status hinzufügen
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.csModel
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.csDomain
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.biosReleaseDate
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.windowsProductName
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.windowsVersion
                            wrap = $true
                        }
                    )
                    actions = @(
                        @{
                            type = "Action.OpenUrl"
                            title = "Ticket erstellen"
                            url = $dataJson.ticketUrl
                        }
                    )
                }
            }
        )
        admins = $adminsString  # Administratoren-Liste als eigenständiges Feld hinzufügen
    }

    $AdaptiveCardJson = $AdaptiveCard | ConvertTo-Json -Depth 10
    Write-Host "AdaptiveCardJson:" $AdaptiveCardJson  # Debug-Ausgabe

    # Direktes Senden des JSON als String
    Invoke-RestMethod -Method POST -Uri $webhookUrl -Body $AdaptiveCardJson -ContentType 'application/json'
}

Write-Host "Windows Version: " + $dataJson.windowsVersion
Write-Host "Windows Product Name: " + $dataJson.windowsProductName
Write-Host "BIOS Release Date: " + $dataJson.biosReleaseDate
Write-Host "Domain: " + $dataJson.csDomain
Write-Host "Model: " + $dataJson.csModel
Write-Host "BitLocker Status: " + $dataJson.bitlockerStatus 
Write-Host "Administrators: " + $dataJson.adminsString
Write-Host "PC Name: " + $dataJson.pcName
Write-Host "SN" + $biosSerialNr

