echo "#####################################"
echo "# Update AD proxy Adress von CSV    #"
echo "# Version 0.9                       #"
echo "# Author: premium3722               #"
echo "#####################################"


########Variablen zum anpassen! #########
# Definiere den Pfad zur CSV-Datei
$csvPath = "C:\TEMP\DEMO.CSV"
# Definiere das Trennzeichen für die CSV-Datei
$csvDelimiter = ';'

#######Default Variablen #############
$LogfilePath = "C:\Windows\Temp\rmm\logs"
$Logfile = "C:\Windows\Temp\rmm\logs\logfile.html"
$start_time = Get-Date

# Aktivieren des Debug-Modus (true oder false)
$debugMode = $false
$debugModefirst3Rows = $true

############################################

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


# Listen für hinzugefügte und nicht hinzugefügte UPNs
$addedAddresses = @()
$noChanges = @()

# Lese die CSV-Datei ein und gebe das Trennzeichen an
$csvData = Import-Csv -Path $csvPath -Delimiter $csvDelimiter
WriteLog "CSV Import" -ForegroundColor Black

# Überprüfe, ob die CSV-Daten korrekt geladen wurden
if ($csvData -eq $null -or $csvData.Count -eq 0) {
    Write-Host "Fehler: CSV-Datei konnte nicht geladen werden oder ist leer."
    WriteLog "Fehler: CSV-Datei konnte nicht geladen werden oder ist leer." -ForegroundColor Black
    exit
}
else
{
  WriteLog "OK: CSV-Datei hat Daten drin" -ForegroundColor Black
}

# Wenn der Debug-Modus aktiviert ist, nur die ersten 3 Zeilen nehmen
if ($debugModefirst3Rows) {
    Write-Host "Im Debug-Modus: Verarbeite nur die ersten 3 Zeilen."
     WriteLog "Im Debug-Modus: Verarbeite nur die ersten 3 Zeilen." -ForegroundColor Yellow
    $csvData = $csvData | Select-Object -First 3
}

# Schleife durch jeden Benutzer im CSV-Daten
foreach ($user in $csvData) {
    # Debugging: Ausgabe des gesamten Benutzerdatensatzes
    Write-Host "Verarbeite Benutzer: $($user | Format-Table -AutoSize | Out-String)"
    WriteLog "Verarbeite Benutzer: $($user | Format-Table -AutoSize | Out-String)" -ForegroundColor Black

    # Hol den UserPrincipalName aus der CSV-Datei
    $userPrincipalName = $user.UserPrincipalName
    WriteLog "$userPrincipalName aus CSV" -ForegroundColor Black
    
    # Überprüfe, ob der UserPrincipalName leer ist
    if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
        Write-Host "Fehler: UserPrincipalName für einen Benutzer ist leer."
        WriteLog "Fehler: UserPrincipalName für einen Benutzer ist leer." -ForegroundColor Black
        continue
    }

    # Suche den Benutzer im Active Directory
    $adUser = Get-ADUser -Filter { UserPrincipalName -eq $userPrincipalName } -Properties proxyAddresses

    # Überprüfe, ob der Benutzer gefunden wurde
    if ($adUser) {
        # Bereite die neuen Proxy-Adressen vor
        $x500Address = ($user.ProxyAddresses -split ",") | Where-Object { $_ -like "X500:*" } | ForEach-Object { $_.ToString() }
        $smtpAddressLower = ($user.ProxyAddresses -split ",") | Where-Object { $_ -cmatch "^smtp:" } | ForEach-Object { $_.ToString() }

        # Überprüfe und zeige an, was hinzugefügt werden würde
        $addressesToAdd = @()

        if ($x500Address -and $adUser.proxyAddresses -notcontains $x500Address) {
            $addressesToAdd += $x500Address
            WriteLog "Benutzer $adUser.proxyAddresses hat kein X500" -ForegroundColor Black
        }

        if ($smtpAddressLower -and $adUser.proxyAddresses -notcontains $smtpAddressLower) {
            $addressesToAdd += $smtpAddressLower
            WriteLog "Benutzer $adUser.proxyAddresses hat kein smtp sekundaer" -ForegroundColor Black
        }

        # Wenn Adressen hinzugefügt werden müssen
        if ($addressesToAdd.Count -gt 0) {
            if ($debugMode) {
                Write-Host "Im Debug-Modus: Die folgenden Adressen würden für $userPrincipalName hinzugefügt: $($addressesToAdd -join ', ')"
            } else {
                # Füge die neuen Adressen als String hinzu (nicht als PSObject)
                Set-ADUser -Identity $adUser -Add @{proxyAddresses = $addressesToAdd}
                Write-Host "Proxy-Adressen für $userPrincipalName aktualisiert: $($addressesToAdd -join ', ')"
                WriteLog "Proxy-Adressen für $userPrincipalName aktualisiert: $($addressesToAdd -join ', ')" -ForegroundColor Black
            }

            # UPN zur Liste der hinzugefügten Adressen hinzufügen
            $addedAddresses += $userPrincipalName
        } else {
            # UPN zur Liste der Benutzer ohne Änderungen hinzufügen
            $noChanges += $userPrincipalName
        }
    } else {
        Write-Host "Benutzer mit UserPrincipalName $userPrincipalName nicht gefunden."
        # Auch Benutzer, die nicht gefunden wurden, zur Liste "keine Änderungen" hinzufügen
        $noChanges += $userPrincipalName
    }
}

# Übersicht am Ende anzeigen
Write-Host "`n===== Zusammenfassung ====="
Write-Host "`nBenutzer, bei denen Proxy-Adressen hinzugefügt wurden:"
WriteLog "`nBenutzer, bei denen Proxy-Adressen hinzugefügt wurden:"
$addedAddresses | ForEach-Object { Write-Host $_ }
WriteLog "$addedAddresses | ForEach-Object { Write-Host $_ }"

Write-Host "`nBenutzer, bei denen keine Änderungen vorgenommen wurden:"
WriteLog "`nBenutzer, bei denen keine Änderungen vorgenommen wurden:"
$noChanges | ForEach-Object { Write-Host $_ }
WriteLog "$noChanges | ForEach-Object { Write-Host $_ }"
