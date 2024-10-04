echo "#####################################"
echo "# Update AD proxy Adress von CSV    #"
echo "# Version 0.9                       #"
echo "# Author: premium3722               #"
echo "#####################################"



########Variablen zum anpassen! #########
# Definiere den Pfad zur CSV-Datei
$csvPath = "C:\Temp\UserExport.CSV"
# Definiere das Trennzeichen für die CSV-Datei
$csvDelimiter = ';'
# Optional: Definiere die OU, in der gesucht werden soll (leer lassen, um im gesamten AD zu suchen)
$searchBaseOU = "OU=EmployeeAccounts,OU=Standort,OU=Kunde,DC=Schweiz,DC=lan"  # Beispiel für eine OU
#$searchBaseOU = $null  # Wenn du im gesamten AD suchen willst

#######Default Variablen #############
$LogfilePath = "C:\Temp\"
$Logfile = "C:\Temp\logfile.html"
$start_time = Get-Date

# Aktivieren des Debug-Modus (true oder false)
$debugMode = $false
$debugModefirst3Rows = $false

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
    Write-Host "Im Debug-Modus: Verarbeite nur die ersten 3 Zeilen." -ForegroundColor Magenta
     WriteLog "Im Debug-Modus: Verarbeite nur die ersten 3 Zeilen." -ForegroundColor Magenta
    $csvData = $csvData | Select-Object -First 3
}

# Schleife durch jeden Benutzer im CSV-Daten
foreach ($user in $csvData) {
    # Debugging: Ausgabe des gesamten Benutzerdatensatzes
    Write-Host "Verarbeite Benutzer: $($user | Format-Table -AutoSize | Out-String)"
    WriteLog "Verarbeite Benutzer: $($user.DisplayName)" -ForegroundColor Black

    # Hol den UserPrincipalName aus der CSV-Datei
    $userPrincipalName = $user.UserPrincipalName
    WriteLog "$userPrincipalName aus CSV" -ForegroundColor Black
    
    # Überprüfe, ob der UserPrincipalName leer ist
    if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
        Write-Host "Fehler: UserPrincipalName für einen Benutzer ist leer."
        WriteLog "Fehler: UserPrincipalName für einen Benutzer ist leer." -ForegroundColor Black
        continue
    }

    # Suchfilter für den Benutzer
    $filter = { UserPrincipalName -eq $userPrincipalName }

    # Suche den Benutzer im Active Directory, optional mit einer OU
    if ($searchBaseOU) 
    {
        Write-Host "Suche nur in $searchBaseOU"
        WriteLog "Suche nur in $searchBaseOU"
        $adUser = Get-ADUser -Filter $filter -Properties proxyAddresses -SearchBase $searchBaseOU
        
    } else {
        Write-Host "Suche in ganzem AD" 
        WriteLog "Suche in ganzem AD"
        $adUser = Get-ADUser -Filter $filter -Properties proxyAddresses
    }

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
                Write-Host "Im Debug-Modus: Die folgenden Adressen würden für $userPrincipalName hinzugefügt: $($addressesToAdd -join ', ')" -ForegroundColor Magenta
                WriteLog "Im Debug-Modus: Die folgenden Adressen würden für $userPrincipalName hinzugefügt: $($addressesToAdd -join ', ')" -ForegroundColor Magenta
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
    WriteLog " "
}

# Übersicht am Ende anzeigen
Write-Host "`n===== Zusammenfassung ====="
Write-Host "`nBenutzer, bei denen Proxy-Adressen hinzugefügt wurden:"
WriteLog "`nBenutzer, bei denen Proxy-Adressen hinzugefügt wurden:"
$addedAddresses | ForEach-Object { Write-Host $_ }
WriteLog "$addedAddresses"
Write-Host "`nBenutzer, bei denen keine Änderungen vorgenommen wurden:"
WriteLog "`nBenutzer, bei denen keine Änderungen vorgenommen wurden:"
$noChanges | ForEach-Object { Write-Host $_ }
WriteLog "$noChanges"
