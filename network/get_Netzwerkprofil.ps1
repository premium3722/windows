#############################################
Ersteller: Levin von Känel

#############################################
Prüft was für ein Netzwerkprofil (Öffentlich, Privat, Domain) der Netzwerkadapter ist
Bei Domain wird eine 0 ausgegeben und sonst eine 1.
Dies kann für PRTG genutzt werden
#############################################

$ConProf = Get-NetConnectionProfile | Select -ExpandProperty NetworkCategory
If ($ConProf -eq "DomainAuthenticated") {write-Host 0:Domain} else {write-Host 1:Non-Domain}

