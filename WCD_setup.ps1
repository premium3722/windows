$provisioning = "$($env:ProgramData)\provisioning"

ni $provisioning -ItemType Directory -Force | Out-Null

gci -Path . | ?{$_.name -ne "setup.ps1"} | %{
    cp $_.FullName (Join-Path -Path $provisioning -ChildPath $_.name) -Force
}

New-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "execute_provisioning" -Value ("cmd /c powershell.exe -ExecutionPolicy Bypass -File `"{0}\WCD_provisioning_start.ps1`"" -f $provisioning)
Set-ExecutionPolicy unrestricted


