$certname = "Name von Zertifikat (z.B vpn.schweiz.ch)"
{
  Get-ChildItem Cert:\LocalMachine\Root\ | where{$_.DnsNameList -eq $certname}
  WriteLog "Zertifikat besteht $certname"
  }
  catch
  {
    Write-Output "Zertifikat fehlt! $certname"
  }
}
