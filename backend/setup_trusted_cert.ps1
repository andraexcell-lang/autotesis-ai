$cert = Get-ChildItem Cert:\CurrentUser\TrustedPeople | Where-Object { $_.Subject -match "localhost" } | Select-Object -First 1
if (-not $cert) {
    Write-Host "Creating new self signed cert in TrustedPeople..."
    $cert = New-SelfSignedCertificate -DnsName 'localhost' -CertStoreLocation 'Cert:\CurrentUser\My'
    Move-Item -Path $cert.PSPath -Destination 'Cert:\CurrentUser\TrustedPeople'
}

Write-Host "Found cert with Thumbprint: $($cert.Thumbprint)"

$pfxPassword = ConvertTo-SecureString -String "autotesis123" -Force -AsPlainText
$pfxPath = Join-Path $PSScriptRoot "server.pfx"
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPassword

Write-Host "Exported PFX to: $pfxPath"
