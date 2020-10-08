
param (
    [string]$file_path = "C:\Temp",
    [string]$cert_path = "Cert:\CurrentUser\My\*99E9FCA2F8D1EFBD1BFB5A62D85CC1E1DB3438FA",
    [string]$ts_svr = "http://timestamp.digicert.com"
)

Set-AuthenticodeSignature $file_path -Certificate (Get-ChildItem -path $cert_path) -TimestampServer $ts_svr
