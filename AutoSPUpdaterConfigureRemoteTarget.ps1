
# Configures the server for WinRM and WSManCredSSP
Write-Host -ForegroundColor White " - Configuring PowerShell remoting..."
$winRM = Get-Service -Name winrm
If ($winRM.Status -ne "Running") {Start-Service -Name winrm}
Set-ExecutionPolicy Bypass -Force
Enable-PSRemoting -Force
Enable-WSManCredSSP -Role Server -Force | Out-Null
# Increase the local memory limit to 1 GB
Set-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB 1024 -WarningAction SilentlyContinue

# Check if we are running this from an Azure File Share, and if so set the credentials for it
$0 = $myInvocation.MyCommand.Definition
$env:dp0 = [System.IO.Path]::GetDirectoryName($0)
$bits = Get-Item $env:dp0 | Split-Path -Parent
if ($bits -like "*file.core.windows.net*")
{
    $storageAccountFQDN = $bits -replace '\\\\',''
    $storageAccountFQDN,$null = $storageAccountFQDN -split '\\'
    $storageAccountPrimaryKey = '0Nf7A48tOByWIv4s6CSz/Y8PoRr6xJnFqRbf57+wMpRIKNEwDuKFhOYIsw2OzDXzP0gN7DsFCZUMVymhSUU5aQ=='
    # Get the storage account username from the FQDN portion of the path
    $storageAccountUsername,$null = $storageAccountFQDN -split "\."
    # Store credentials locally to access the Azure File Share
    Start-Process -FilePath cmdkey.exe -ArgumentList "/add:$storageAccountFQDN /user:$storageAccountUsername /pass:$storageAccountPrimaryKey" -Wait -NoNewWindow
}

#Get out of this PowerShell process
Stop-Process -Id $PID -Force