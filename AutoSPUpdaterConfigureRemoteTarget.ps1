
# Configures the server for WinRM and WSManCredSSP
Write-Host -ForegroundColor White " - Configuring PowerShell remoting..."
$winRM = Get-Service -Name winrm
If ($winRM.Status -ne "Running")
{
    Start-Service -Name winrm
}
# Only change the PowerShell execution policy if we need to
if ((Get-ExecutionPolicy) -ne "Unrestricted" -and (Get-ExecutionPolicy) -ne "Bypass")
{
    Write-Host -ForegroundColor White " - Setting PowerShell execution policy..."
    Set-ExecutionPolicy Bypass -Force
}
else
{
    Write-Host -ForegroundColor White " - PowerShell execution policy already set to `"$(Get-ExecutionPolicy)`"."
}
Enable-PSRemoting -Force
Enable-WSManCredSSP -Role Server -Force | Out-Null
# Increase the local memory limit to 1 GB
Set-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB 1024 -WarningAction SilentlyContinue

#Get out of this PowerShell process
Stop-Process -Id $PID -Force