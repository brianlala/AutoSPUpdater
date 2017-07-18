<#
.SYNOPSIS
    Applies SharePoint 2010/2013/2016 updates (Service Packs + Cumulative/Public Updates) farm-wide, centrally from any server in the farm.
.DESCRIPTION
    Consisting of a module and a "launcher" script, AutoSPUpdater will install SharePoint 201x updates in two phases: binary installation and PSConfig (AKA
    the command-line equivalent of the "Products and Technologies Configuration Wizard"). AutoSPUpdater leverages PowerShell remoting and will test connectivity
    to other servers in the farm (automatically detected using Get-SPFarm) via ping, so this must be allowed through Windows Firewall. The script will prompt when
    the binary installation has completed on each server prior to running PSConfig. The script will also pause the SharePoint 2013 Search Service Application to 
    speed up patching (only required on SP2013). For best results, run the script from a UNC/shared path (NOT a mapped drive) e.g. "\\server\share$\SP\Scripts". 
    You can also run this from a regular local path but ONLY if the script and update files exist identically on each server in the farm. Currently, Azure file shares
    (e.g. *.file.core.windows.net) don't work as UNC sources, probably due to the way authentication is implemented. In general, you should make sure that all
    servers in your farm have connectivity and access to the path you run this script from.
.EXAMPLE
    .\AutoSPUpdaterLaunch.ps1 -patchPath C:\SP\2013\Updates -remoteAuthPassword fuzzyBunny99
.EXAMPLE
    & C:\SP\AutoSPInstaller\AutoSPUpdaterLaunch.ps1
.PARAMETER patchPath
    AutoSPUpdater will attempt to find updates located in the path structure used by AutoSPInstaller and AutoSPSourceBuilder (related projects). For example, if you
    are running AutoSPUpdater from C:\SP\AutoSPInstaller\, we will search for and attempt to install all updates found in C:\SP\201x\Updates (where 201x is the automatically-
    detected version of SharePoint). If this relative path doesn’t exist, the script will look in the “default” path used by AutoSPInstaller and AutoSPSourceBuilder – C:\SP\201x\Updates.
    Otherwise, you can just specify another path.
.PARAMETER remoteAuthPassword
    Optionally provide (in clear text, yikes) the password of the currently-logged in user for use in remote authentication to the other servers in the farm. If omitted, 
    the script will prompt you for it (in this case it will be obfuscated and encrypted). This parameter is only provided for maximum automation; normally it's best to leave it out.
.PARAMETER skipParallelInstall
    By default, AutoSPUpdater will install binaries on the local server first, then install binaries on each other server in the farm in parallel. This can significantly speed
    up patch installation. Use the -skipParallelInstall switch if you would instead like to install updates serially, one server at-a-time.
.PARAMETER useSqlSnapshot
    By default, AutoSPUpdater will attempt to use a SQL snapshot (only available if the SQL instance(s) are running Enterprise Edition. This can avoid unecessary downtime by pointing
    end-users to a read-only snapshot copy of the content database while the "real" database is being upgraded. Make sure your SQL server is indeed Enterprise Edition before attempting to use this option.
.LINK
    https://github.com/brianlala/autospsourcebuilder
    http://blogs.msdn.com/b/russmax/archive/2013/04/01/why-sharepoint-2013-cumulative-update-takes-5-hours-to-install.aspx
.NOTES
    Created & maintained by Brian Lalancette (@brianlala), 2012-2017.
#>
param
(
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()]
    [string]$patchPath,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()]
    [string]$remoteAuthPassword,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()]
    [Switch]$skipParallelInstall = $false,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()]
    [Switch]$useSqlSnapshot = $false
)

$servicesToStop = ("SPTimerV4","SPSearch4","OSearch14","OSearch15","OSearch16","SPSearchHostController")
# Same set of services, just in a slightly different order
$servicesToStart = ("SPSearchHostController","OSearch14","OSearch15","OSearch16","SPTimerV4","SPSearch4")

#region Check If Admin
# First check if we are running this under an elevated session. Pulled from the script at http://gallery.technet.microsoft.com/scriptcenter/1b5df952-9e10-470f-ad7c-dc2bdc2ac946
If (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning " - You must run this script under an elevated PowerShell prompt. Launch an elevated PowerShell prompt by right-clicking the PowerShell shortcut and selecting `"Run as Administrator`"."
    break
}
#endregion

#region Set Up Paths & Environment

$Host.UI.RawUI.WindowTitle = "-- $env:COMPUTERNAME (AutoSPUpdater) --"
$Host.UI.RawUI.BackgroundColor = "Black"
$0 = $myInvocation.MyCommand.Definition
$launchPath = [System.IO.Path]::GetDirectoryName($0)
$bits = Get-Item $launchPath | Split-Path -Parent
# Check if we are running this from an Azure File Share. This doesn't really work for some reason.
if ($bits -like "*file.core.windows.net*")
{
    $storageAccountFQDN = $bits -replace '\\\\',''
    $storageAccountFQDN,$null = $storageAccountFQDN -split '\\'
    $storageAccountPrimaryKey = ''
    # Get the storage account username from the FQDN portion of the path
    $storageAccountUsername,$null = $storageAccountFQDN -split "\."
    # Store credentials locally to access the Azure File Share
    Start-Process -FilePath cmdkey.exe -ArgumentList "/add:$storageAccountFQDN /user:$storageAccountUsername /pass:$storageAccountPrimaryKey" -Wait -NoNewWindow -LoadUserProfile
}
Write-Host -ForegroundColor White " - Loading SharePoint PowerShell Snapin..."
# Added the line below to match what the SharePoint.ps1 file implements (normally called via the SharePoint Management Shell Start Menu shortcut)
if (!($Host.Name -eq "ServerRemoteHost")) {$Host.Runspace.ThreadOptions = "ReuseThread"}
Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
Import-Module -Name "$launchPath\AutoSPUpdaterModule.psm1" -DisableNameChecking -Global -Force -ErrorAction Inquire
If (Confirm-LocalSession)
{
    Clear-Host
    if (!$startDate) {$startDate = Get-Date}
    StartTracing # Only start tracing if this is a local session
}
$spYears = @{"14" = "2010"; "15" = "2013"; "16" = "2016"}
$spVersions = @{"2010" = "14"; "2013" = "15"; "2016" = "16"}
if ($null -eq $spVer)
{
    [string]$spVer = (Get-SPFarm).BuildVersion.Major
    if (!$?)
    {
        Start-Sleep 10
        throw "Could not determine version of farm."
    }
}
$spYear = $spYears.$spVer
if ([string]::IsNullOrEmpty($patchPath))
{
    $patchPath = $bits+"\$spYear\Updates"
}
if (!(Test-Path -Path $patchPath -ErrorAction SilentlyContinue))
{
    Write-Host -ForegroundColor Yellow " - Patch path `"$patchPath`" does not appear to be valid; checking in standard location `"C:\SP\$spYear\Updates`"..."
    if (Test-Path -Path "C:\SP\$spYear\Updates")
    {
        $patchPath = "C:\SP\$spYear\Updates"
    }
    else
    {
        throw "Patch path `"$patchPath`" does not appear to be valid."
    }
}
Write-Host -ForegroundColor White " - `$patchPath is: $patchPath"
$PSConfig = "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\$spVer\BIN\psconfig.exe"
$PSConfigUI = "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\$spVer\BIN\psconfigui.exe"

UnblockFiles -path $patchPath
#endregion

#region Get Farm Servers & Credentials 
[array]$farmServers = (Get-SPFarm).Servers | Where-Object {$_.Role -ne "Invalid"}
if ($patchPath -like "*:*" -and $farmServers.Count -gt 1)
{
    Write-Host -ForegroundColor Yellow " - The path where updates reside ($patchPath) is identified by a local drive letter."
    Write-Host -ForegroundColor Yellow " - You should either use a UNC path that all farm servers can access (recommended),"
    Write-Host -ForegroundColor Yellow " - or create identical paths and copy all required files on each farm server."
    Write-Host -ForegroundColor White " - Ctrl-C to exit, or"
    Pause "continue updating" "y"
}

if ((Confirm-LocalSession) -and $farmServers.Count -gt 1) # Only do this stuff on the first (local) server, and only if we have other servers in the farm.
{
    Write-Host -ForegroundColor White " - Updating $env:COMPUTERNAME and additional farm server(s):"
    foreach ($farmserver in $farmServers | Where-Object {$_.Name -ne $env:COMPUTERNAME})
    {
        if (Confirm-LocalSession) {Write-Host -ForegroundColor White "  - $($farmserver.Name)"}
        [array]$remoteFarmServers += $farmServer.Name
    }
    if ([string]::IsNullOrEmpty($remoteAuthPassword)) {$password = Read-Host -AsSecureString -Prompt "Please enter the password for $env:USERDOMAIN\$env:USERNAME"}
    elseif ($remoteAuthPassword.GetType().Name -ne "SecureString")
    {
        $password = ConvertTo-SecureString -String $remoteAuthPassword -AsPlainText -Force
    }
    else
    {
        $password = $remoteAuthPassword
    }
    while ($credentialVerified -ne $true)
    {
        if ($password) # In case this is an automatic re-launch of the local script, re-use the password from the remote auth credential
        {
            Write-Host -ForegroundColor White " - Using pre-provided credentials..."
            $credential = New-Object System.Management.Automation.PsCredential $env:USERDOMAIN\$env:USERNAME,$password
        }
        if (!$credential) # Otherwise prompt for the remote auth or AutoAdminLogon credential
        {
            Write-Host -ForegroundColor White " - Prompting for remote/autologon credentials..."
            $credential = $host.ui.PromptForCredential("AutoSPUpdater - Remote/Automatic Install", "Enter Credentials for Remote/Automatic Authentication:", "$env:USERDOMAIN\$env:USERNAME", "NetBiosUserName")
        }
        $currentDomain = "LDAP://" + ([ADSI]"").distinguishedName
        $null,$user = $credential.Username -split "\\"
        if (($user -ne $null) -and ($credential.Password -ne $null)) {$passwordPlain = ConvertTo-PlainText $credential.Password}
        else
        {
            throw "Valid credentials are required for remote authentication."
            Pause "exit"
        }
        Write-Host -ForegroundColor White " - Checking credentials: `"$($credential.Username)`"..." -NoNewline
        $dom = New-Object System.DirectoryServices.DirectoryEntry($currentDomain,$user,$passwordPlain)
        If ($dom.Path -ne $null)
        {
            Write-Host -ForegroundColor Black -BackgroundColor Green "Verified."
            $credentialVerified = $true
        }
        else
        {
            Write-Host -BackgroundColor Red -ForegroundColor Black "Invalid - please try again."
            Remove-Variable -Name remoteAuthPassword -ErrorAction SilentlyContinue
            Remove-Variable -Name remoteAuthPasswordPlain -ErrorAction SilentlyContinue
            Remove-Variable -Name password -ErrorAction SilentlyContinue
            Remove-Variable -Name passwordPlain -ErrorAction SilentlyContinue
            Remove-Variable -Name credential -ErrorAction SilentlyContinue
        }
    }
}
#endregion

#region Stop AV
# Stop Symantec AV
[array]$avPaths = @("C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\Smc.exe","C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\12.1.1000.157.105\Bin64\Smc.exe")
foreach ($avPath in $avPaths)
{
    if (Test-Path -Path $avPath -ErrorAction SilentlyContinue)
    {
        Write-Host -ForegroundColor White " - Stopping antivirus (can speed up patching)..."
        Start-Process -FilePath $avPath -ArgumentList "-stop" -Wait -NoNewWindow
        break
    }
}
#endregion

#region Pause Search Service Application
# Only need to pause the Search Service Application(s) if running SharePoint 2013 and only attempt on the first (local) server in the farm
if ($spVer -eq 15 -and (Confirm-LocalSession))
{
    Request-SPSearchServiceApplicationStatus -desiredStatus Paused
}
#endregion

#region Stop Services
Write-Host -ForegroundColor White " - Temporarily disabling and stopping services..."
foreach ($service in $servicesToStop)
{
    $serviceExists = Get-Service $service -ErrorAction SilentlyContinue
    if ($serviceExists -and (Get-Service $service).Status -eq "Running")
    {
        Write-Host -ForegroundColor White "  - Stopping service $((Get-Service -Name $service).DisplayName)..."
        Set-Service -Name $service -StartupType Disabled
        Stop-Service -Name $service -Force
        New-Variable $service"WasRunning" -Value $true
    }
}
Write-Host -ForegroundColor White " - Services are now stopped."
#endregion

#region Install Remote Patch Binaries
<#
Write-Host -ForegroundColor White "-----------------------------------"
Write-Host -ForegroundColor White "| Automated SP$spYear patch script |"
Write-Host -ForegroundColor White "| Started on: $startDate |"
Write-Host -ForegroundColor White "-----------------------------------"
#>

# In case we are running this from a non-SharePoint farm server, only do these steps for farm member servers
if ($farmservers | Where-Object {$_ -match $env:COMPUTERNAME}) # Had to do it this way for PowerShell backward compatibility
{
    try
    {
        # We only want to Install-Remote if we aren't already *in* a remote session, and if there are actually remote servers to install!
        if ((Confirm-LocalSession) -and !([string]::IsNullOrEmpty($remoteFarmServers))) {Install-Remote -skipParallelInstall $skipParallelInstall -remoteFarmServers $remoteFarmServers -credential $credential -launchPath $launchPath -patchPath $patchPath}
    }
    catch
    {
        $EndDate = Get-Date
        Write-Host -ForegroundColor White "-----------------------------------"
        Write-Host -ForegroundColor White "| Automated SP$spYear patching script |"
        Write-Host -ForegroundColor White "| Started on: $startDate |"
        Write-Host -ForegroundColor White "| Aborted:    $EndDate |"
        Write-Host -ForegroundColor White "-----------------------------------"
        $aborted = $true
        if (!$scriptCommandLine -and (!(Confirm-LocalSession))) {Pause "exit"}
    }
    finally
    {}
}
# If the local server isn't a SharePoint farm server, just attempt remote installs
else
{
    if (Confirm-LocalSession)
    {
        Install-Remote -skipParallelInstall $skipParallelInstall -remoteFarmServers $remoteFarmServers -credential $credential -launchPath $launchPath -patchPath $patchPath
    }
}
#endregion

#region Install Local Patch Binaries
InstallUpdatesFromPatchPath -patchPath $patchPath -spVer $spVer
#endregion

#region Clear Configuration Cache
Clear-SPConfigurationCache
#endregion

#region Start Services
Write-Host -ForegroundColor White " - Re-enabling & starting services..."
ForEach ($service in $servicesToStart)
{
    if ($service -like "OSearch*") # The OSearch* service by default has startup type "Manual" so let's keep it that way
    {
        $startupType = "Manual"
    }
    else
    {
        $startupType = "Automatic"
    }
    if ((Get-Variable -Name $service"WasRunning" -ValueOnly -ErrorAction SilentlyContinue) -eq $true)
    {
        Set-Service -Name $service -StartupType $startupType
        Write-Host -ForegroundColor White "  - Starting service $((Get-Service -Name $service).DisplayName)..."
        Start-Service -Name $service
    }
}
Write-Host -ForegroundColor White " - Services are now started." 
#endregion

#region Get-SPProduct
Write-Host -ForegroundColor White " - Getting/updating local patch status (Get-SPProduct)..."
Get-SPProduct -Local
#endregion

#region Launch Central Admin - Servers In Farm
if (Confirm-LocalSession)
{
    $caWebApp = Get-SPWebApplication -IncludeCentralAdministration | ? {$_.IsAdministrationWebApplication}
    $caWebAppUrl = ($caWebApp.Url).TrimEnd("/")
    Write-Host -ForegroundColor White " - Launching `"$caWebAppUrl/_admin/FarmServers.aspx`"..."
    Write-Host -ForegroundColor White " - You can use this to track the status of each server's configuration."
    Start-Process "$caWebAppUrl/_admin/FarmServers.aspx" -WindowStyle Minimized
}
#endregion

#region Resume Search Service Application
# Only need to resume a paused Search Service Application(s) if running SharePoint 2013
if ($spVer -eq 15)
{
   Request-SPSearchServiceApplicationStatus -desiredStatus Online
}
#endregion

#region PSConfig
# Only upgrade databases if PSConfig is also required to be run
if (Test-UpgradeRequired -eq $true)
{
    #region Upgrade Content Databases
    # Get all servers in the farm running the Foundation Web Application service
    $foundationWebAppServiceInstances = Get-SPServiceInstance | Where-Object {$_.GetType().ToString() -eq "Microsoft.SharePoint.Administration.SPWebServiceInstance" -and $_.Name -ne "WSS_Administration"} # Need to filter out WSS_Administration because the Central Administration service instance shares the same Type as the Foundation Web Application Service
    # Get the service on the local server
    $foundationWebAppServiceInstance = $foundationWebAppServiceInstances | Where-Object {$_.Server.Address -eq "$env:COMPUTERNAME"}
    # See if the service is Online locally, or attempt to do the content DB upgrade if for some reason we can't query the Status of $foundationWebAppServiceInstance.Status
    if ($foundationWebAppServiceInstance.Status -eq "Online" -or $null -eq $foundationWebAppServiceInstance.Status)
    {
        Write-Host -ForegroundColor Cyan " - The script has determined that content databases may need to be upgraded."
        [array]$contentDatabases = Get-SPContentDatabase | Sort-Object Name
        Write-Host -ForegroundColor White " - Content databases found ($($contentDatabases.Count)):"
        foreach ($contentDatabase in $contentDatabases)
        {
            Write-Host -ForegroundColor Cyan "  - $($contentDatabase.Name)"
        }
        Write-Host -ForegroundColor White " - If any content databases are in a SQL Availability Group, you can `"Suspend Data Movement`" to speed up the upgrade."
        # Only need to pause if this isn't the only server in the farm
        if ($farmServers.Count -gt 1)
        {
            Write-Host -ForegroundColor Yellow " - Please ensure that all servers in the farm have completed the binary install phase before proceeding."
            Pause "proceed with content database upgrade" "y"
        }
        #region Launch Central Admin - Database Status
        if (Confirm-LocalSession)
        {
            $caWebApp = Get-SPWebApplication -IncludeCentralAdministration | ? {$_.IsAdministrationWebApplication}
            $caWebAppUrl = ($caWebApp.Url).TrimEnd("/")
            Write-Host -ForegroundColor White " - Launching `"$caWebAppUrl/_admin/DatabaseStatus.aspx`"..."
            Write-Host -ForegroundColor White " - You can use this to track the status of each content database upgrade."
            Start-Sleep -Seconds 3
            Start-Process "$caWebAppUrl/_admin/DatabaseStatus.aspx" -WindowStyle Minimized
        }
        #endregion
        $databaseUpgradeAttempted = $true
        Upgrade-ContentDatabases -spVer $spVer
    }
    else
    {
        Write-Host -ForegroundColor Yellow " - Content databases likely need to be upgraded, but this should be done from a web front-end server."
        Write-Host -ForegroundColor Yellow " - Please switch to a remote window with a prompt to upgrade content databases, and proceed from there prior to running PSConfig.exe."
        $databaseUpgradeAttempted = $false
    }
    #endregion
    # Good post for troubleshooting PSConfig: http://itgroove.net/mmman/2015/04/29/how-to-resolve-failures-in-the-sharepoint-product-config-psconfig-tool/
    Write-Host -ForegroundColor Cyan " - The script has determined that PSConfig needs to be run on this server ($env:COMPUTERNAME)."
    Write-Host -ForegroundColor White " - Running: $PSConfig"
    # Only need to pause if this isn't the only server in the farm, and if the DB upgrade hasn't already been attempted
    if ($farmServers.Count -gt 1 -and (!$databaseUpgradeAttempted))
    {
        Write-Host -ForegroundColor Yellow " - Please ensure that all servers in the farm have completed the binary install phase before proceeding."
        Pause "proceed with farm configuration wizard (PSConfig.exe)" "y"
    }
    # Display a message about no PSConfig progress over remote session
    if (!(Confirm-LocalSession))
    {
        Write-Host -ForegroundColor White " - Note that while PSConfig is running remotely there is no progress shown and it may take several minutes to complete."
        $passThruParameter = @{PassThru = $true}
    }
    else
    {
        $passThruParameter = @{}
    }
    $attemptNumber = 1
    Start-Process -FilePath $PSConfig -ArgumentList "-cmd upgrade -inplace b2b -wait -force -cmd applicationcontent -install -cmd installfeatures -cmd secureresources" -NoNewWindow -Wait @passThruParameter
    $PSConfigLastError = Check-PSConfig
    while (!([string]::IsNullOrEmpty($PSConfigLastError)) -and $attemptNumber -le 1)
    {
        Write-Warning $PSConfigLastError.Line
        Write-Host -ForegroundColor White " - An error occurred running PSConfig, trying again ($attemptNumber)..."
        Start-Sleep -Seconds 5
        $attemptNumber += 1
        Start-Process -FilePath $PSConfig -ArgumentList "-cmd upgrade -inplace b2b -wait -force -cmd applicationcontent -install -cmd installfeatures -cmd secureresources" -NoNewWindow -Wait -PassThru
        $PSConfigLastError = Check-PSConfig
    }
    if ($attemptNumber -ge 2)
    {
        if (Confirm-LocalSession)
        {
            Write-Host -ForegroundColor White " - After $attemptNumber attempts to run PSConfig, trying GUI-based..."
            Start-Process -FilePath $PSConfigUI -NoNewWindow -Wait
        }
    }
    if (Test-UpgradeRequired -eq $true)
    {
        Write-Host -ForegroundColor Yellow " - PSConfig has failed after $attemptNumber attempts. Please diagnose locally on $env:COMPUTERNAME."
    }
    else
    {
        Write-Host -ForegroundColor White " - PSConfig completed successfully."
    }
    Clear-Variable -Name PSConfigLastError -ErrorAction SilentlyContinue
    Clear-Variable -Name PSConfigLog -ErrorAction SilentlyContinue
    Clear-Variable -Name retryNum -ErrorAction SilentlyContinue
}
else
{
    Write-Host -ForegroundColor White " - The script has determined that running PSConfig is not required on this server ($env:COMPUTERNAME)."
}
#endregion

#region Start AV
# Start Symantec AV
[array]$avPaths = @("C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\Smc.exe","C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\12.1.1000.157.105\Bin64\Smc.exe")
foreach ($avPath in $avPaths)
{
    if (Test-Path -Path $avPath -ErrorAction SilentlyContinue)
    {
        Write-Host -ForegroundColor White " - (Re-)starting antivirus..."
        Start-Process -FilePath $avPath -ArgumentList "-start" -Wait -NoNewWindow
        break
    }
}
#endregion

#region Completed
Write-Host -ForegroundColor White " - Completed!`a"
$Host.UI.RawUI.WindowTitle = "-- Completed ($env:COMPUTERNAME) --"
$EndDate = Get-Date
try
{
    Stop-Transcript -ErrorAction SilentlyContinue
    if (!$?) {throw}
}
catch
{}
$script:isTracing = $false
#endregion

#region Launch Central Admin - Patch Status
if (Confirm-LocalSession)
{
    $caWebApp = Get-SPWebApplication -IncludeCentralAdministration | ? {$_.IsAdministrationWebApplication}
    $caWebAppUrl = ($caWebApp.Url).TrimEnd("/")
    Write-Host -ForegroundColor White " - Launching `"$caWebAppUrl/_admin/PatchStatus.aspx`"..."
    Write-Host -ForegroundColor White " - Review the patch status to ensure everything was applied OK."
    Start-Process "$caWebAppUrl/_admin/PatchStatus.aspx" -WindowStyle Minimized
}
#endregion

#region Wrap Up
If (!$aborted)
{
    If (Confirm-LocalSession) # Only do this stuff if this was a local session and it succeeded
    {
        Write-Host -ForegroundColor White "-----------------------------------"
        Write-Host -ForegroundColor White "| Automated SP$spYear patch script |"
        Write-Host -ForegroundColor White "| Started on: $startDate |"
        Write-Host -ForegroundColor White "| Completed:  $EndDate |"
        Write-Host -ForegroundColor White "-----------------------------------"
        try
        {
            Stop-Transcript -ErrorAction SilentlyContinue
            if (!$?) {throw}
        }
        catch
        {}
        $script:isTracing = $false
    }
    # Remove any lingering LogTime values in the registry
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\AutoSPUpdater\" -Name "LogTime" -ErrorAction SilentlyContinue
}
#endregion