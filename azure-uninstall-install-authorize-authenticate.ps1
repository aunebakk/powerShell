# & '.\install\azure-uninstall-install-authorize-authenticate.ps1' -doTest -doEcho -zenUninstall -zenInstall -zenAuthorize -zenAuthenticate -comment 'cool'
# & '.\install\azure-uninstall-install-authorize-authenticate.ps1' -doDevelopment -doEcho -zenUninstall -zenInstall -zenAuthorize -zenAuthenticate -comment 'cool'
param (
    [string]$scriptName = 'azure-uninstall-install-authorize-authenticate',
    [string]$scriptStyle = 'original', # original / task
    [string]$scriptStatus = 'status ( todos, learn, learned )',
    [string]$scriptDocumentation = 'fine tune, added offline option',

    [DateTime]$dateTimeStart = [System.DateTime]::UtcNow,
    [DateTime]$dateTimeStop = [System.DateTime]::UtcNow,
    [DateTime]$createdDateTime = '2018.06.20',
    [DateTime]$updateDateTime = '2018.07.05',   

    [switch]$doDevelopment = $false,
    [switch]$doTest = $false,
    [switch]$doStep = $false,
    [switch]$doEcho = $true,

    [switch]$zenUninstall = $false,
        [switch]$zenOnlyShowWhichToUninstall = $false,
    [switch]$zenInstall = $false,
    [switch]$zenInstallFromLocal = $false,
    [switch]$zenAuthorize = $false,
    [switch]$zenAuthenticate = $false,

    [switch]$rethrow = $false,
    [string]$comment = 'no comment',
    [switch]$sendMail = $false,
    [switch]$returnHtml = $false
)
##################################################################################################################
$taskName = 'locals'
#region
##################################################################################################################
[string] $taskLine = ''
[string] $answer = ''
[bool] $mailAnyway = $false
[string] $htmlLog = ''
#endregion
##################################################################################################################
$taskName = 'set startup location'
#region
##################################################################################################################
try {
    [string] $startupDirectory = ''
    if ($doDevelopment) { 
        [string] $startupDirectory = 'C:\SQL2XProjects' + '\' + 'sql2x' + '\' + 'Scripts'
    } elseif ($doTest) {
        [string] $startupDirectory = $home + '\' + 'sql2x' + '\' + 'Scripts'   
    }
    if ($startupDirectory -ne '') { 
        Set-Location $startupDirectory -ErrorAction:Stop
    }
} catch [Exception] {
    if ($rethrow) {
        Write-Host ($taskName + ' ' + 'Exception; ' + $_.Exception.Message)
        throw [Exception] ('Failed to; ' + $taskName)
    } else {
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + ('Exception:' + ' ' + $taskName + ' ' + '[' +  $_.Exception.Message + ']' + ' ' + 'Line:[' + $_.InvocationInfo.ScriptLineNumber + ']')
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) -ForegroundColor Red }
        $doDevelopment = $false; $doTest = $true
    }
}
#endregion
##################################################################################################################
$taskName = 'start script:'
#region
##################################################################################################################
$taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + $taskName `
        + ' ' + $MyInvocation.MyCommand.Name `
        + ' ' + $MyInvocation.MyCommand.Arguments
$htmlLog = $htmlLog + $taskLine + '<br>'
if ($doEcho) { Write-Host $taskLine }

if (-not (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).
        IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) { 
    $taskLine = ([System.DateTime]::UtcNow.ToString() + ' ' + 'pwd:' + ' ' + $pwd + ' ' + 'Not Admin')
    $htmlLog = $htmlLog + $taskLine + '<br>'
    if ($doEcho) { Write-Host $taskLine }
} else {
    $taskLine = ([System.DateTime]::UtcNow.ToString() + ' ' + 'pwd:' + ' ' + $pwd + ' ' + 'Admin')
    $htmlLog = $htmlLog + $taskLine + '<br>'
    if ($doEcho) { Write-Host $taskLine }
}

$taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'doDevelopment;' + '[' + $doDevelopment + ']' + ' ' + 'doTest;' + '[' + $doTest + ']' + ' ' + 'doStep;' + '[' + $doStep + ']' + ' ' + 'doEcho;' + '[' + $doEcho + ']' + ' ' + 'rethrow;' + '[' + $rethrow + ']'
$htmlLog = $htmlLog + $taskLine + '<br>'
if ($doEcho) { Write-Host $taskLine }

# check comment
if ([string]::IsNullOrWhiteSpace($comment)) {
    [string] $comment = Read-Host -Prompt ([System.DateTime]::UtcNow.ToString() + ' ' + 'Comment')
    if ([string]::IsNullOrWhiteSpace($comment)) { throw [Exception] "Need a comment" }
}
$taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'comment:' + ' ' + $comment
$htmlLog = $htmlLog + $taskLine + '<br>'
if ($doEcho) { Write-Host $taskLine }
#endregion
##################################################################################################################
$taskName = 'uninstall azure modules'
#region
##################################################################################################################
if ($zenUninstall) {
    try {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 
            # powershell version
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'powerShell version:' + ' ' + $PSVersionTable.PSVersion
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine }

            # azure version
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'azure version:' + ' ' + (get-module azure).version
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine }

            # find current azure modules
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'find current azure modules'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine }

            [System.Int32] $azureModules = Get-Module -ListAvailable | Where-Object { $_.Name -Match 'Azs' -Or $_.Name -Match 'Azure' -Or $_.Name -Match 'AzureRM.profile'  -Or $_.Name -Match 'AzureRM.sql' } | Measure-Object | Select-Object -ExpandProperty Count
            while ($azureModules -gt 0) {

                # log current azure module count
                $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'uninstalling' + ' ' + $azureModules + ' ' + 'azure modules'
                $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
                if ($doEcho) { Write-Host $taskLine }

                # | Select-Object -ExpandProperty Name
                # if uninstall fails repeatedly then delete them
                # Remove-Item $home\Documents\WindowsPowerShell\Modules\Azure* -Recurse -Force
                # $module = Get-Module -ListAvailable | Where-Object { $_.Name -Match 'AzureRM.WebSites' }
                # $module | Remove-Module -Force
                # $module | Uninstall-Module -AllVersions -Force
                [System.Array] $modules = Get-Module -ListAvailable | Where-Object { $_.Name -Match 'Azs' -Or $_.Name -Match 'Azure' -Or $_.Name -Match 'AzureRM.profile' -Or $_.Name -Match 'AzureRM.sql' }
                foreach ($module in $modules) {
                    # log azure module name
                    $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'uninstall' + ';' + ' ' + $module
                    $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
                    if ($doEcho) { Write-Host $taskLine }

                    if (!$zenOnlyShowWhichToUninstall) {
                        $module | Remove-Module -Force
                        $module | Uninstall-Module -AllVersions -Force

                        # check if implementation is still around ( copied from local? )
                        if (Test-Path $module.ModuleBase) {
                            # module path exists
                            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'module path exists' + ' ' + (Get-Item $module.ModuleBase ).Parent.FullName
                            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
                            if ($doEcho) { Write-Host $taskLine }

                            Remove-Item (Get-Item $module.ModuleBase ).Parent.FullName -Recurse -Force
                        }
                    }
                }

                [System.Int32] $azureModules = Get-Module -ListAvailable | Where-Object { $_.Name -Match 'Azs' -Or $_.Name -Match 'Azure' -Or $_.Name -Match 'AzureRM.profile'  -Or $_.Name -Match 'AzureRM.sql' } | Measure-Object | Select-Object -ExpandProperty Count
                # } else {
                #     [System.Int32] $azureModules = 0
                # }
            }
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    } catch [Exception] {
        if ($rethrow) {
            Write-Host ($taskName + ' ' + 'Exception; ' + $_.Exception.Message)
            throw [Exception] ('Failed to; ' + $taskName)
        } else {
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + ('Exception:' + ' ' + $taskName + ' ' + '[' +  $_.Exception.Message + ']' + ' ' + 'Line:[' + $_.InvocationInfo.ScriptLineNumber + ']')
            $htmlLog = $htmlLog + $taskLine + '<br>'
            if ($doEcho) { Write-Host ( $taskLine ) -ForegroundColor Red }
        }
    } finally {
        if ($module) { try { Remove-Variable -Name module } catch {}}
        if ($modules) { try { Remove-Variable -Name modules } catch {}}
        if ($azureModules) { try { Remove-Variable -Name azureModules } catch {}}
    }
}
#endregion
##################################################################################################################
$taskName = 'install azure modules'
#region
##################################################################################################################
if ($zenInstall) {
    try {
        [string] $local = ''

        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 
            # install latest azure modules
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

            # Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
            Install-Module -Name Azure              -Scope CurrentUser
            Install-Module -Name AzureRM.Sql        -RequiredVersion 4.6.0 -Scope CurrentUser
            Install-Module -Name AzureRM.Resources  -Scope CurrentUser
            Install-Module -Name AzureRM.Websites   -Scope CurrentUser

            Get-Module -ListAvailable

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'installed Azure and AzureRM.Sql'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine }
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    } catch [Exception] {
        if ($rethrow) {
            Write-Host ($taskName + ' ' + 'Exception; ' + $_.Exception.Message)
            throw [Exception] ('Failed to; ' + $taskName)
        } else {
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + ('Exception:' + ' ' + $taskName + ' ' + '[' +  $_.Exception.Message + ']' + ' ' + 'Line:[' + $_.InvocationInfo.ScriptLineNumber + ']')
            $htmlLog = $htmlLog + $taskLine + '<br>'
            if ($doEcho) { Write-Host ( $taskLine ) -ForegroundColor Red }
        }
    } finally {
        if ($local) { try { Remove-Variable -Name local } catch {}}
    }
}
#endregion
##################################################################################################################
$taskName = 'install azure modules from local'
#region
##################################################################################################################
if ($zenInstallFromLocal) {
    try {
        [string] $local = ''

        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 

            # copy over latest azure modules and hope that means they are installed
            Copy-Item `
                -Recurse `
                -Force `
                -LiteralPath ($home + '\install\modules') `
                -Destination ($home + '\Documents\WindowsPowerShell')

            import-Module ($home + '\Documents\WindowsPowerShell\Modules\Azure\5.3.0\Azure.psm1')
            import-Module ($home + '\Documents\WindowsPowerShell\Modules\Azure.Storage\4.3.1\Azure.Storage.psm1')
            import-Module ($home + '\Documents\WindowsPowerShell\Modules\AzureRM.profile\5.3.2\AzureRM.Profile.psm1')
            import-Module ($home + '\Documents\WindowsPowerShell\Modules\AzureRM.Resources\6.2.0\AzureRM.Resources.psm1')
            import-Module ($home + '\Documents\WindowsPowerShell\Modules\AzureRM.Sql\4.10.0\AzureRM.Sql.psm1')
            import-Module ($home + '\Documents\WindowsPowerShell\Modules\AzureRM.Websites\5.0.4\AzureRM.Websites.psm1')

            #Get-Module -ListAvailable

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'installed loads of modules from local'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine }
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    } catch [Exception] {
        if ($rethrow) {
            Write-Host ($taskName + ' ' + 'Exception; ' + $_.Exception.Message)
            throw [Exception] ('Failed to; ' + $taskName)
        } else {
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + ('Exception:' + ' ' + $taskName + ' ' + '[' +  $_.Exception.Message + ']' + ' ' + 'Line:[' + $_.InvocationInfo.ScriptLineNumber + ']')
            $htmlLog = $htmlLog + $taskLine + '<br>'
            if ($doEcho) { Write-Host ( $taskLine ) -ForegroundColor Red }
        }
    } finally {
        if ($local) { try { Remove-Variable -Name local } catch {}}
    }
}
#endregion
##################################################################################################################
$taskName = 'authorize azure'
#region
##################################################################################################################
if ($zenAuthorize) {
    try {
        [string] $local = ''

        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 

            # do this in a new session / process so that the current session is 'clean' for later uninstall of azure
            powerShell ' `
                [string] $username = ''xxx''; `
                [SecureString] $securePassword = ConvertTo-SecureString -String ''xxx'' -AsPlainText -Force; `
                $credentials = New-Object System.Management.Automation.PSCredential ($username, $securePassword); `
                Add-AzureRmAccount -Credential $credentials; `
                Select-AzureRmSubscription -SubscriptionName ''Free Trial''; `
                Get-AzureRmSubscription
                '
                #Set-AzureRmContext -SubscriptionId ""
            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'authorized'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine }
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    } catch [Exception] {
        if ($rethrow) {
            Write-Host ($taskName + ' ' + 'Exception; ' + $_.Exception.Message)
            throw [Exception] ('Failed to; ' + $taskName)
        } else {
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + ('Exception:' + ' ' + $taskName + ' ' + '[' +  $_.Exception.Message + ']' + ' ' + 'Line:[' + $_.InvocationInfo.ScriptLineNumber + ']')
            $htmlLog = $htmlLog + $taskLine + '<br>'
            if ($doEcho) { Write-Host ( $taskLine ) -ForegroundColor Red }
        }
    } finally {
        if ($local) { try { Remove-Variable -Name local } catch {}}
    }
}
#endregion
##################################################################################################################
$taskName = 'authenticate azure'
#region
##################################################################################################################
if ($zenAuthenticate) {
    try {
        [string] $local = ''

        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 
            # do this in a new session / process so that the current session is 'clean' for later uninstall of azure

            # find all databases in all resources in all azure sql servers
            powerShell ' `
                Get-AzureRmResourceGroup | Get-AzureRmSqlServer | Get-AzureRmSqlDatabase | Select-Object DatabaseName `
                '

            # un authorize ( correct word? )
            powerShell ' `
                Logout-AzureRmAccount -ErrorAction:SilentlyContinue `
                #Remove-AzureRmContext -ErrorAction:SilentlyContinue `
                Remove-AzureRmAccount -ErrorAction:SilentlyContinue `
                '

            # logout azure message
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'logout azure RM'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine }
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    } catch [Exception] {
        if ($rethrow) {
            Write-Host ($taskName + ' ' + 'Exception; ' + $_.Exception.Message)
            throw [Exception] ('Failed to; ' + $taskName)
        } else {
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + ('Exception:' + ' ' + $taskName + ' ' + '[' +  $_.Exception.Message + ']' + ' ' + 'Line:[' + $_.InvocationInfo.ScriptLineNumber + ']')
            $htmlLog = $htmlLog + $taskLine + '<br>'
            if ($doEcho) { Write-Host ( $taskLine ) -ForegroundColor Red }
        }
    } finally {
        if ($serverName) { try { Remove-Variable -Name serverName } catch {}}
    }
}
#endregion
##################################################################################################################
$taskName = 'mail status'
#region
##################################################################################################################
try {  
    $answer = 'yes'

    # log
    $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
    $htmlLog = $htmlLog + $taskLine + '<br>'
    if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
    elseif ($doEcho) { Write-Host ( $taskLine ) }

    if ($sendMail -and $answer -ne 'no' -and ($doDevelopment -or $doTest -or $mailAnyway)) { 
        $SMTPClient = New-Object System.Net.Mail.SmtpClient
        $SMTPClient.Port = 587
        $SMTPClient.Host = 'smtp.live.com'
        $SMTPClient.EnableSsl = $true
        $SMTPClient.Timeout = 10000;
        $SMTPClient.DeliveryMethod = ([System.Net.Mail.SmtpDeliveryMethod]::Network)
        $SMTPClient.UseDefaultCredentials = $false

        $SMTPClient.Credentials = New-Object System.Net.NetworkCredential( `
            'xxx' , `
            'xxx' `
            );

        $emailMessage = New-Object System.Net.Mail.MailMessage
        $emailMessage.From = 'xxx'
        $emailMessage.To.Add('xxx')
        $emailMessage.Subject = ($env:computername + ' : ' + $MyInvocation.MyCommand.Name)
        $emailMessage.Body = $htmlLog
        $emailMessage.IsBodyHtml = $true
        $emailMessage.BodyEncoding = ([System.Text.UTF8Encoding]::UTF8)
        $emailMessage.DeliveryNotificationOptions = ([System.Net.Mail.DeliveryNotificationOptions]::OnFailure)

        $SMTPClient.Send( $emailMessage )
        $SMTPClient.Dispose()
    }

    # log
    $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
    $htmlLog = $htmlLog + $taskLine + '<br>'
    if ($doEcho) { Write-Host ( $taskLine ) }
} catch [Exception] {
    if ($rethrow) {
        Write-Host ($taskName + ' ' + 'Exception; ' + $_.Exception.Message)
        throw [Exception] ('Failed to; ' + $taskName)
    } else {
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + ('Exception:' + ' ' + $taskName + ' ' + '[' +  $_.Exception.Message + ']' + ' ' + 'Line:[' + $_.InvocationInfo.ScriptLineNumber + ']')
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) -ForegroundColor Red }
    }
} finally {
    if ($SMTPClient) { try { Remove-Variable -Name SMTPClient } catch {} }
    if ($emailMessage) { try { Remove-Variable -Name emailMessage } catch {} }
}   
#endregion
##################################################################################################################
$taskName = 'end script:'
#region
##################################################################################################################
$dateTimeStop = [System.DateTime]::UtcNow
$taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + $taskName `
    + ' ' + '[' + $MyInvocation.MyCommand.Name `
    + ' ' + $MyInvocation.MyCommand.Arguments + ']' `
    + ' ' + 'from' + ' ' + '[' + $dateTimeStart + ']' `
    + ' ' + 'to' + ' ' + '[' + $dateTimeStop  + ']'
$htmlLog = $htmlLog + $taskLine + '<br>'
if ($doEcho) { Write-Host ( $taskLine ) }
#endregion
##################################################################################################################
$taskName = 'cleanup'
#region
##################################################################################################################
if ($startupDirectory) { try { Remove-Variable -Name startupDirectory } catch {} }

if ($mailAnyway) { try { Remove-Variable -Name mailAnyway } catch {} }
if ($answer) { try { Remove-Variable -Name answer } catch {} }
if ($taskLine) { try { Remove-Variable -Name taskLine } catch {} }
if ($taskName) { try { Remove-Variable -Name taskName } catch {} }

if ($returnHtml -and $htmlLog) { 
    try { $htmlLog; Remove-Variable -Name htmlLog; return } catch {} } 
elseif ($htmlLog) { 
    try { Remove-Variable -Name htmlLog } catch {} }
#endregion
