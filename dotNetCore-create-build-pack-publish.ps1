# '.\zzz\documentation\Developing Microsoft Azure Solutions\dotNetCore-create-build-pack-publish.ps1' -doDevelopment -projectName Tiney2 -zenPublish -zenClean -zenCreate -zenPack -zenBuild -sendMail -comment 'fricking cool'
param (
    [string]$scriptName = 'For dotNetCore, create build pack and publish a project',
    [string]$scriptStyle = 'original', # original / task
    [string]$scriptStatus = 'status ( todos, learn, learned )',
    [string]$scriptDocumentation = `
           'a) this script is meant to fix this issue: https://github.com/Azure/azure-powershell/issues/6374 | ' `
         + 'b) Publish-AzureWebSiteProject AppOffline switch is not available, but previous restart should do it, follow this thread: https://github.com/Azure/azure-powershell/issues/669',
    [DateTime]$dateTimeStart = [System.DateTime]::UtcNow,
    [DateTime]$dateTimeStop = [System.DateTime]::UtcNow,
    [DateTime]$createdDateTime = '2018.08.19',
    [DateTime]$updateDateTime = '2018.08.19',

    [switch]$doDevelopment = $false,
    [switch]$doTest = $false,
    [switch]$doStep = $false,
    [switch]$doEcho = $true,

    [switch]$rethrow = $false,
    [string]$comment = 'no comment',
    [switch]$sendMail = $false,
    [switch]$returnHtml = $false,

    [string]$emailAddressFrom = 'xxx',
    [string]$emailAddressFromPassword = 'xxx',
    [string]$emailAddressTo = 'xxx',

    [switch]$zenClean = $false,
    [switch]$zenCreate = $false,
        [string]$projectName = '',
        [string]$dotNetCoreTemplate = 'razor', # console, classlib, web, mvc, webapi, sln
    [switch]$zenRun = $false,

    [switch]$zenBuild = $false,
    [switch]$zenPack = $false,
    [switch]$zenPublish = $false
)
##################################################################################################################
$taskName = 'locals'
#region
##################################################################################################################
[string] $taskLine = ''
[string] $answer = ''
[bool] $mailAnyway = $false
[string] $htmlLog = ''

[string] $creationsDirectory = 'C:\SQL2XProjects'
[string] $projectDirectory = $creationsDirectory + '\' + $projectName
[string] $publishDirectory = 'bin\Debug\netcoreapp2.0\publish\'
[string] $publish2AzureDirectory = 'bin\azure'
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
$taskName = 'help'
#region
# log scriptName
$taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'Name' + '; ' + $scriptName
$htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
if ($doEcho) { Write-Host $taskLine }
# log scriptStyle
$taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'Style' + '; ' + $scriptStyle
$htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
if ($doEcho) { Write-Host $taskLine }
# log scriptStatus
$taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'Status' + '; ' + $scriptStatus
$htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
if ($doEcho) { Write-Host $taskLine }
# log scriptDocumentation
$taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'Documentation' + '; ' + $scriptDocumentation
$htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
if ($doEcho) { Write-Host $taskLine }
##################################################################################################################
#endregion
##################################################################################################################
$taskName = 'zenClean'
#region
##################################################################################################################
if ($zenClean) {
    try {
        [string] $backupDirectory = ($creationsDirectory + '\' + 'bak' + '\' + $projectName + '.')

        # make sure backup directory follows windows standard, change colon to period
        $backupDirectory = $backupDirectory + [System.DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ').Replace(':', ' ') 

        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($projectName.Length -eq 0) {
            throw [Exception] ('ProjectName is empty')
        } 
        
        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 
            if (Test-Path($projectDirectory)) {
                if (Test-Path($creationsDirectory + '\' + 'bak')) {
                    Move-Item `
                        -LiteralPath $projectDirectory `
                        -Destination $backupDirectory 

                    # log
                    $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'Project' + ';' + ' ' + '[' + $projectName + ']' + ' ' + 'moved to' + ' ' + '[' + $backupDirectory + ']'
                    $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
                    if ($doEcho) { Write-Host $taskLine }
                }
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
        if ($backupDirectory) { try { Remove-Variable -Name backupDirectory } catch {}}
    }
}
#endregion
##################################################################################################################
$taskName = 'zenCreate'
#region
##################################################################################################################
if ($zenCreate) {
    try {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 

            dotnet new $dotNetCoreTemplate --name $projectName --output $projectDirectory

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'project created'
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
    }
}
#endregion
##################################################################################################################
$taskName = 'zenRun'
#region
##################################################################################################################
if ($zenRun) {
    try {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 

            Start-Process 'microsoft-edge:http://localhost:5000' # todo, make this a parameter
            
            dotnet run --project ($projectDirectory + '\' + $projectName + '.csproj')

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'project ran'
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
$taskName = 'zenBuild'
#region
##################################################################################################################
if ($zenBuild) {
    try {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 
            dotnet publish ($projectDirectory + '\' + $projectName + '.csproj') --output ($projectDirectory + '\' + $publishDirectory)

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'project built'
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
    }
}
#endregion
##################################################################################################################
$taskName = 'zenPack'
#region
##################################################################################################################
if ($zenPack) {
    try {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + ($projectDirectory + '\' + $publish2AzureDirectory)
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine }

            # check if previous zip exists
            if (Test-Path(($projectDirectory + '\' + $publish2AzureDirectory + '\' + 'publish.zip'))) {
                Remove-Item `
                    -Path ($projectDirectory + '\' + $publish2AzureDirectory + '\' + 'publish.zip') 
            }

            # check if directory exists
            if (!(Test-Path(($projectDirectory + '\' + $publish2AzureDirectory)))) {
                New-Item `
                    -Path ($projectDirectory + '\' + $publish2AzureDirectory) `
                    -ItemType Directory
            }

            # zip new publish file
            Add-Type -Assembly System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::
                CreateFromDirectory(
                    ($projectDirectory + '\' + $publishDirectory),
                    ($projectDirectory + '\' + $publish2AzureDirectory + '\' + 'publish.zip'),
                    [System.IO.Compression.CompressionLevel]::Optimal, 
                    $false,  #includeBaseDirectory
                    [System.Text.Encoding]::Default
                    )

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'build packed'
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
    }
}
#endregion
##################################################################################################################
$taskName = 'zenPublish'
#region
##################################################################################################################
if ($zenPublish) {
    try {
        [string] $local = ''

        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 

            (Get-AzureRmWebApp).GetEnumerator() | # Get-AzureRmWebApp can use -Name instead, but this illustrates the power of .GetEnumerator() 
                Where-Object {$_.Name -eq $projectName} | 
                Restart-AzureRmWebApp 

            # pause for restart to go through
            Start-Sleep -Seconds 60

            Publish-AzureWebSiteProject  `
                -Name $projectName `
                -Package ($projectDirectory + '\' + $publish2AzureDirectory + '\' + 'publish.zip')

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + ''
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
            $emailAddressFrom , `
            $emailAddressFromPassword `
            );

        $emailMessage = New-Object System.Net.Mail.MailMessage
        $emailMessage.From = $emailAddressFrom
        $emailMessage.To.Add($emailAddressTo)
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

if ($publish2AzureDirectory) { try { Remove-Variable -Name publishDirectory } catch {}}
if ($publishDirectory) { try { Remove-Variable -Name publishDirectory } catch {}}
if ($creationsDirectory) { try { Remove-Variable -Name creationsDirectory } catch {}}
if ($projectDirectory) { try { Remove-Variable -Name projectDirectory } catch {}}

if ($returnHtml -and $htmlLog) { 
    try { $htmlLog; Remove-Variable -Name htmlLog; return } catch {} } 
elseif ($htmlLog) { 
    try { Remove-Variable -Name htmlLog } catch {} }
#endregion
