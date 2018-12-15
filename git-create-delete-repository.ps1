<#
 github api: https://developer.github.com/
 .\sql2x\git\git-create-delete-repository.ps1 -doDevelopment -doEcho -comment 'testing'
 #>
param (
    [string]$scriptName = 'create and delete git repository',
    [string]$scriptStyle = 'original', # original / task
    [string]$scriptStatus = 'status ( todos, learn, learned )',
    [string]$scriptDocumentation = 'research',

    [DateTime]$dateTimeStart = [System.DateTime]::UtcNow,
    [DateTime]$dateTimeStop = [System.DateTime]::UtcNow,
    [DateTime]$createdDateTime = '2018.12.04',
    [DateTime]$updateDateTime = '2018.12.15',

    [switch]$doDevelopment = $false,
    [switch]$doTest = $false,
    [switch]$doStep = $false,
    [switch]$doEcho = $true,

    [string]$githubUser = 'xxx',
    [string]$githubToken = 'xxx',
    [string]$repositoryName = 'testY',
    [string]$fileName = 'NOTICE',
    [string]$fileContent = 'xyz',

    [switch]$doCreateRepository = $false,
    [switch]$doDeleteRepository = $false,

    [switch]$doFileAdd = $false,
    [switch]$doFileGet = $false,
    [switch]$doFileUpdate = $false,
    [switch]$doFileDelete = $false,

    [switch]$rethrow = $false,
    [string]$comment = 'no comment',
    [switch]$sendMail = $false,
    [switch]$returnHtml = $false,

    [switch]$whatIf = $false,
    [string]$emailAddressFrom = 'xxx',
    [string]$emailAddressFromPassword = 'xxx',
    [string]$emailAddressTo = 'xxx'
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
$taskName = 'create repository'
#region
##################################################################################################################
try {
    if ($doCreateRepository) {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) {

            if (!$whatIf) {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                $base64AuthInfo = [Convert]::ToBase64String( `
                    [Text.Encoding]::ASCII.GetBytes( `
                        ("{0}:{1}" -f `
                            $githubUser, `
                            $githubToken `
                            )));
                
                $authHeader = @{"Authorization"="Basic $base64AuthInfo"};
                
                $body = @{
                    name = $repositoryName;
                    description = '';
                    homepage = '';
                    private = $false;
                    has_issues = $true;
                    has_wiki = $true;
                    has_downloads = $true;
                    auto_init = $true;
                    gitignore_template = 'C';
                    license_template = 'unlicense';
                } | ConvertTo-Json -Compress;
                
                $repoCreationResult = 
                    Invoke-RestMethod `
                        -Uri 'https://api.github.com/user/repos' `
                        -Headers $authHeader `
                        -Method Post `
                        -Body $body;
            }

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'repository created'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine -ForegroundColor Green}
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    }
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
#endregion
##################################################################################################################
$taskName = 'delete repository'
#region
##################################################################################################################
try {
    if ($doDeleteRepository) {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) {

            if (!$whatIf) {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                $base64AuthInfo = [Convert]::ToBase64String( `
                    [Text.Encoding]::ASCII.GetBytes( `
                        ("{0}:{1}" -f `
                            $githubUser, `
                            $githubToken `
                            )));
                
                $authHeader = @{"Authorization"="Basic $base64AuthInfo"};
                
                $repoCreationResult = 
                    Invoke-WebRequest `
                        -Uri ('https://api.github.com/repos/' + $githubUser + '/' + $repositoryName) `
                        -Headers $authHeader `
                        -Method Delete;               
            }

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'repository deleted'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine -ForegroundColor Green}
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    }
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
#endregion
##################################################################################################################
$taskName = 'add file to repository'
#region
##################################################################################################################
try {
    if ($doFileAdd) {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) {

            if (!$whatIf) {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                $base64AuthInfo = [Convert]::ToBase64String( `
                    [Text.Encoding]::ASCII.GetBytes( `
                        ("{0}:{1}" -f `
                            $githubUser, `
                            $githubToken `
                            )));
                
                $authHeader = @{"Authorization"="Basic $base64AuthInfo"};
                
                $body = @{
                    message = ('Added file ' + $fileName);
                    content = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($fileContent));
                } | ConvertTo-Json -Compress;
                
                $gitUri = ('https://api.github.com/repos/{0}/contents/' + $fileName) -f ($githubUser + '/' + $repositoryName);
                $gitCreationResult = Invoke-RestMethod -Uri $gitUri -Headers $authHeader -Method Put -Body $body;
            }

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'file added to repository'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine -ForegroundColor Green}
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    }
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
#endregion
##################################################################################################################
$taskName = 'get file from repository'
#region
##################################################################################################################
try {
    if ($doFileGet) {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) {

            if (!$whatIf) {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                $base64AuthInfo = [Convert]::ToBase64String( `
                    [Text.Encoding]::ASCII.GetBytes( `
                        ("{0}:{1}" -f `
                            $githubUser, `
                            $githubToken `
                            )));
                
                $authHeader = @{"Authorization"="Basic $base64AuthInfo"};
                
                $gitUri = ('https://api.github.com/repos/{0}/contents/' + $fileName) -f ($githubUser + '/' + $repositoryName);
                $gitCreationResult = Invoke-RestMethod -Uri $gitUri -Headers $authHeader -Method Get;
            }

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'got file from repository'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine -ForegroundColor Green}
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    }
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
#endregion
##################################################################################################################
$taskName = 'update file in repository'
#region
##################################################################################################################
try {
    if ($doFileUpdate) {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) {

            if (!$whatIf) {
                # get file sha
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                $base64AuthInfo = [Convert]::ToBase64String( `
                    [Text.Encoding]::ASCII.GetBytes( `
                        ("{0}:{1}" -f `
                            $githubUser, `
                            $githubToken `
                            )));
                
                $authHeader = @{"Authorization"="Basic $base64AuthInfo"};
                
                $gitUri = ('https://api.github.com/repos/{0}/contents/' + $fileName) -f ($githubUser + '/' + $repositoryName);
                $gitCreationResult = Invoke-RestMethod -Uri $gitUri -Headers $authHeader -Method Get;

                # update file
                $body = @{
                    message = ('Changed file' + $fileName);
                    content = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($fileContent));
                    sha = $gitCreationResult.sha
                } | ConvertTo-Json -Compress;

                $gitUri = ('https://api.github.com/repos/{0}/contents/' + $fileName) -f ($githubUser + '/' + $repositoryName);
                $gitCreationResult = Invoke-RestMethod -Uri $gitUri -Headers $authHeader -Method Put -Body $body;
            }

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'file updated in repository'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine -ForegroundColor Green}
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    }
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
#endregion
##################################################################################################################
$taskName = 'delete file in repository'
#region
##################################################################################################################
try {
    if ($doFileDelete) {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) {

            if (!$whatIf) {
                # get file sha
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                $base64AuthInfo = [Convert]::ToBase64String( `
                    [Text.Encoding]::ASCII.GetBytes( `
                        ("{0}:{1}" -f `
                            $githubUser, `
                            $githubToken `
                            )));
                
                $authHeader = @{"Authorization"="Basic $base64AuthInfo"};
                
                $gitUri = ('https://api.github.com/repos/{0}/contents/' + $fileName) -f ($githubUser + '/' + $repositoryName);
                $gitCreationResult = Invoke-RestMethod -Uri $gitUri -Headers $authHeader -Method Get;

                # update file
                $body = @{
                    message = ('deleted file' + $fileName);
                    sha = $gitCreationResult.sha
                } | ConvertTo-Json -Compress;

                $gitUri = ('https://api.github.com/repos/{0}/contents/' + $fileName) -f ($githubUser + '/' + $repositoryName);
                $gitCreationResult = Invoke-RestMethod -Uri $gitUri -Headers $authHeader -Method Delete -Body $body;
            }

            # log
            $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'file deleted in repository'
            $htmlLog = $htmlLog + ' ' + $taskLine + '<br>'
            if ($doEcho) { Write-Host $taskLine -ForegroundColor Green}
        }

        # log
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'end:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doEcho) { Write-Host ( $taskLine ) }
    }
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

if ($returnHtml -and $htmlLog) { 
    try { $htmlLog; Remove-Variable -Name htmlLog; return } catch {} } 
elseif ($htmlLog) { 
    try { Remove-Variable -Name htmlLog } catch {} }
#endregion
