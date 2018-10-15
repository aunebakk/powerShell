# & '.\zzz\documentation\Windows PowerShell Step by Step\desktop-windows.ps1' -doDevelopment
[Diagnostics.CodeAnalysis.SuppressMessageAttribute `
    ('PSPossibleIncorrectComparisonWithNull','')]
param (
    [string]$scriptName = 'Desktop Windows - saves and retrieves window state and positions for windows apps',
    [string]$scriptStyle = 'original', # original / task
    [string]$scriptStatus = 'Research & Development, issue when having more extended screens',
    [string]$scriptDocumentation = 'https://superuser.com/questions/1324007/setting-window-size-and-position-in-powershell-5-and-6',

    [DateTime]$dateTimeStart = [System.DateTime]::UtcNow,
    [DateTime]$dateTimeStop = [System.DateTime]::UtcNow,
    [DateTime]$createdDateTime = '2018.10.15',
    [DateTime]$updateDateTime = '2018.10.15',

    [switch]$doDevelopment = $false,
    [switch]$doTest = $false,
    [switch]$doStep = $false,
    [switch]$doEcho = $true,

    [switch]$zenLocate = $false,
    [switch]$zenSave = $false,
    [switch]$zenRestore = $false,

    [switch]$rethrow = $false,
    [string]$comment = 'no comment',
    [switch]$sendMail = $false,
    [switch]$returnHtml = $false
)
##################################################################################################################
$taskName = 'types'
#region
##################################################################################################################
Try{
    [void][Window]
} Catch {
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class Window {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("User32.dll")]
        public extern static bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);
        }
        public struct RECT
        {
        public int Left;        // x position of upper-left corner
        public int Top;         // y position of upper-left corner
        public int Right;       // x position of lower-right corner
        public int Bottom;      // y position of lower-right corner
        }
"@
}

Add-Type -AssemblyName UIAutomationClient
class DesktopWindow {
    [System.String]$processName
    [System.String]$windowVisualState
    [int]$left;        # x position of upper-left corner
    [int]$top;         # y position of upper-left corner
    [int]$right;       # x position of lower-right corner
    [int]$bottom;      # y position of lower-right corner
}
#endregion
##################################################################################################################
$taskName = 'locals'
#region
##################################################################################################################
[string] $taskLine = ''
[string] $answer = ''
[bool] $mailAnyway = $false
[string] $htmlLog = ''

[DesktopWindow[]] $desktopWindows = $null
[string] $desktopWindowsFile = $home + '\' + 'sql2x' + '\' + 'Scripts' + '\' + 'garbage' + '\' + 'desktop-windows.xml'
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
#endregion
##################################################################################################################
$taskName = 'Locate desktop windows state, position and size'
#region
##################################################################################################################
if ($zenLocate -or ($zenSave)) {
    try {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 
            # get windows with a handle ( apps )
            $processList = `
                Get-Process | Where-Object { $_.MainWindowHandle -ne 0 }

            $processList | ForEach-Object {

                # process information
                [int] $processId = $_.Id
                [System.ComponentModel.Component] $process = `
                    Get-Process | Where-Object { $_.Id -eq $processId } | Select-Object -First 1

                # new desktop window object
                $desktopWindow = New-Object DesktopWindow
                $desktopWindow.processName = $process.ProcessName

                # visual state of process
                try {
                    $actionElement = [System.Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
                    if ($actionElement -ne $null) {
                        $windowPattern = $actionElement.GetCurrentPattern([System.Windows.Automation.WindowPatternIdentifiers]::Pattern)
                        if ($windowPattern -ne $null) {
                            $desktopWindow.windowVisualState = $windowPattern.Current.WindowVisualState
                        }
                    }
                } catch { 
                    $desktopWindow.windowVisualState = 'N/A'
                }

                # window position and size
                $rectangle = New-Object RECT
                foreach ( $handle in $process.MainWindowHandle ) {      
                    if ( $handle -eq [System.IntPtr]::Zero ) { Continue }

                    if ([Window]::GetWindowRect(    
                                    $handle,
                                    [ref]$rectangle
                                    )) {
                        $desktopWindow.left = $rectangle.Left
                        $desktopWindow.top = $rectangle.Top
                        $desktopWindow.right = $rectangle.Right
                        $desktopWindow.bottom = $rectangle.Bottom
                    }
                }

                $desktopWindows += $desktopWindow
            }

            if ($doEcho) {
                # host window locations
                Clear-Host
                foreach ( $desktopWindow in $desktopWindows ) {
                    Write-Host ($desktopWindow.ProcessName + ' ' + $desktopWindow.windowVisualState + ' ' + $desktopWindow.Left + ' ' + $desktopWindow.Right + ' ' + $desktopWindow.Top + ' ' + $desktopWindow.Bottom)
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
        if ($local) { try { Remove-Variable -Name local } catch {}}
    }
}
#endregion
##################################################################################################################
$taskName = 'Save desktop windows state, position and size'
#region
##################################################################################################################
if ($zenSave) {
    try {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 
            # save desktop windows location
            ConvertTo-Json -InputObject $desktopWindows | Out-File $desktopWindowsFile -Encoding Ascii
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
$taskName = 'Restore desktop windows state, position and size'
#region
##################################################################################################################
if ($zenRestore) {
    try {
        # log
        $answer = 'yes'
        $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'start:' + ' ' + $taskName
        $htmlLog = $htmlLog + $taskLine + '<br>'
        if ($doStep) { $answer = Read-Host -Prompt ( $taskLine ) }
        elseif ($doEcho) { Write-Host ( $taskLine ) }

        if ($answer -ne 'no' -and ($doDevelopment -or $doTest)) { 
            # forget desktop windows location
            if ($desktopWindows) { try { Remove-Variable -Scope:Script -Name desktopWindows } catch {} }

            # restore desktop windows location
            $jsonDesktopWindow = Get-Content ($desktopWindowsFile) -Raw
            $desktopWindows = ConvertFrom-Json -InputObject $jsonDesktopWindow

            foreach ( $desktopWindow in $desktopWindows ) {

                # log
                $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'move app' + ':' `
                    + ' ' + ($desktopWindow.ProcessName + ' ' + $desktopWindow.windowVisualState + ' ' + $desktopWindow.Left + ' ' + $desktopWindow.Right + ' ' + $desktopWindow.Top + ' ' + $desktopWindow.Bottom)
                $htmlLog = $htmlLog + $taskLine + '<br>'
                if ($doEcho) { Write-Host ( $taskLine ) }

                $process = (Get-Process -Name $desktopWindow.ProcessName -ErrorAction:SilentlyContinue)
                if ($process -ne $null) {
                    $handles = (Get-Process -Name $desktopWindow.ProcessName).MainWindowHandle
                    foreach ( $handle in $handles ) {
                        if ( $handle -eq [System.IntPtr]::Zero ) { Continue }

                        $rectangle = New-Object RECT
                        if ([Window]::GetWindowRect($handle,[ref]$rectangle)) {
                            if (!([Window]::MoveWindow(
                                    $handle, 
                                    $desktopWindow.Left,
                                    $desktopWindow.Top,
                                    $desktopWindow.Right - $desktopWindow.Left,
                                    $desktopWindow.Bottom - $desktopWindow.Top, 
                                    $True
                                    )) ) {
       
                                # log
                                $taskLine = [System.DateTime]::UtcNow.ToString() + ' ' + 'failed to move app' + ':'
                                    + ' ' + ($desktopWindow.ProcessName + ' ' + $desktopWindow.windowVisualState + ' ' + $desktopWindow.Left + ' ' + $desktopWindow.Right + ' ' + $desktopWindow.Top + ' ' + $desktopWindow.Bottom)
                                $htmlLog = $htmlLog + $taskLine + '<br>'
                                if ($doEcho) { Write-Host ( $taskLine ) }
                            }
                        }
                    }
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
        if ($jsonDesktopWindow) { try { Remove-Variable -Name jsonDesktopWindow } catch {}}
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
if ($desktopWindowsFile) { try { Remove-Variable -Name desktopWindowsFile } catch {} }
if ($desktopWindows) { try { Remove-Variable -Name desktopWindows } catch {} }
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
