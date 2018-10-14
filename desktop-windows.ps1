<#

Desktop Windows - retrieves window state and positions for windows apps

Drawn on experience from:
    https://superuser.com/questions/1324007/setting-window-size-and-position-in-powershell-5-and-6
    https://community.spiceworks.com/topic/1739320-powershell-manipulate-window-size-and-position-ps4

State:
    R&D

#>
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
# get windows state, x, y, left, top, width, height for all active apps
Add-Type -AssemblyName UIAutomationClient

# get windows with a handle ( apps )
$processList = `
    Get-Process | Where-Object { $_.MainWindowHandle -ne 0 }

# $processList
# break

class DesktopWindow {
    [System.String]$processName
    [System.String]$windowVisualState
    [int]$left;        # x position of upper-left corner
    [int]$top;         # y position of upper-left corner
    [int]$right;       # x position of lower-right corner
    [int]$bottom;      # y position of lower-right corner
}

[DesktopWindow[]] $desktopWindows = $null

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

[string] $desktopWindowsFile = $home + '\' + 'sql2x' + '\' + 'Scripts' + '\' + 'garbage' + '\' + 'desktop-windows.xml'

# save desktop windows location
ConvertTo-Json -InputObject $desktopWindows | Out-File $desktopWindowsFile -Encoding Ascii

# forget desktop windows location
if ($desktopWindows) { try { Remove-Variable -Scope:Script -Name desktopWindows } catch {} }

# restore desktop windows location
$jsonDesktopWindow = Get-Content ($desktopWindowsFile) -Raw
$desktopWindows = ConvertFrom-Json -InputObject $jsonDesktopWindow

# host window locations
Clear-Host
foreach ( $desktopWindow in $desktopWindows ) {
    Write-Host ($desktopWindow.ProcessName + ' ' + $desktopWindow.windowVisualState + ' ' + $desktopWindow.Left + ' ' + $desktopWindow.Right + ' ' + $desktopWindow.Top + ' ' + $desktopWindow.Bottom)
}

<#

leftovers

# $monitor = Get-Wmiobject Win32_Videocontroller
# $monitor.CurrentHorizontalResolution
# $monitor.CurrentVerticalResolution

# [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
# [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
# $res=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds
# $Forms = New-Object System.Windows.Forms.Form
# $Forms.Size = New-Object System.Drawing.Size($res.Width,$res.Height)

#get the processes
# $ProcessList = Get-Process  

#loop over them checking for chrome
#($Process.MainWindowTitle -like '*code*')

Get-Process | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object Name, MainWindowTitle

[System.ComponentModel.Component] $process = Get-Process | Where-Object { $_.Name -match "code" } | Select-Object -First 1 
$process.ProcessName
Set-Window -ProcessName:($process.ProcessName) -X 10 -y 10 -Width 500 -Height 500 -Passthru


#>
