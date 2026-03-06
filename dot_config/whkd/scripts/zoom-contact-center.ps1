param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('call-selected', 'pickup', 'hangup', 'finish-wrap-up')]
    [string]$Action
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeWin32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT {
        public uint type;
        public MOUSEINPUT mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    public static void LeftClick() {
        INPUT[] inputs = new INPUT[2];
        inputs[0].type = 0;
        inputs[0].mi.dwFlags = 0x0002;
        inputs[1].type = 0;
        inputs[1].mi.dwFlags = 0x0004;
        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    public static void VerticalWheel(int delta) {
        INPUT[] inputs = new INPUT[1];
        inputs[0].type = 0;
        inputs[0].mi.dwFlags = 0x0800;
        inputs[0].mi.mouseData = unchecked((uint)delta);
        SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
    }
}
"@

function Send-Shortcut {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Keys
    )

    $shell = New-Object -ComObject WScript.Shell
    $shell.SendKeys($Keys)
}

function Get-SelectedPhoneNumber {
    Add-Type -AssemblyName System.Windows.Forms

    $shell = New-Object -ComObject WScript.Shell
    $originalClipboard = $null
    $hadClipboardText = $false

    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        $originalClipboard = [System.Windows.Forms.Clipboard]::GetText()
        $hadClipboardText = $true
    }

    $selectedText = ''
    foreach ($attempt in 1..3) {
        $shell.SendKeys('^c')

        for ($i = 0; $i -lt 10; $i++) {
            Start-Sleep -Milliseconds 100

            if (-not [System.Windows.Forms.Clipboard]::ContainsText()) {
                continue
            }

            $clipboardText = [System.Windows.Forms.Clipboard]::GetText()
            if ([string]::IsNullOrWhiteSpace($clipboardText)) {
                continue
            }

            if ($hadClipboardText -and $clipboardText -eq $originalClipboard) {
                continue
            }

            $selectedText = $clipboardText
            break
        }

        if ($selectedText) {
            break
        }
    }

    $trimmed = $selectedText.Trim()
    if (-not $trimmed) {
        throw 'No new selected phone number was copied. Select the number first, or copy it immediately before retrying.'
    }

    $normalized = $trimmed -replace '[^\d+]', ''
    if ($normalized.StartsWith('+')) {
        $normalized = '+' + ($normalized.Substring(1) -replace '[^\d]', '')
    }
    else {
        $normalized = $normalized -replace '[^\d]', ''
    }

    if ($normalized -notmatch '^\+?\d{7,}$') {
        throw "Selected text does not look like a phone number: $trimmed"
    }

    if ($hadClipboardText) {
        [System.Windows.Forms.Clipboard]::SetText($originalClipboard)
    }

    return $normalized
}

function Get-ZoomWindowProcess {
    $candidates = Get-Process |
        Where-Object {
            $_.MainWindowHandle -ne 0 -and (
                $_.ProcessName -match '^zoom' -or
                $_.MainWindowTitle -match 'zoom'
            )
        }

    $preferred = $candidates |
        Where-Object {
            $_.MainWindowTitle -eq 'Zoom Workplace'
        } |
        Sort-Object StartTime -Descending

    if ($preferred) {
        return $preferred | Select-Object -First 1
    }

    return $candidates | Sort-Object StartTime -Descending | Select-Object -First 1
}

function Focus-ZoomWindow {
    $zoom = Get-ZoomWindowProcess
    if (-not $zoom) {
        throw 'Zoom Workplace window not found.'
    }

    [NativeWin32]::ShowWindowAsync($zoom.MainWindowHandle, 9) | Out-Null
    Start-Sleep -Milliseconds 80
    [NativeWin32]::SetForegroundWindow($zoom.MainWindowHandle) | Out-Null
    Start-Sleep -Milliseconds 120

    return $zoom
}

function Invoke-ZoomCallButton {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

    $deadline = (Get-Date).AddSeconds(10)
    do {
        $zoom = Focus-ZoomWindow
        if ($zoom) {
            $root = [System.Windows.Automation.AutomationElement]::FromHandle($zoom.MainWindowHandle)
            if ($root) {
                $inputCondition = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::AutomationIdProperty,
                    'contactSearchInput'
                )
                $buttonCondition = New-Object System.Windows.Automation.AndCondition(
                    (New-Object System.Windows.Automation.PropertyCondition(
                        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                        [System.Windows.Automation.ControlType]::Button
                    )),
                    (New-Object System.Windows.Automation.PropertyCondition(
                        [System.Windows.Automation.AutomationElement]::NameProperty,
                        'Call'
                    ))
                )

                $searchInput = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $inputCondition)
                if ($searchInput -and -not $searchInput.Current.IsOffscreen) {
                    $searchInput.SetFocus()
                    Start-Sleep -Milliseconds 80
                    Send-Shortcut '{ENTER}'
                    return
                }

                $callButton = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $buttonCondition)
                if ($callButton -and $callButton.Current.IsEnabled -and -not $callButton.Current.IsOffscreen) {
                    $pattern = $null
                    if ($callButton.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) {
                        $pattern.Invoke()
                        return
                    }

                    $bounds = $callButton.Current.BoundingRectangle
                    if ($bounds.Width -gt 0 -and $bounds.Height -gt 0) {
                        $x = [int]($bounds.X + ($bounds.Width / 2))
                        $y = [int]($bounds.Y + ($bounds.Height / 2))
                        [NativeWin32]::SetCursorPos($x, $y) | Out-Null
                        Start-Sleep -Milliseconds 100
                        [NativeWin32]::LeftClick()
                        return
                    }

                    Start-Sleep -Milliseconds 100
                    Send-Shortcut ' '
                    return
                }
            }
        }

        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)

    throw 'Zoom opened the number, but the call button could not be triggered automatically.'
}

function Click-ZoomButtonByName {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Names,
        [int]$TimeoutSeconds = 5,
        [switch]$SkipFocus
    )

    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $zoom = if ($SkipFocus) { Get-ZoomWindowProcess } else { Focus-ZoomWindow }
        if ($zoom) {
            $root = [System.Windows.Automation.AutomationElement]::FromHandle($zoom.MainWindowHandle)
            if ($root) {
                $buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                    [System.Windows.Automation.ControlType]::Button
                )

                $buttons = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $buttonCondition)
                for ($i = 0; $i -lt $buttons.Count; $i++) {
                    $button = $buttons.Item($i)
                    $name = $button.Current.Name
                    if ([string]::IsNullOrWhiteSpace($name)) {
                        continue
                    }

                    if ($Names -notcontains $name) {
                        continue
                    }

                    if (-not $button.Current.IsEnabled -or $button.Current.IsOffscreen) {
                        continue
                    }

                    $pattern = $null
                    if ($button.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) {
                        $pattern.Invoke()
                        return $true
                    }

                    $bounds = $button.Current.BoundingRectangle
                    if ($bounds.Width -gt 0 -and $bounds.Height -gt 0) {
                        $x = [int]($bounds.X + ($bounds.Width / 2))
                        $y = [int]($bounds.Y + ($bounds.Height / 2))
                        [NativeWin32]::SetCursorPos($x, $y) | Out-Null
                        Start-Sleep -Milliseconds 40
                        [NativeWin32]::LeftClick()
                        return $true
                    }
                }
            }
        }

        Start-Sleep -Milliseconds 100
    } while ((Get-Date) -lt $deadline)

    return $false
}

function Scroll-WrapUpPaneRight {
    param(
        [int]$Steps = 10
    )

    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

    $zoom = Focus-ZoomWindow
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($zoom.MainWindowHandle)

    $paneCondition = New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Group
        )),
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ClassNameProperty,
            'content-wrap phone-tab'
        ))
    )

    $pane = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $paneCondition)
    if (-not $pane) {
        throw 'Could not find the wrap-up pane in Zoom.'
    }

    $bounds = $pane.Current.BoundingRectangle
    $x = [int]($bounds.X + ($bounds.Width / 2))
    $y = [int]($bounds.Y + ($bounds.Height / 2))

    [NativeWin32]::SetCursorPos($x, $y) | Out-Null
    Start-Sleep -Milliseconds 30

    # Hold Shift so the wheel event pans horizontally in the web view.
    [NativeWin32]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero)
    try {
        foreach ($step in 1..$Steps) {
            [NativeWin32]::VerticalWheel(-240)
            Start-Sleep -Milliseconds 40
        }
    }
    finally {
        [NativeWin32]::keybd_event(0x10, 0, 0x0002, [UIntPtr]::Zero)
    }
}

function Finish-ZoomWrapUp {
    Focus-ZoomWindow | Out-Null

    # The wrap-up form lives in a horizontally pannable pane. Scroll it right
    # until the save action moves into view, then click it.
    foreach ($attempt in 1..5) {
        if (Click-ZoomButtonByName -Names @('Save and close wrap-up') -TimeoutSeconds 0 -SkipFocus) {
            return
        }

        Scroll-WrapUpPaneRight -Steps 10
        Start-Sleep -Milliseconds 80
    }

    throw 'Could not bring the Save and close wrap-up button into view.'
}

function Invoke-ZoomWrapUp {
    Finish-ZoomWrapUp
}

switch ($Action) {
    'call-selected' {
        $number = Get-SelectedPhoneNumber
        Start-Process "tel:$number"
        Start-Sleep -Milliseconds 500
        Invoke-ZoomCallButton
    }
    'pickup' { Send-Shortcut '^+a' }
    'hangup' {
        if (-not (Click-ZoomButtonByName -Names @('End') -TimeoutSeconds 2)) {
            throw 'Could not find the active-call End button in Zoom.'
        }
    }
    'finish-wrap-up' { Invoke-ZoomWrapUp }
}
