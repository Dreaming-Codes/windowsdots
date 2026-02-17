Add-Type -AssemblyName System.Windows.Forms
$primary = [System.Windows.Forms.Screen]::AllScreens | Where-Object { $_.Primary }
$displayNum = ($primary.DeviceName -replace '.*DISPLAY','').PadLeft(4,'0')
Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Enum\DISPLAY' -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
    if ($props.PSObject.Properties.Name -contains 'Driver' -and $props.Driver -like "*\$displayNum") {
        $parts = $_.Name -split '\\'
        $hw = $parts[$parts.Count-2]
        $inst = $parts[$parts.Count-1]
        if ($hw -ne 'Default_Monitor' -and $hw -notlike 'MS_*') {
            Write-Output "$hw-$inst"
        }
    }
}
