# Prints {"<registry font name>": "<absolute file path>"} for every installed font (machine + user).
$map = [ordered]@{}
$sys = Join-Path $env:SystemRoot 'Fonts'
foreach ($key in 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts', 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts') {
    if (-not (Test-Path $key)) { continue }
    $props = Get-ItemProperty $key
    foreach ($p in $props.PSObject.Properties) {
        if ($p.Name -like 'PS*') { continue }
        $file = [string]$p.Value
        if (-not [System.IO.Path]::IsPathRooted($file)) { $file = Join-Path $sys $file }
        $map[$p.Name] = $file
    }
}
$map | ConvertTo-Json -Compress
