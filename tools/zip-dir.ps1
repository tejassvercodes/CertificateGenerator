# Zips a folder into an Office file: forward-slash entry names, no directory entries, and
# [Content_Types].xml first -- the layout Word insists on (bsdtar and Compress-Archive both
# produce zips Word calls "corrupted").
param([Parameter(Mandatory=$true)][string]$Dir, [Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path $Out) { Remove-Item $Out -Force }
$root = (Resolve-Path $Dir).Path.TrimEnd('\')
$files = Get-ChildItem -Path $root -Recurse -File | Sort-Object { if ($_.FullName.Substring($root.Length + 1) -eq '[Content_Types].xml') { 0 } else { 1 } }, FullName
$zip = [System.IO.Compression.ZipFile]::Open($Out, 'Create')
try {
    foreach ($f in $files) {
        $name = $f.FullName.Substring($root.Length + 1).Replace('\', '/')
        [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $f.FullName, $name, [System.IO.Compression.CompressionLevel]::Optimal)
    }
} finally { $zip.Dispose() }
