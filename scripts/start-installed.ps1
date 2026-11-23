$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$package = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$start = New-Object Diagnostics.ProcessStartInfo
$start.FileName = Join-Path $package 'runtime\node.exe'
$start.Arguments = '"' + (Join-Path $PSScriptRoot 'serve-installed.mjs') + '"'
$start.WorkingDirectory = Split-Path -Parent $PSScriptRoot
$start.UseShellExecute = $false
$start.CreateNoWindow = $true
$process = New-Object Diagnostics.Process
$process.StartInfo = $start
try {
    if (-not $process.Start()) { throw 'The local service process could not start.' }
    $process.WaitForExit()
    $result = $process.ExitCode
} finally { $process.Dispose() }
exit $result
