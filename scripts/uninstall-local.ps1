$ErrorActionPreference = 'Stop'
$package = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('triobridge-uninstall-' + [Guid]::NewGuid().ToString())
try {
    New-Item -ItemType Directory -Path $temporary | Out-Null
    Copy-Item -LiteralPath (Join-Path $package 'runtime\node.exe') -Destination (Join-Path $temporary 'node.exe')
    & (Join-Path $temporary 'node.exe') (Join-Path $PSScriptRoot 'install-local.mjs') remove
    $result = $LASTEXITCODE
} finally {
    Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue
}
exit $result
