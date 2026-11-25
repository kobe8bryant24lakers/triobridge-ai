param(
    [switch]$Verify,
    [string]$ExpectedSid,
    [string]$Node,
    [string]$WinRAR
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

if ($Verify) {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    if (($identity.Name -split '\\')[-1] -notmatch '^tb[0-9a-f]{12}$' -or
        $identity.User.Value -ne $ExpectedSid -or
        $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'The installer smoke test must run as the disposable standard user.'
    }
    # Start-Process can inherit paths from the elevated Actions worker. Use the
    # folders owned by the newly loaded profile for WebView2, DPAPI and TEMP.
    $env:USERPROFILE = [Environment]::GetFolderPath('UserProfile')
    $env:LOCALAPPDATA = [Environment]::GetFolderPath('LocalApplicationData')
    $env:APPDATA = [Environment]::GetFolderPath('ApplicationData')
    $env:TEMP = $env:TMP = Join-Path $env:LOCALAPPDATA 'Temp'
    [void][IO.Directory]::CreateDirectory($env:TEMP)
    $env:TRIOBRIDGE_TEST_WINRAR = $WinRAR
    $env:RUNNER_ENVIRONMENT = 'github-hosted'
    Set-Location -LiteralPath $root
    & $Node offline/verify-install.mjs
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    & $Node offline/verify-ui.mjs --test-model
    exit $LASTEXITCODE
}

if ($env:RUNNER_ENVIRONMENT -ne 'github-hosted') {
    throw 'Use an ephemeral hosted Windows runner for this acceptance test.'
}
$Node = (Get-Command node.exe).Source
$arch = & $Node -p 'process.arch'
$build = Join-Path $root "build/native-win32-$arch"
$name = 'tb' + [Guid]::NewGuid().ToString('N').Substring(0, 12)
$password = ConvertTo-SecureString ('Tb1!' + [Guid]::NewGuid().ToString('N')) -AsPlainText -Force
$account = New-LocalUser -Name $name -Password $password -AccountNeverExpires
try {
    $sid = $account.SID.Value
    & icacls.exe $root /grant "*${sid}:(OI)(CI)RX" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Could not grant access to the checkout.' }
    & icacls.exe $build /grant "*${sid}:(OI)(CI)M" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Could not grant access to the extracted installer and smoke tests.' }
    $arguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                   '-File', "`"$PSCommandPath`"", '-Verify', '-ExpectedSid', $sid, '-Node', "`"$Node`"")
    if ($env:TRIOBRIDGE_TEST_WINRAR) {
        $arguments += @('-WinRAR', "`"$env:TRIOBRIDGE_TEST_WINRAR`"")
    }
    $output = Join-Path $build "$name-out.txt"
    $errors = Join-Path $build "$name-err.txt"
    $options = @{
        FilePath = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
        ArgumentList = $arguments
        Credential = [PSCredential]::new("$env:COMPUTERNAME\$name", $password)
        LoadUserProfile = $true
        WorkingDirectory = $root
        RedirectStandardOutput = $output
        RedirectStandardError = $errors
        Wait = $true
        PassThru = $true
    }
    $result = Start-Process @options
    Get-Content -LiteralPath $output
    Get-Content -LiteralPath $errors
    if ($result.ExitCode -ne 0) { throw "Standard-user installer acceptance exited with $($result.ExitCode)." }
}
finally {
    Remove-LocalUser -SID $account.SID
}
