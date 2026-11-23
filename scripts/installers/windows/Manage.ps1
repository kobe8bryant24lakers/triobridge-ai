param([ValidateSet('Install', 'Remove')][string]$Action = 'Install')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'windows-office-discovery.ps1')

function Assert-OwnManifest([string]$Path) {
    if (Test-Path -LiteralPath $Path) {
        if ((Get-Item -LiteralPath $Path).Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw 'Refusing to modify a linked manifest.'
        }
        $text = [IO.File]::ReadAllText($Path)
        if (-not $text.Contains('<Id>__HOSTED_MANIFEST_ID__</Id>')) {
            throw 'An unrelated manifest already uses this filename. No changes were made.'
        }
    }
}

function Invoke-ShareSetup([string]$ShareAction) {
    # Only this helper is elevated. HKCU below always belongs to the original user,
    # including when a different administrator supplies the UAC credentials.
    $arguments = @($helper, $ShareAction, $directory, $account, $shareName) | ForEach-Object {
        "'" + $_.Replace("'", "''") + "'"
    }
    $command = "try { & $($arguments[0]) -Action $($arguments[1]) -Directory $($arguments[2]) -Account $($arguments[3]) -Name $($arguments[4]); exit 0 } catch { Write-Error `$_; exit 1 }"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $process = Start-Process -FilePath "$PSHOME\powershell.exe" -Verb RunAs -Wait -PassThru `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $encoded)
    if ($process.ExitCode -ne 0) {
        throw 'Could not configure the private catalog share. Check administrator permissions and the Windows Server service, then retry. No Office trust was added.'
    }
}

try {
    Assert-TrioBridgeOfficeClosed
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $account = $identity.Name
    $shareName = 'TrioBridgeAI-' + $identity.User.Value + '$'
    $directory = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'TrioBridgeAI\OnlineCatalog'
    $manifestPath = Join-Path $directory 'triobridge-ai-online.xml'
    $source = Join-Path $PSScriptRoot 'manifest.xml'
    $helper = Join-Path $PSScriptRoot 'Share.ps1'
    $catalogId = '{__HOSTED_MANIFEST_ID__}'
    $registryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
    $catalogUrl = "\\$env:COMPUTERNAME\$shareName"

    if (Test-Path -LiteralPath $registryPath) {
        $registered = Get-ItemProperty -LiteralPath $registryPath
        if ($registered.Url -ne $catalogUrl -or $registered.Id -ne $catalogId) {
            throw 'An unrelated trusted catalog uses this ID. No changes were made.'
        }
    }
    foreach ($path in @((Split-Path -Parent $directory), $directory)) {
        if ((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'Refusing to modify a linked catalog folder.'
        }
    }
    Assert-OwnManifest $manifestPath
    if ($Action -eq 'Remove') {
        Invoke-ShareSetup 'Remove'
        if (Test-Path -LiteralPath $registryPath) { Remove-Item -LiteralPath $registryPath }
        if (Test-Path -LiteralPath $manifestPath) { Remove-Item -LiteralPath $manifestPath }
        Write-Host 'Removed TrioBridge installation files and its catalog. Model settings and keys are unchanged.'
        Write-Host 'Restart Office. If a cached entry remains, follow the Microsoft Office cache instructions.'
    } else {
        if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne '__MANIFEST_SHA256__') {
            throw 'The bundled manifest is missing or damaged. Download the installer again.'
        }
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
        # Restrict the dedicated folder to its owner, SYSTEM and administrators.
        $acl = New-Object Security.AccessControl.DirectorySecurity
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($sid in @($identity.User.Value, 'S-1-5-18', 'S-1-5-32-544')) {
            $principal = New-Object Security.Principal.SecurityIdentifier($sid)
            $rule = New-Object Security.AccessControl.FileSystemAccessRule($principal, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
            $acl.AddAccessRule($rule)
        }
        Set-Acl -LiteralPath $directory -AclObject $acl
        Copy-Item -LiteralPath $source -Destination $manifestPath -Force
        Write-Host 'Windows will request permission to create a catalog share accessible only to your account.'
        Invoke-ShareSetup 'Install'
        if (-not (Test-Path -LiteralPath (Join-Path $catalogUrl 'triobridge-ai-online.xml'))) {
            throw 'The catalog cannot be reached. Check Windows file sharing or ask your administrator. No Office trust was added.'
        }
        New-Item -Path $registryPath -Force | Out-Null
        New-ItemProperty -LiteralPath $registryPath -Name Id -Value $catalogId -PropertyType String -Force | Out-Null
        New-ItemProperty -LiteralPath $registryPath -Name Url -Value $catalogUrl -PropertyType String -Force | Out-Null
        New-ItemProperty -LiteralPath $registryPath -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null
        Write-Host 'Installed. Reopen Word, Excel and PowerPoint; add the entry separately in each application.'
        Write-Host 'Go to Home > Add-ins > More Add-ins / Advanced > Shared Folder, then add TrioBridge AI (Online).'
        Write-Host 'In Model settings, choose a provider, paste your API key, and connect.'
    }
    exit 0
} catch {
    Write-Host ("Installation action failed: " + $_.Exception.Message) -ForegroundColor Red
    Write-Host 'The setup may be incomplete. Retry, or run Uninstall.cmd to remove its files and catalog.'
    Write-Host 'Help: https://kobe8bryant24lakers.github.io/triobridge-ai/#install'
    exit 1
}
