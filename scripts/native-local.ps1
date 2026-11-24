param(
    [Parameter(Mandatory = $true)][ValidateSet('Preflight', 'Certificate', 'Trust', 'RemoveCertificate', 'Start', 'Stop', 'TaskStatus', 'Catalog', 'RemoveCatalog', 'RemoveService', 'RegisterUninstaller')][string]$Action,
    [Parameter(Mandatory = $true)][string]$Root,
    [string]$CertificateDirectory,
    [string]$InstallationId,
    [string]$ManifestPath,
    [ValidateSet(0, 1)][int]$RequireClosedOffice = 0
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
trap {
    if ($Action -in @('Start', 'Stop')) {
        [Console]::Error.WriteLine(('TRIOBRIDGE_TASK_ERROR 0x{0:X8}' -f $_.Exception.HResult))
    }
    break
}
. (Join-Path $PSScriptRoot 'windows-office-discovery.ps1')
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$sid = $identity.User.Value
$account = $identity.Name
$expectedRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'TrioBridgeAI\LocalApp'
if ([IO.Path]::GetFullPath($Root) -ne $expectedRoot) { throw 'Unexpected local installation directory.' }
$state = [IO.File]::ReadAllText((Join-Path $Root 'installation.json')) | ConvertFrom-Json
if ($state.product -ne 'triobridge-ai-local') { throw 'The directory is not a TrioBridge installation.' }
$taskName = 'TrioBridgeAI-Local-' + $sid
$taskDescription = 'TrioBridge AI local shared service'
$catalog = Join-Path (Split-Path -Parent $Root) 'LocalCatalog'
$shareName = 'TrioBridgeAI-Local-' + $sid + '$'
$catalogUrl = "\\$env:COMPUTERNAME\$shareName"
$catalogId = '{d7cf6e8b-6d03-4a54-9c31-5b7d5427618e}'
$catalogRegistry = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
$uninstallRegistry = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\TrioBridgeAI.Local'

function Assert-OrdinaryPath([string]$Path) {
    if (Test-Path -LiteralPath $Path) {
        $item = Get-Item -LiteralPath $Path -Force
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked installation paths are not supported.' }
        if ((Get-Acl -LiteralPath $Path).GetOwner([Security.Principal.SecurityIdentifier]).Value -ne $sid) {
            throw 'Installation files must belong to the current account.'
        }
    }
}

function Protect-Directory([string]$Path) {
    Assert-OrdinaryPath $Path
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
    $acl = New-Object Security.AccessControl.DirectorySecurity
    $acl.SetOwner($identity.User)
    $acl.SetAccessRuleProtection($true, $false)
    foreach ($value in @($sid, 'S-1-5-18', 'S-1-5-32-544')) {
        $principal = New-Object Security.Principal.SecurityIdentifier($value)
        $rule = New-Object Security.AccessControl.FileSystemAccessRule($principal, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $acl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $Path -AclObject $acl
}

function Protect-NewManifest([string]$Path) {
    $item = Get-Item -LiteralPath $Path -Force
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked manifests are not supported.' }
    $acl = New-Object Security.AccessControl.FileSecurity
    $acl.SetOwner($identity.User)
    $acl.SetAccessRuleProtection($true, $false)
    foreach ($value in @($sid, 'S-1-5-18', 'S-1-5-32-544')) {
        $principal = New-Object Security.Principal.SecurityIdentifier($value)
        $rule = New-Object Security.AccessControl.FileSystemAccessRule($principal, 'FullControl', 'Allow')
        $acl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $Path -AclObject $acl
}

function Get-OwnedTask {
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($task) {
        $prefix = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + (Join-Path $Root 'versions\')
        $argumentsPattern = '^' + [Regex]::Escape($prefix) + '\d+\.\d+\.\d+-[a-f0-9]{12}(-[a-f0-9]{8})?' + [Regex]::Escape('\app\scripts\start-installed.ps1"') + '$'
        if ($task.Description -ne $taskDescription -or @($task.Actions).Count -ne 1 -or
            $task.Actions[0].Execute -ne (Join-Path $PSHOME 'powershell.exe') -or
            $task.Actions[0].Arguments -notmatch $argumentsPattern) {
            throw 'An unrelated scheduled task uses the TrioBridge name. No changes were made.'
        }
    }
    return $task
}

function Assert-Catalog {
    Assert-OrdinaryPath $catalog
    $manifest = Join-Path $catalog 'triobridge-ai.xml'
    Assert-OrdinaryPath $manifest
    if ((Test-Path -LiteralPath $manifest) -and -not ([IO.File]::ReadAllText($manifest)).Contains('<Id>d7cf6e8b-6d03-4a54-9c31-5b7d5427618e</Id>')) {
        throw 'Another add-in owns the local manifest filename.'
    }
    if (Test-Path -LiteralPath $catalogRegistry) {
        $entry = Get-ItemProperty -LiteralPath $catalogRegistry
        if ($entry.Id -ne $catalogId -or $entry.Url -ne $catalogUrl) { throw 'Another catalog owns the local registration ID.' }
    }
}

function Stop-OwnedServiceProcess {
    $recordPath = Join-Path $Root 'service-process.json'
    if (-not (Test-Path -LiteralPath $recordPath)) { return }
    Assert-OrdinaryPath $recordPath
    $record = [IO.File]::ReadAllText($recordPath) | ConvertFrom-Json
    if ($record.installationId -ne $state.id -or $record.version -notmatch '^\d+\.\d+\.\d+-[a-f0-9]{12}(-[a-f0-9]{8})?$' -or
        $record.pid -notmatch '^\d+$' -or [long]$record.pid -le 0 -or [long]$record.pid -gt [int]::MaxValue) {
        throw 'Invalid service process record; no process was stopped.'
    }
    $version = Join-Path $Root ('versions\' + $record.version)
    $executable = Join-Path $version 'runtime\node.exe'
    $entry = Join-Path $version 'app\scripts\serve-installed.mjs'
    $target = Get-Process -Id ([int]$record.pid) -ErrorAction SilentlyContinue
    if (-not $target) { return }
    try {
        # Hold the process handle so a recycled PID cannot select another process.
        $null = $target.Handle
        if ($target.MainModule.FileName -ne $executable) { return }
        $processInfo = Get-CimInstance Win32_Process -Filter ('ProcessId = ' + [int]$record.pid)
        $commandPattern = '^"?' + [Regex]::Escape($executable) + '"?\s+"' + [Regex]::Escape($entry) + '"\s*$'
        if (-not $processInfo -or $processInfo.CommandLine -notmatch $commandPattern) { return }
        $target.Kill()
        if (-not $target.WaitForExit(15000)) { throw 'The owned local service did not stop.' }
    } finally { $target.Dispose() }
}

function Invoke-Share([string]$Operation) {
    $helper = Join-Path $PSScriptRoot 'local-catalog-share.ps1'
    $quoted = @($helper, $Operation, $catalog, $account, $shareName) | ForEach-Object { "'" + $_.Replace("'", "''") + "'" }
    $command = "try { & $($quoted[0]) -Action $($quoted[1]) -Directory $($quoted[2]) -Account $($quoted[3]) -Name $($quoted[4]); exit 0 } catch { Write-Error `$_; exit 1 }"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $process = Start-Process -FilePath "$PSHOME\powershell.exe" -Verb RunAs -Wait -PassThru -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $encoded)
    if ($process.ExitCode -ne 0) { throw 'The private catalog share was not configured. Approve the Windows administrator prompt and verify the Server service is available.' }
}

function Get-OwnedCertificate {
    if (-not $CertificateDirectory -or -not ([IO.Path]::GetFullPath($CertificateDirectory)).StartsWith($Root + '\', [StringComparison]::OrdinalIgnoreCase) -or $InstallationId -ne $state.id) {
        throw 'Invalid certificate directory or installation identity.'
    }
    Assert-OrdinaryPath $CertificateDirectory
    $record = [IO.File]::ReadAllText((Join-Path $CertificateDirectory 'identity.json')) | ConvertFrom-Json
    $certificate = New-Object Security.Cryptography.X509Certificates.X509Certificate2((Join-Path $CertificateDirectory 'server.cer'))
    if ($record.installationId -ne $InstallationId -or $record.thumbprint -ne $certificate.Thumbprint) { throw 'Certificate ownership mismatch.' }
    return $certificate
}

switch ($Action) {
    'TaskStatus' {
        Get-TrioBridgeTaskStatus | ConvertTo-Json -Compress
    }
    'Preflight' {
        if ($RequireClosedOffice) { Assert-TrioBridgeOfficeClosed }
        Assert-OrdinaryPath $Root
        Assert-Catalog
        $null = Get-OwnedTask
    }
    'Certificate' {
        if (-not $CertificateDirectory -or -not ([IO.Path]::GetFullPath($CertificateDirectory)).StartsWith($Root + '\', [StringComparison]::OrdinalIgnoreCase) -or $InstallationId -ne $state.id) { throw 'Invalid certificate request.' }
        Protect-Directory $CertificateDirectory
        $certificate = New-SelfSignedCertificate -Type Custom -Subject "CN=localhost, O=TrioBridge AI, OU=$InstallationId" `
            -FriendlyName "TrioBridge AI local $InstallationId" -CertStoreLocation 'Cert:\CurrentUser\My' `
            -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable `
            -KeyUsage DigitalSignature, KeyEncipherment -NotAfter (Get-Date).AddDays(365) `
            -TextExtension @('2.5.29.17={text}DNS=localhost&IPAddress=127.0.0.1&IPAddress=::1', '2.5.29.37={text}1.3.6.1.5.5.7.3.1', '2.5.29.19={text}CA=false')
        [IO.File]::WriteAllBytes((Join-Path $CertificateDirectory 'server.cer'), $certificate.RawData)
        $utf8 = New-Object Text.UTF8Encoding($false)
        $record = @{ installationId = $InstallationId; thumbprint = $certificate.Thumbprint } | ConvertTo-Json -Compress
        [IO.File]::WriteAllText((Join-Path $CertificateDirectory 'identity.json'), $record, $utf8)
        $pem = "-----BEGIN CERTIFICATE-----`r`n" + [Convert]::ToBase64String($certificate.RawData, [Base64FormattingOptions]::InsertLineBreaks) + "`r`n-----END CERTIFICATE-----`r`n"
        [IO.File]::WriteAllText((Join-Path $CertificateDirectory 'server.crt'), $pem, $utf8)
        $random = New-Object byte[] 32
        $generator = [Security.Cryptography.RandomNumberGenerator]::Create()
        try { $generator.GetBytes($random) } finally { $generator.Dispose() }
        $password = [Convert]::ToBase64String($random)
        [IO.File]::WriteAllText((Join-Path $CertificateDirectory 'server.password'), $password, $utf8)
        $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
        Export-PfxCertificate -Cert $certificate -FilePath (Join-Path $CertificateDirectory 'server.pfx') -Password $securePassword -CryptoAlgorithmOption AES256_SHA256 -ChainOption EndEntityCertOnly -NoProperties | Out-Null
    }
    'Trust' {
        $certificate = Get-OwnedCertificate
        if (Test-Path -LiteralPath ('Cert:\CurrentUser\Root\' + $certificate.Thumbprint)) { break }
        $store = New-Object Security.Cryptography.X509Certificates.X509Store('Root', 'CurrentUser')
        try { $store.Open('ReadWrite'); $store.Add($certificate) } finally { $store.Close() }
        if (-not (Test-Path -LiteralPath ('Cert:\CurrentUser\Root\' + $certificate.Thumbprint))) { throw 'The localhost certificate was not trusted for the current user.' }
    }
    'RemoveCertificate' {
        $certificate = Get-OwnedCertificate
        foreach ($store in @('Root', 'My')) {
            $path = "Cert:\CurrentUser\$store\" + $certificate.Thumbprint
            if (Test-Path -LiteralPath $path) {
                if ($store -eq 'My') { Remove-Item -LiteralPath $path -DeleteKey }
                else { Remove-Item -LiteralPath $path }
            }
        }
    }
    'Stop' {
        if (Get-OwnedTask) { Stop-ScheduledTask -TaskName $taskName }
        Stop-OwnedServiceProcess
    }
    'Start' {
        $null = Get-OwnedTask
        if ($state.active -notmatch '^\d+\.\d+\.\d+-[a-f0-9]{12}(-[a-f0-9]{8})?$') { throw 'Invalid active release.' }
        $version = Join-Path $Root ('versions\' + $state.active)
        $executable = Join-Path $PSHOME 'powershell.exe'
        $entry = Join-Path $version 'app\scripts\start-installed.ps1'
        $actionDefinition = New-ScheduledTaskAction -Execute $executable -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $entry + '"') -WorkingDirectory (Join-Path $version 'app')
        $trigger = New-ScheduledTaskTrigger -AtLogOn -User $account
        $principal = New-ScheduledTaskPrincipal -UserId $sid -LogonType Interactive -RunLevel Limited
        $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
        Register-ScheduledTask -TaskName $taskName -Description $taskDescription -Action $actionDefinition -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName
    }
    'Catalog' {
        Assert-TrioBridgeOfficeClosed
        Assert-Catalog
        if (-not $ManifestPath -or -not ([IO.File]::ReadAllText($ManifestPath)).Contains('<Id>d7cf6e8b-6d03-4a54-9c31-5b7d5427618e</Id>')) { throw 'Missing generated private manifest.' }
        Protect-Directory $catalog
        Copy-Item -LiteralPath $ManifestPath -Destination (Join-Path $catalog 'triobridge-ai.xml') -Force
        Protect-NewManifest (Join-Path $catalog 'triobridge-ai.xml')
        Invoke-Share 'Install'
        if (-not (Test-Path -LiteralPath (Join-Path $catalogUrl 'triobridge-ai.xml'))) { throw 'The local catalog share is not reachable.' }
        New-Item -Path $catalogRegistry -Force | Out-Null
        New-ItemProperty -LiteralPath $catalogRegistry -Name Id -Value $catalogId -PropertyType String -Force | Out-Null
        New-ItemProperty -LiteralPath $catalogRegistry -Name Url -Value $catalogUrl -PropertyType String -Force | Out-Null
        New-ItemProperty -LiteralPath $catalogRegistry -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null
    }
    'RemoveCatalog' {
        Assert-TrioBridgeOfficeClosed
        Assert-Catalog
        Invoke-Share 'Remove'
        if (Test-Path -LiteralPath $catalogRegistry) { Remove-Item -LiteralPath $catalogRegistry }
        $manifest = Join-Path $catalog 'triobridge-ai.xml'
        if (Test-Path -LiteralPath $manifest) { Remove-Item -LiteralPath $manifest }
    }
    'RemoveService' {
        if (Get-OwnedTask) { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false }
        if (Test-Path -LiteralPath $uninstallRegistry) {
            $registration = Get-ItemProperty -LiteralPath $uninstallRegistry
            if ($registration.InstallLocation -ne $Root) { throw 'Uninstall registration belongs to another installation.' }
            Remove-Item -LiteralPath $uninstallRegistry
        }
    }
    'RegisterUninstaller' {
        if (Test-Path -LiteralPath $uninstallRegistry) {
            if ((Get-ItemProperty -LiteralPath $uninstallRegistry).InstallLocation -ne $Root) { throw 'Uninstall registration belongs to another installation.' }
        }
        $version = Join-Path $Root ('versions\' + $state.active)
        $metadata = [IO.File]::ReadAllText((Join-Path $version 'app\package.json')) | ConvertFrom-Json
        $wrapper = '@"%~dp0versions\' + $state.active + '\Uninstall.cmd"' + "`r`n"
        [IO.File]::WriteAllText((Join-Path $Root 'Uninstall.cmd'), $wrapper, [Text.Encoding]::ASCII)
        New-Item -Path $uninstallRegistry -Force | Out-Null
        foreach ($item in @{ DisplayName = 'TrioBridge AI (Local Shared)'; DisplayVersion = $metadata.version; Publisher = 'TrioBridge AI'; InstallLocation = $Root; UninstallString = ('"' + (Join-Path $Root 'Uninstall.cmd') + '"') }.GetEnumerator()) {
            New-ItemProperty -LiteralPath $uninstallRegistry -Name $item.Key -Value $item.Value -PropertyType String -Force | Out-Null
        }
        New-ItemProperty -LiteralPath $uninstallRegistry -Name NoModify -Value 1 -PropertyType DWord -Force | Out-Null
        New-ItemProperty -LiteralPath $uninstallRegistry -Name NoRepair -Value 1 -PropertyType DWord -Force | Out-Null
    }
}
