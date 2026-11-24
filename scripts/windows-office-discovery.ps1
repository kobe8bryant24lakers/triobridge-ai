param([switch]$Diagnose)

function Get-TrioBridgeRunningOfficeApps {
    $session = [Diagnostics.Process]::GetCurrentProcess().SessionId
    $hosts = [ordered]@{ Word = 'WINWORD'; Excel = 'EXCEL'; PowerPoint = 'POWERPNT' }
    foreach ($name in $hosts.Keys) {
        $running = @(Get-Process -Name $hosts[$name] -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $session })
        if ($running.Count) { $name }
    }
}

function Assert-TrioBridgeOfficeClosed {
    $running = @(Get-TrioBridgeRunningOfficeApps)
    if ($running.Count) {
        throw ('OFFICE_OPEN: 请先保存文档并完全退出 Word、Excel、PowerPoint，再重新运行安装或卸载。当前仍在运行：' + ($running -join '、') + '。未关闭任何文档或修改加载项目录。')
    }
}

function Get-TrioBridgeManifestStatus([string]$Path, [string]$ExpectedId) {
    $result = [ordered]@{ validXml = $false; matchesId = $false; Word = $false; Excel = $false; PowerPoint = $false }
    $reader = $null
    try {
        if ((Get-Item -LiteralPath $Path -ErrorAction Stop).Length -gt 1048576) { throw 'Manifest is too large.' }
        $settings = New-Object Xml.XmlReaderSettings
        $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $reader = [Xml.XmlReader]::Create($Path, $settings)
        $document = New-Object Xml.XmlDocument
        $document.XmlResolver = $null
        $document.Load($reader)
        $namespaces = New-Object Xml.XmlNamespaceManager($document.NameTable)
        $namespaces.AddNamespace('o', 'http://schemas.microsoft.com/office/appforoffice/1.1')
        $id = $document.SelectSingleNode('/o:OfficeApp/o:Id', $namespaces)
        $result.validXml = $null -ne $id
        $result.matchesId = $null -ne $id -and $id.InnerText -eq $ExpectedId
        $names = @($document.SelectNodes('/o:OfficeApp/o:Hosts/o:Host', $namespaces) | ForEach-Object { $_.GetAttribute('Name') })
        $result.Word = $names -contains 'Document'
        $result.Excel = $names -contains 'Workbook'
        $result.PowerPoint = $names -contains 'Presentation'
    } catch {
        # The report intentionally excludes raw XML, paths and exception contents.
    } finally { if ($reader) { $reader.Dispose() } }
    return [PSCustomObject]$result
}

function Get-TrioBridgeOfficeVersion([string]$Executable) {
    foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
        foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
            $base = $null
            $key = $null
            try {
                $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                $key = $base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\App Paths\' + $Executable)
                if (-not $key) { continue }
                $path = [string]$key.GetValue('')
                $path = $path.Trim('"')
                if ($path -match '^[A-Za-z]:\\') {
                    $version = [Diagnostics.FileVersionInfo]::GetVersionInfo($path).FileVersion
                    if ($version -match '^\d+\.\d+\.\d+\.\d+$') { return $version }
                }
            } catch {
                # Missing registrations do not prove that an Office app is absent.
            } finally {
                if ($key) { $key.Dispose() }
                if ($base) { $base.Dispose() }
            }
        }
    }
    return 'not detected'
}

function Get-TrioBridgeCatalogReport([string]$Edition) {
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    if ($Edition -eq 'local') {
        $id = 'd7cf6e8b-6d03-4a54-9c31-5b7d5427618e'
        $share = 'TrioBridgeAI-Local-' + $sid + '$'
        $filename = 'triobridge-ai.xml'
    } else {
        $id = '9d977989-12ea-46cf-920e-b54a12b4d728'
        $share = 'TrioBridgeAI-' + $sid + '$'
        $filename = 'triobridge-ai-online.xml'
    }
    $expected = "\\$env:COMPUTERNAME\$share"
    $path = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{' + $id + '}'
    $result = [ordered]@{ edition = $Edition; registered = $false; urlMatches = $false; showInMenu = $false; reachable = $false; manifest = $null }
    try {
        $entry = Get-ItemProperty -LiteralPath $path -ErrorAction Stop
        $result.registered = $entry.Id -eq ('{' + $id + '}')
        $result.urlMatches = $entry.Url -eq $expected
        $result.showInMenu = ([int]$entry.Flags -band 1) -eq 1
        # Never follow an arbitrary registry URL or contact a remote SMB server.
        if ($result.registered -and $result.urlMatches) {
            $manifest = Join-Path $expected $filename
            $result.reachable = Test-Path -LiteralPath $manifest -PathType Leaf
            if ($result.reachable) { $result.manifest = Get-TrioBridgeManifestStatus $manifest $id }
        }
    } catch {
        # Report only allowlisted booleans, never user names or registry contents.
    }
    return [PSCustomObject]$result
}

function Get-TrioBridgeTaskStatus {
    $result = [ordered]@{ state = 'Unavailable'; result = $null }
    try {
        $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        $name = 'TrioBridgeAI-Local-' + $sid
        $task = Get-ScheduledTask -TaskPath '\' -TaskName $name -ErrorAction Stop
        if ([string]$task.State -in @('Unknown', 'Disabled', 'Queued', 'Ready', 'Running')) {
            $result.state = [string]$task.State
        }
        $info = Get-ScheduledTaskInfo -TaskPath '\' -TaskName $name -ErrorAction Stop
        $result.result = '0x{0:X8}' -f ([long]$info.LastTaskResult -band 4294967295L)
    } catch {
        if ($_.CategoryInfo.Category -eq 'ObjectNotFound') { $result.state = 'Missing' }
    }
    return [PSCustomObject]$result
}

function Get-TrioBridgeLastStartup {
    $root = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'TrioBridgeAI\LocalApp'
    try {
        $item = Get-Item -LiteralPath (Join-Path $root 'last-startup.json') -Force -ErrorAction Stop
        if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $item.Length -gt 1024) { return }
        $record = [IO.File]::ReadAllText($item.FullName) | ConvertFrom-Json -ErrorAction Stop
        if ($record.phase -isnot [string] -or $record.code -isnot [string] -or
            $record.phase -cnotin @('stop', 'activate', 'credential', 'start', 'https', 'models', 'ready') -or
            $record.code -cnotin @('OK', 'TASK_STOP', 'TASK_START', 'ACTIVATION', 'CERTIFICATE_READ', 'CAPABILITY_ACCESS',
                'PORT_BUSY', 'PORT_DENIED', 'NO_LISTENER', 'TIMEOUT', 'CONNECTION_RESET', 'TLS_CERTIFICATE', 'SERVICE_IDENTITY',
                'LOCAL_AUTH', 'MODEL_STORAGE', 'INVALID_RESPONSE', 'TRANSPORT')) { return }
        $safe = [ordered]@{ phase = [string]$record.phase; code = [string]$record.code }
        if ($record.PSObject.Properties['hresult'] -and $record.hresult -is [string] -and $record.hresult -cmatch '^0x[0-9A-F]{8}\z') {
            $safe.hresult = [string]$record.hresult
        }
        if ($record.PSObject.Properties['task'] -and $record.task.state -is [string] -and $record.task.state -cin @('Unknown', 'Disabled', 'Queued', 'Ready', 'Running', 'Missing', 'Unavailable')) {
            $safe.taskState = [string]$record.task.state
            if ($record.task.result -is [string] -and $record.task.result -cmatch '^0x[0-9A-F]{8}\z') { $safe.taskResult = [string]$record.task.result }
        }
        return [PSCustomObject]$safe
    } catch { return }
}

function Get-TrioBridgeLocalServiceReport {
    $client = New-Object Net.Sockets.TcpClient
    $listening = $false
    try {
        $probe = $client.BeginConnect('127.0.0.1', 3000, $null, $null)
        if ($probe.AsyncWaitHandle.WaitOne(750)) { $client.EndConnect($probe); $listening = $true }
    } catch {
        # TCP reachability alone cannot identify the process or confirm TLS.
    } finally { $client.Close() }
    return [PSCustomObject][ordered]@{
        tcp3000 = $listening
        task = Get-TrioBridgeTaskStatus
        lastStartup = Get-TrioBridgeLastStartup
    }
}

if ($Diagnose) {
    $ErrorActionPreference = 'Stop'
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Windows is required.' }
    [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
    [ordered]@{
        format = 2
        runningOffice = @(Get-TrioBridgeRunningOfficeApps)
        detectedVersions = [ordered]@{
            Word = Get-TrioBridgeOfficeVersion 'WINWORD.EXE'
            Excel = Get-TrioBridgeOfficeVersion 'EXCEL.EXE'
            PowerPoint = Get-TrioBridgeOfficeVersion 'POWERPNT.EXE'
        }
        catalogs = @((Get-TrioBridgeCatalogReport 'local'), (Get-TrioBridgeCatalogReport 'online'))
        localService = Get-TrioBridgeLocalServiceReport
    } | ConvertTo-Json -Depth 5
}
