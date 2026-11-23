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

if ($Diagnose) {
    $ErrorActionPreference = 'Stop'
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Windows is required.' }
    [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
    [ordered]@{
        format = 1
        runningOffice = @(Get-TrioBridgeRunningOfficeApps)
        detectedVersions = [ordered]@{
            Word = Get-TrioBridgeOfficeVersion 'WINWORD.EXE'
            Excel = Get-TrioBridgeOfficeVersion 'EXCEL.EXE'
            PowerPoint = Get-TrioBridgeOfficeVersion 'POWERPNT.EXE'
        }
        catalogs = @((Get-TrioBridgeCatalogReport 'local'), (Get-TrioBridgeCatalogReport 'online'))
    } | ConvertTo-Json -Depth 5
}
