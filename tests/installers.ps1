# Portable syntax and helper tests. No registry writes, SMB shares or elevation.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path -Parent $PSScriptRoot
$templates = Join-Path $root 'scripts/installers/windows'
foreach ($file in @('Manage.ps1', 'Share.ps1')) {
    $tokens = $null
    $parseErrors = $null
    $null = [Management.Automation.Language.Parser]::ParseFile((Join-Path $templates $file), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
}

$testState = @{existingShare = $null; created = 0; removed = 0; grantedAccount = ''; shareAccess = @()}
function Import-Module { [CmdletBinding()]param([string]$Name) }
function Get-SmbShare {
    [CmdletBinding()]param([string]$Name)
    return $testState.existingShare
}
function New-SmbShare {
    [CmdletBinding()]param($Name, $Path, $FullAccess, $FolderEnumerationMode, $Description)
    $testState.existingShare = [PSCustomObject]@{Name = $Name; Path = $Path; Description = $Description}
    $testState.grantedAccount = $FullAccess
    $testState.shareAccess = @([PSCustomObject]@{AccountName = $FullAccess; AccessControlType = 'Allow'; AccessRight = 'Full'})
    $testState.created++
}
function Get-SmbShareAccess {
    [CmdletBinding()]param($Name)
    return $testState.shareAccess
}
function Remove-SmbShare {
    [CmdletBinding(SupportsShouldProcess = $true)]param($Name, [switch]$Force)
    $testState.existingShare = $null
    $testState.removed++
}
function Assert([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

$helper = Join-Path $templates 'Share.ps1'
$arguments = @{Directory = 'C:\Users\Test User\OnlineCatalog'; Account = 'PC\Test User'; Name = 'TrioBridgeAI-S-1-5-21-123$'}
& $helper -Action Install @arguments
& $helper -Action Install @arguments
Assert ($testState.created -eq 1) 'Repeated installation must reuse its own share.'
Assert ($testState.grantedAccount -eq 'PC\Test User') 'Only the installing account should receive share access.'
$testState.shareAccess += [PSCustomObject]@{AccountName = 'Everyone'; AccessControlType = 'Allow'; AccessRight = 'Read'}
$rejected = $false
try { & $helper -Action Install @arguments } catch { $rejected = $true }
Assert $rejected 'An existing share with broader access must not be reused.'
& $helper -Action Remove @arguments
& $helper -Action Remove @arguments
Assert ($testState.removed -eq 1) 'Repeated removal must be harmless.'

foreach ($conflict in @(
    [PSCustomObject]@{Path = 'C:\Unrelated'; Description = 'TrioBridge AI online catalog'},
    [PSCustomObject]@{Path = $arguments.Directory; Description = 'Another application'}
)) {
    foreach ($action in @('Install', 'Remove')) {
        $testState.existingShare = $conflict
        $rejected = $false
        try { & $helper -Action $action @arguments } catch { $rejected = $true }
        Assert $rejected 'An unrelated share must be rejected.'
        Assert ($testState.existingShare -eq $conflict) 'An unrelated share must remain untouched.'
    }
}
Write-Host 'PASS: PowerShell syntax, share idempotence, account scope, ACL drift and collision handling. Windows integration still needs a real host.'
