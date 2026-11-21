param(
    [Parameter(Mandatory = $true)][ValidateSet('Install', 'Remove')][string]$Action,
    [Parameter(Mandatory = $true)][string]$Directory,
    [Parameter(Mandatory = $true)][string]$Account,
    [Parameter(Mandatory = $true)][ValidatePattern('^TrioBridgeAI-S-1-[0-9-]+\$$')][string]$Name
)
$ErrorActionPreference = 'Stop'
Import-Module SmbShare -ErrorAction Stop
$share = Get-SmbShare -Name $Name -ErrorAction SilentlyContinue
if ($share -and ($share.Path -ne $Directory -or $share.Description -ne 'TrioBridge AI online catalog')) {
    throw 'The share name is already used by another folder. No changes were made.'
}
if ($Action -eq 'Remove') {
    if ($share) { Remove-SmbShare -Name $Name -Force -Confirm:$false }
} elseif ($share) {
    $access = @(Get-SmbShareAccess -Name $Name)
    if ($access.Count -ne 1 -or $access[0].AccountName -ne $Account -or $access[0].AccessControlType -ne 'Allow' -or $access[0].AccessRight -ne 'Full') {
        throw 'The existing catalog has unexpected share permissions. Review it or uninstall before retrying.'
    }
} else {
    # Only the installing user can access this share. Do not grant Everyone access.
    New-SmbShare -Name $Name -Path $Directory -FullAccess $Account `
        -FolderEnumerationMode AccessBased -Description 'TrioBridge AI online catalog' | Out-Null
}
