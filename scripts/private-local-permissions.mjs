import { execFile } from 'node:child_process';
import { chmod, lstat } from 'node:fs/promises';
import { win32 } from 'node:path';
import { promisify } from 'node:util';

const execute = promisify(execFile);
// An intermediate Node process can inherit PowerShell 7 module paths. Windows
// PowerShell then resolves built-in cmdlets to incompatible modules. Let 5.1
// construct its own search path without changing the parent environment.
export function windowsPowerShellEnvironment(environment = process.env) {
  return Object.fromEntries(
    Object.entries(environment).filter(([name]) => name.toLowerCase() !== 'psmodulepath'),
  );
}
// Only non-secret paths and a fixed kind travel through the child environment.
// The command is fixed: no path, token, API key or caller input is interpolated.
export const WINDOWS_PRIVATE_ACL_SCRIPT = `
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$stage = 'INSPECT'
try {
$path = [Environment]::GetEnvironmentVariable('TRIOBRIDGE_PRIVATE_PATH', 'Process')
$kind = [Environment]::GetEnvironmentVariable('TRIOBRIDGE_PRIVATE_KIND', 'Process')
$fresh = [Environment]::GetEnvironmentVariable('TRIOBRIDGE_PRIVATE_CREATED', 'Process') -eq '1'
if ([String]::IsNullOrWhiteSpace($path) -or $kind -notin @('file', 'directory')) { throw 'Invalid private path request.' }
$stage = 'GET_ITEM'
$item = Get-Item -LiteralPath $path -Force
$stage = 'TYPE'
if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Linked private paths are not supported.' }
if (($kind -eq 'directory') -ne [bool]$item.PSIsContainer) { throw 'Private path type mismatch.' }
$stage = 'IDENTITY'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$stage = 'GET_ACL'
$current = Get-Acl -LiteralPath $path
$stage = 'READ_OWNER'
$owner = $current.GetOwner([Security.Principal.SecurityIdentifier]).Value
if ($owner -ne $identity.User.Value) {
  $stage = 'OWNER'
  $principal = New-Object Security.Principal.WindowsPrincipal($identity)
  if (-not ($fresh -and $owner -eq 'S-1-5-32-544' -and $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
    throw 'TBACL_OWNER: Private path belongs to another OS account.'
  }
}
$stage = 'BUILD'
if ($kind -eq 'directory') { $acl = New-Object Security.AccessControl.DirectorySecurity }
else { $acl = New-Object Security.AccessControl.FileSecurity }
$acl.SetOwner($identity.User)
$acl.SetAccessRuleProtection($true, $false)
$allowed = @($identity.User.Value, 'S-1-5-18', 'S-1-5-32-544') | Select-Object -Unique
foreach ($sid in $allowed) {
  $principal = New-Object Security.Principal.SecurityIdentifier($sid)
  if ($kind -eq 'directory') {
    $rule = New-Object Security.AccessControl.FileSystemAccessRule($principal, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
  } else {
    $rule = New-Object Security.AccessControl.FileSystemAccessRule($principal, 'FullControl', 'Allow')
  }
  [void]$acl.AddAccessRule($rule)
}
$stage = 'APPLY'
Set-Acl -LiteralPath $path -AclObject $acl
$stage = 'VERIFY'
$verified = Get-Acl -LiteralPath $path
if (-not $verified.AreAccessRulesProtected -or $verified.GetOwner([Security.Principal.SecurityIdentifier]).Value -ne $identity.User.Value) { throw 'Private ACL verification failed.' }
foreach ($rule in $verified.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier])) {
  if ($rule.IdentityReference.Value -notin $allowed -or $rule.AccessControlType -ne 'Allow' -or $rule.IsInherited) { throw 'Private ACL contains an unexpected rule.' }
}
} catch {
  [Console]::Error.WriteLine('TBACL_' + $stage)
  exit 1
}
`;

export async function protectWindowsPath(
  path,
  kind,
  { runner = execute, environment = process.env, created = false } = {},
) {
  if (typeof path !== 'string' || !path || !['file', 'directory'].includes(kind))
    throw new Error('Invalid private path request.');
  const systemRoot = environment.SystemRoot || environment.SYSTEMROOT;
  if (!systemRoot || !win32.isAbsolute(systemRoot))
    throw new Error('Windows system directory is unavailable.');
  const executable = win32.join(
    systemRoot,
    'System32',
    'WindowsPowerShell',
    'v1.0',
    'powershell.exe',
  );
  try {
    await runner(
      executable,
      [
        '-NoProfile',
        '-NonInteractive',
        '-EncodedCommand',
        Buffer.from(WINDOWS_PRIVATE_ACL_SCRIPT, 'utf16le').toString('base64'),
      ],
      {
        env: {
          ...windowsPowerShellEnvironment(environment),
          TRIOBRIDGE_PRIVATE_PATH: path,
          TRIOBRIDGE_PRIVATE_KIND: kind,
          TRIOBRIDGE_PRIVATE_CREATED: created ? '1' : '0',
        },
        windowsHide: true,
        timeout: 60_000,
        maxBuffer: 8192,
      },
    );
  } catch (error) {
    const stage = String(error.stderr || '').match(
      /TBACL_(INSPECT|GET_ITEM|TYPE|IDENTITY|GET_ACL|READ_OWNER|OWNER|BUILD|APPLY|VERIFY)\b/,
    )?.[1];
    throw new Error(
      'Could not enforce private permissions for the current Windows account.' +
        (stage ? ` (ACL_${stage})` : ''),
    );
  }
}

const protectedWindowsPaths = new Map();
/** POSIX modes on Mac; protected user/SYSTEM/Administrators DACL on Windows. */
export async function enforcePrivatePath(path, kind, { created = false } = {}) {
  const info = await lstat(path);
  if (
    (kind === 'directory' ? !info.isDirectory() : !info.isFile()) ||
    info.isSymbolicLink() ||
    (process.getuid && info.uid !== process.getuid())
  )
    throw new Error('Private path is invalid or belongs to another OS account.');
  if (process.platform === 'win32') {
    const identity = `${info.dev}:${info.ino}:${info.birthtimeMs}`;
    if (protectedWindowsPaths.get(path) !== identity) {
      await protectWindowsPath(path, kind, { created });
      protectedWindowsPaths.set(path, identity);
    }
  } else {
    await chmod(path, kind === 'directory' ? 0o700 : 0o600);
  }
}
