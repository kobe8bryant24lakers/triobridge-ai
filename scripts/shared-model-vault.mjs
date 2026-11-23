/** Secrets remain in the current OS user's credential vault, never in CLI arguments/files. */
export function createNativeModelVault({
  platform = process.platform,
  loadBinding = () => import('@napi-rs/keyring'),
} = {}) {
  if (!['darwin', 'win32'].includes(platform))
    throw new Error(
      'Shared model credentials require macOS Keychain or Windows Credential Manager.',
    );
  let binding;
  const entry = async (reference) => {
    binding ??= loadBinding();
    const { AsyncEntry } = await binding;
    return new AsyncEntry('ai.triobridge.shared-models.v1', reference);
  };
  return {
    kind: platform === 'darwin' ? 'macos-keychain' : 'windows-credential-manager',
    async get(reference, signal) {
      return (await entry(reference)).getPassword(signal);
    },
    async set(reference, value, signal) {
      signal?.throwIfAborted();
      const credential = await entry(reference);
      signal?.throwIfAborted();
      // A native worker may still be writing when its JS AbortSignal rejects.
      // Wait for this atomic OS operation, then the catalog transaction handles cancellation
      // and deletes the staged entry. Otherwise a late native write could outlive rollback.
      await credential.setPassword(value);
    },
    async delete(reference) {
      await (await entry(reference)).deleteCredential();
    },
  };
}
