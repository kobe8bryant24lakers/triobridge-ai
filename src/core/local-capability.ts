const CAPABILITY_KEY = 'triobridge.local-capability.v1';
const TOKEN_PATTERN = /^[a-zA-Z0-9_-]{43,128}$/;
let memoryCapability: { origin: string; value: string } | undefined;

interface BootstrapContext {
  location: Pick<Location, 'href' | 'origin'>;
  history: Pick<History, 'replaceState' | 'state'>;
  sessionStorage: Storage;
}

export function isLocalSharedOrigin(origin: string): boolean {
  try {
    const url = new URL(origin);
    return url.protocol === 'https:' && ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
  } catch {
    return false;
  }
}

/** A locally installed Office pane must never silently fall back to browser credentials. */
export function requiresLocalSharedMode(
  location: Pick<Location, 'origin' | 'search'> = window.location,
): boolean {
  return (
    isLocalSharedOrigin(location.origin) &&
    new URLSearchParams(location.search).get('office') === '1'
  );
}

/** Consume the sideload-only bootstrap before rendering or initiating requests. */
export function bootstrapLocalCapability(context: BootstrapContext = window): void {
  const url = new URL(context.location.href);
  const fragment = new URLSearchParams(url.hash.slice(1));
  const values = fragment.getAll('local-token');
  if (!values.length) return;
  fragment.delete('local-token');
  url.hash = fragment.toString();
  // Clear the URL even if the token is invalid or session storage is blocked.
  context.history.replaceState(
    context.history.state,
    '',
    `${url.pathname}${url.search}${url.hash}`,
  );
  if (
    values.length !== 1 ||
    !TOKEN_PATTERN.test(values[0]) ||
    !isLocalSharedOrigin(context.location.origin)
  )
    return;
  try {
    context.sessionStorage.setItem(CAPABILITY_KEY, values[0]);
    memoryCapability = undefined;
  } catch {
    // The current pane remains usable; no persistent-storage fallback for this secret.
    memoryCapability = { origin: context.location.origin, value: values[0] };
  }
}

export function getLocalCapability(
  context: Pick<BootstrapContext, 'location' | 'sessionStorage'> = window,
): string | undefined {
  if (!isLocalSharedOrigin(context.location.origin)) return;
  try {
    const saved = context.sessionStorage.getItem(CAPABILITY_KEY);
    if (saved && TOKEN_PATTERN.test(saved)) return saved;
  } catch {
    // A restricted WebView may only permit the in-memory bootstrap.
  }
  if (memoryCapability?.origin === context.location.origin) return memoryCapability.value;
}
