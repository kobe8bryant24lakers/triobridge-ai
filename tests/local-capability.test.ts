import { describe, it, expect, vi } from 'vitest';
import {
  bootstrapLocalCapability,
  getLocalCapability,
  isLocalSharedOrigin,
  requiresLocalSharedMode,
} from '../src/core/local-capability';

const token = 'synthetic_capability_'.padEnd(43, 'x');
function context(origin = 'https://localhost:3000', fragment = `#local-token=${token}`) {
  return {
    location: { origin, href: `${origin}/?office=1${fragment}` },
    history: { state: null, replaceState: vi.fn() },
    sessionStorage,
  };
}

describe('per-user local capability bootstrap', () => {
  it('requires shared mode for local Office entries even when their token is missing', () => {
    for (const origin of ['https://localhost:3000', 'https://127.0.0.1:3000', 'https://[::1]:3000'])
      expect(requiresLocalSharedMode({ origin, search: '?office=1' })).toBe(true);
    for (const location of [
      { origin: 'https://example.com', search: '?office=1' },
      { origin: 'https://localhost.evil.test', search: '?office=1' },
      { origin: 'http://localhost:3000', search: '?office=1' },
      { origin: 'https://localhost:3000', search: '' },
    ])
      expect(requiresLocalSharedMode(location)).toBe(false);
  });
  it('consumes the fragment before use and only retains it in session storage', () => {
    const target = context();
    bootstrapLocalCapability(target);
    expect(target.history.replaceState).toHaveBeenCalledWith(null, '', '/?office=1');
    expect(getLocalCapability(target)).toBe(token);
    expect(localStorage.length).toBe(0);
    expect(sessionStorage.length).toBe(1);
  });
  it('removes invalid and duplicate tokens without accepting them', () => {
    for (const fragment of ['#local-token=short', `#local-token=${token}&local-token=${token}`]) {
      const target = context(undefined, fragment);
      bootstrapLocalCapability(target);
      expect(target.history.replaceState).toHaveBeenCalledWith(null, '', '/?office=1');
      expect(getLocalCapability(target)).toBeUndefined();
    }
  });
  it('never exposes a capability to remote or insecure origins', () => {
    for (const origin of [
      'https://example.com',
      'http://localhost:3000',
      'https://localhost.evil.test',
    ]) {
      const target = context(origin);
      bootstrapLocalCapability(target);
      expect(getLocalCapability(target)).toBeUndefined();
      expect(sessionStorage.length).toBe(0);
    }
    expect(isLocalSharedOrigin('https://127.0.0.1:3000')).toBe(true);
    expect(isLocalSharedOrigin('https://[::1]:3000')).toBe(true);
  });
  it('preserves other fragment parameters and observes session clearing', () => {
    const target = context(undefined, `#page=settings&local-token=${token}`);
    bootstrapLocalCapability(target);
    expect(target.history.replaceState).toHaveBeenCalledWith(null, '', '/?office=1#page=settings');
    sessionStorage.clear();
    expect(getLocalCapability(target)).toBeUndefined();
  });
});
