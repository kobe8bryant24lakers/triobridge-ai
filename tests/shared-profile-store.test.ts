import { afterEach, describe, it, expect, vi } from 'vitest';
import {
  createProfileController,
  emptyProfile,
  loadProfiles,
  PROFILE_KEY,
  saveProfiles,
  SESSION_KEY,
} from '../src/core/profile-store';

const token = 'synthetic_token_'.padEnd(43, 'x');
const p = {
  id: 'shared-one',
  name: 'Shared',
  baseUrl: 'https://example.com/v1',
  model: 'one',
  timeoutSeconds: 60,
  hasKey: true,
  revision: 1,
};
const snapshot = (revision = 1, activeId = p.id, profiles = [p]) => ({
  version: 2,
  revision,
  activeId,
  profiles,
  storage: 'macos-keychain',
});
const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status });
const controllers: ReturnType<typeof createProfileController>[] = [];
function controller(fetcher: typeof fetch, capability: () => string | undefined = () => token) {
  const instance = createProfileController({ fetcher, capability });
  controllers.push(instance);
  return instance;
}
afterEach(() => {
  controllers.splice(0).forEach((c) => c.stop());
  vi.useRealTimers();
});

describe('multi-host shared profile state', () => {
  it('stays local without capability and never probes the server', () => {
    const fetcher = vi.fn();
    const store = controller(fetcher, () => undefined);
    store.start();
    expect(store.getSnapshot().mode).not.toBe('shared');
    expect(fetcher).not.toHaveBeenCalled();
  });
  it('enables shared state without loading legacy credentials', async () => {
    saveProfiles({
      profiles: [{ ...emptyProfile(), name: 'Legacy', model: 'old', apiKey: 'legacy-secret' }],
      activeId: '',
    });
    const read = vi.spyOn(localStorage, 'getItem');
    const store = controller(vi.fn().mockResolvedValue(json(snapshot())));
    await store.actions.refresh();
    expect(store.getSnapshot()).toMatchObject({
      mode: 'shared',
      connection: 'ready',
      revision: 1,
      legacyAvailable: true,
    });
    expect(read).not.toHaveBeenCalledWith(PROFILE_KEY);
    expect(JSON.stringify(store.getSnapshot())).not.toContain('legacy-secret');
  });
  it('rejects stale polling and keeps shared state through an outage', async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(json(snapshot(3)))
      .mockResolvedValueOnce(json(snapshot(2)))
      .mockRejectedValueOnce(new TypeError('offline'));
    const store = controller(fetcher);
    await store.actions.refresh();
    await store.actions.refresh();
    expect(store.getSnapshot()).toMatchObject({ revision: 3 });
    await expect(store.actions.refresh()).rejects.toThrow('offline');
    expect(store.getSnapshot()).toMatchObject({
      mode: 'shared',
      revision: 3,
      connection: 'offline',
      profiles: [{ id: p.id }],
    });
  });
  it('updates from a conflict without silently saving the stale draft', async () => {
    const newer = { ...p, model: 'updated', revision: 2 };
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(json(snapshot()))
      .mockResolvedValueOnce(
        json(
          {
            error: { code: 'conflict', message: 'review changes' },
            state: snapshot(2, p.id, [newer]),
          },
          409,
        ),
      );
    const store = controller(fetcher);
    await store.actions.refresh();
    await expect(
      store.actions.upsert({
        expectedRevision: 1,
        profile: {
          id: p.id,
          name: p.name,
          baseUrl: p.baseUrl,
          model: 'stale draft',
          timeoutSeconds: 60,
        },
        credential: { action: 'keep' },
      }),
    ).rejects.toMatchObject({ status: 409 });
    expect(store.getSnapshot()).toMatchObject({ revision: 2, profiles: [{ model: 'updated' }] });
    expect(fetcher).toHaveBeenCalledTimes(2);
  });
  it('selectively migrates chosen profiles, keeping unsupported fixtures and their keys', async () => {
    const legacy = {
      ...emptyProfile(),
      id: 'legacy',
      name: 'Cloud',
      model: 'cloud',
      apiKey: 'cloud-private',
    };
    const fixture = {
      ...emptyProfile(),
      id: 'fixture',
      name: 'Fixture',
      baseUrl: '/model-proxy/v1',
      model: 'fixed',
      apiKey: 'fixture-private',
    };
    saveProfiles({ profiles: [legacy, fixture], activeId: legacy.id });
    const imported = {
      id: legacy.id,
      name: legacy.name,
      baseUrl: legacy.baseUrl,
      model: legacy.model,
      timeoutSeconds: 60,
      hasKey: true,
      revision: 1,
    };
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(json(snapshot(0, '', [])))
      .mockResolvedValueOnce(json(snapshot(1, legacy.id, [imported])));
    const store = controller(fetcher);
    await store.actions.refresh();
    const list = await store.actions.listLegacyProfiles();
    expect(list).toHaveLength(2);
    expect(JSON.stringify(list)).not.toContain('private');
    await store.actions.migrateLocal(0, [legacy.id]);
    expect(loadProfiles().profiles.map((x) => x.id)).toEqual(['fixture']);
    expect(sessionStorage.getItem(SESSION_KEY)).toContain('fixture-private');
    expect(sessionStorage.getItem(SESSION_KEY)).not.toContain('cloud-private');
    const body = JSON.parse(fetcher.mock.calls[1][1].body);
    expect(body.profiles).toHaveLength(1);
    expect(body.profiles[0].profile).not.toHaveProperty('apiKey');
    expect(body.profiles[0].credential).toEqual({ action: 'replace', value: 'cloud-private' });
  });
  it('preserves all local sources when atomic import fails', async () => {
    const legacy = { ...emptyProfile(), name: 'Legacy', model: 'old', apiKey: 'keep-private' };
    saveProfiles({ profiles: [legacy], activeId: legacy.id });
    const before = [localStorage.getItem(PROFILE_KEY), sessionStorage.getItem(SESSION_KEY)];
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(json(snapshot(0, '', [])))
      .mockResolvedValueOnce(
        json({ error: { code: 'config', message: 'proxy not configured' } }, 400),
      );
    const store = controller(fetcher);
    await store.actions.refresh();
    await expect(store.actions.migrateLocal(0, [legacy.id])).rejects.toThrow(
      'proxy not configured',
    );
    expect([localStorage.getItem(PROFILE_KEY), sessionStorage.getItem(SESSION_KEY)]).toEqual(
      before,
    );
  });
  it('retains concurrently changed local credentials after successful import', async () => {
    const legacy = {
      ...emptyProfile(),
      id: 'legacy',
      name: 'Legacy',
      model: 'old',
      apiKey: 'original-private',
    };
    saveProfiles({ profiles: [legacy], activeId: legacy.id });
    const imported = { ...p, id: legacy.id };
    const fetcher = vi
      .fn()
      .mockResolvedValueOnce(json(snapshot(0, '', [])))
      .mockImplementationOnce(async () => {
        saveProfiles({ profiles: [{ ...legacy, apiKey: 'changed-private' }], activeId: legacy.id });
        return json(snapshot(1, legacy.id, [imported]));
      });
    const store = controller(fetcher);
    await store.actions.refresh();
    const result = await store.actions.migrateLocal(0, [legacy.id]);
    expect(result.warning).toContain('发生变化');
    expect(loadProfiles().profiles[0].apiKey).toBe('changed-private');
  });
  it('refreshes on visibility and stops polling on disposal', async () => {
    vi.useFakeTimers();
    const fetcher = vi.fn().mockImplementation(async () => json(snapshot()));
    const store = controller(fetcher);
    store.start();
    await vi.waitFor(() => expect(store.getSnapshot().mode).toBe('shared'));
    await vi.advanceTimersByTimeAsync(5000);
    expect(fetcher).toHaveBeenCalledTimes(2);
    Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'visible' });
    document.dispatchEvent(new Event('visibilitychange'));
    await vi.waitFor(() => expect(fetcher).toHaveBeenCalledTimes(3));
    store.stop();
    await vi.advanceTimersByTimeAsync(10000);
    expect(fetcher).toHaveBeenCalledTimes(3);
  });
  it('keeps shared access usable when the WebView blocks localStorage access', async () => {
    const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'localStorage')!;
    Object.defineProperty(globalThis, 'localStorage', {
      configurable: true,
      get: () => {
        throw new DOMException('blocked', 'SecurityError');
      },
    });
    try {
      const store = controller(vi.fn().mockResolvedValue(json(snapshot())));
      await store.actions.refresh();
      expect(store.getSnapshot()).toMatchObject({
        mode: 'shared',
        connection: 'ready',
        legacyAvailable: false,
      });
    } finally {
      Object.defineProperty(globalThis, 'localStorage', descriptor);
    }
  });
});
