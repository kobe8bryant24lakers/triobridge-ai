import { describe, it, expect, vi } from 'vitest';
import { complete } from '../src/core/model-client';
import {
  createSharedModelClient,
  parseSharedState,
  SharedModelClientError,
} from '../src/core/shared-model-client';
import * as capability from '../src/core/local-capability';

const token = 'test_capability_'.padEnd(43, 'x');
const publicProfile = {
  id: 'shared-one',
  name: 'Shared',
  baseUrl: 'https://example.com/v1',
  model: 'one',
  timeoutSeconds: 60,
  hasKey: true,
  revision: 1,
};
const state = {
  version: 2,
  revision: 1,
  activeId: 'shared-one',
  profiles: [publicProfile],
  storage: 'macos-keychain',
};
const messages = [{ role: 'user' as const, content: 'synthetic request' }];
const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status });

describe('shared metadata and fixed local routing', () => {
  it('keeps shared metadata structurally free of saved keys', () => {
    const parsed = parseSharedState({ ...state, warning: 'cleanup pending' });
    expect(parsed.profiles[0]).toEqual({ ...publicProfile, kind: 'shared', apiFormat: 'openai' });
    expect(
      parseSharedState({ ...state, profiles: [{ ...publicProfile, apiFormat: 'anthropic' }] })
        .profiles[0].apiFormat,
    ).toBe('anthropic');
    expect(() =>
      parseSharedState({ ...state, profiles: [{ ...publicProfile, apiFormat: 'unknown' }] }),
    ).toThrow();
    expect(parsed.profiles[0]).not.toHaveProperty('apiKey');
    expect(parsed.warning).toBe('cleanup pending');
    expect(() =>
      parseSharedState({ ...state, profiles: [{ ...publicProfile, apiKey: 'must-not-return' }] }),
    ).toThrow();
    expect(() => parseSharedState({ ...state, activeId: 'missing' })).toThrow();
    expect(() => parseSharedState({ ...state, revision: 0 })).toThrow();
  });
  it('does not request shared service without a bootstrap capability', async () => {
    const fetcher = vi.fn();
    await expect(createSharedModelClient(fetcher, () => undefined).state()).rejects.toMatchObject({
      code: 'capability',
    });
    expect(fetcher).not.toHaveBeenCalled();
  });
  it('rejects static HTML and transmits only the capability to the fixed same-origin path', async () => {
    const fetcher = vi.fn().mockResolvedValue(new Response('<!doctype html>'));
    await expect(createSharedModelClient(fetcher, () => token).state()).rejects.toMatchObject({
      code: 'unavailable',
    });
    expect(fetcher).toHaveBeenCalledWith(
      '/local-models/state',
      expect.objectContaining({
        method: 'POST',
        credentials: 'omit',
        redirect: 'error',
        body: '{}',
        headers: { 'Content-Type': 'application/json', 'X-TrioBridge-Local-Token': token },
      }),
    );
  });
  it('preserves optimistic conflict state and never retries the mutation', async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValue(
        json({ error: { code: 'conflict', message: 'refresh first' }, state }, 409),
      );
    await expect(
      createSharedModelClient(fetcher, () => token).remove({
        expectedRevision: 0,
        id: publicProfile.id,
      }),
    ).rejects.toMatchObject({
      status: 409,
      code: 'conflict',
      state: { revision: 1 },
    });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
  it('generates via profileId/revision only; no API key or URL/model override', async () => {
    vi.spyOn(capability, 'getLocalCapability').mockReturnValue(token);
    const fetcher = vi
      .fn()
      .mockResolvedValue(json({ choices: [{ message: { content: 'shared result' } }] }));
    const profile = parseSharedState(state).profiles[0];
    expect(await complete(profile, messages, undefined, fetcher)).toBe('shared result');
    expect(fetcher).toHaveBeenCalledWith(
      '/local-models/complete',
      expect.objectContaining({
        body: JSON.stringify({ profileId: 'shared-one', profileRevision: 1, messages }),
        headers: { 'Content-Type': 'application/json', 'X-TrioBridge-Local-Token': token },
      }),
    );
    expect(fetcher.mock.calls[0][1].headers).not.toHaveProperty('Authorization');
  });
  it('distinguishes an unavailable local bridge from an unreachable model API', async () => {
    const failure = vi.fn().mockRejectedValue(new TypeError('private error content'));
    await expect(createSharedModelClient(failure, () => token).state()).rejects.toMatchObject({
      code: 'local_network',
      message: expect.stringContaining('本机共享服务'),
    });
    vi.spyOn(capability, 'getLocalCapability').mockReturnValue(token);
    const profile = parseSharedState(state).profiles[0];
    await expect(complete(profile, messages, undefined, failure)).rejects.toMatchObject({
      code: 'network',
      message: expect.stringContaining('本机共享服务'),
    });
    const upstream = vi
      .fn()
      .mockResolvedValue(
        json({ error: { code: 'network', message: 'API refused: ECONNREFUSED' } }, 502),
      );
    await expect(complete(profile, messages, undefined, upstream)).rejects.toMatchObject({
      code: 'network',
      message: 'API refused: ECONNREFUSED',
    });
  });
  it('rejects stale shared generation without silently reissuing it', async () => {
    vi.spyOn(capability, 'getLocalCapability').mockReturnValue(token);
    const fetcher = vi
      .fn()
      .mockResolvedValue(
        json({ error: { code: 'conflict', message: 'refresh profile' }, state }, 409),
      );
    await expect(
      complete(parseSharedState(state).profiles[0], messages, undefined, fetcher),
    ).rejects.toMatchObject({ code: 'config', message: 'refresh profile' });
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(new SharedModelClientError('x', 'conflict', 409).status).toBe(409);
  });
});
