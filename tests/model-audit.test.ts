import { afterEach, describe, expect, it, vi } from 'vitest';
import { complete } from '../src/core/model-client';
import { MODEL_AUDIT_PATH, normalizeAuditRecord } from '../src/core/model-audit';
import { emptyProfile } from '../src/core/profile-store';

const key = 'audit-test-key-must-never-be-recorded';
const profile = {
  ...emptyProfile(),
  name: 'Synthetic audit test',
  baseUrl: 'https://example.test/v1',
  model: 'requested-model',
  apiKey: key,
};
const messages = [{ role: 'user' as const, content: '合成验收数据：计划 88 小时。' }];
const response = () =>
  new Response(
    JSON.stringify({
      id: 'completion-test',
      model: 'returned-model',
      usage: {
        prompt_tokens: 20,
        completion_tokens: 30,
        total_tokens: 50,
        prompt_tokens_details: { cached_tokens: 8, authorization: key },
        completion_tokens_details: { reasoning_tokens: 5 },
        apiKey: key,
      },
      choices: [{ finish_reason: 'stop', message: { content: `结果；意外回显 ${key}` } }],
      headers: { Authorization: key },
      apiKey: key,
    }),
    { headers: { 'x-request-id': 'request-test', 'x-secret': key, 'set-cookie': key } },
  );

afterEach(() => {
  vi.unstubAllEnvs();
  vi.unstubAllGlobals();
});

describe('explicit real-model development audit', () => {
  it.each([
    { dev: true, flag: '' },
    { dev: false, flag: '1' },
    { dev: true, flag: 'true' },
  ])('never sends audit traffic with DEV=$dev flag=$flag', async ({ dev, flag }) => {
    vi.stubEnv('DEV', dev);
    vi.stubEnv('VITE_REAL_MODEL_AUDIT', flag);
    const auditFetch = vi.fn();
    vi.stubGlobal('fetch', auditFetch);
    const modelFetch = vi.fn().mockResolvedValue(response());
    expect(await complete(profile, messages, undefined, modelFetch)).toContain('结果');
    expect(modelFetch).toHaveBeenCalledOnce();
    expect(auditFetch).not.toHaveBeenCalled();
  });

  it('records selected response evidence and synthetic content without credentials or other headers', async () => {
    vi.stubEnv('DEV', true);
    vi.stubEnv('VITE_REAL_MODEL_AUDIT', '1');
    vi.stubGlobal('Office', { context: { host: 'PowerPoint' } });
    const auditFetch = vi.fn().mockResolvedValue(new Response(null, { status: 204 }));
    vi.stubGlobal('fetch', auditFetch);
    await complete(
      profile,
      [...messages, { role: 'user', content: `Never persist ${key}` }],
      undefined,
      vi.fn().mockResolvedValue(response()),
    );
    expect(auditFetch).toHaveBeenCalledOnce();
    const [url, init] = auditFetch.mock.calls[0];
    expect(url).toBe(MODEL_AUDIT_PATH);
    expect(init.headers).toEqual({ 'Content-Type': 'application/json' });
    expect(init.credentials).toBe('omit');
    expect(init.redirect).toBe('error');
    const serialized = init.body as string;
    expect(serialized).not.toContain(key);
    expect(serialized).not.toContain('Authorization');
    expect(serialized).not.toContain('apiKey');
    expect(serialized).not.toContain('set-cookie');
    expect(serialized).not.toContain('x-secret');
    expect(JSON.parse(serialized)).toMatchObject({
      endpoint: 'https://example.test/v1/chat/completions',
      officeHost: 'PowerPoint',
      requestedModel: 'requested-model',
      returnedModel: 'returned-model',
      completionId: 'completion-test',
      providerRequestId: 'request-test',
      httpStatus: 200,
      finishReason: 'stop',
      usage: {
        prompt_tokens: 20,
        completion_tokens: 30,
        total_tokens: 50,
        prompt_tokens_details: { cached_tokens: 8 },
        completion_tokens_details: { reasoning_tokens: 5 },
      },
      messages: [...messages, { role: 'user', content: 'Never persist [REDACTED]' }],
      responseContent: '结果；意外回显 [REDACTED]',
      outcome: 'success',
    });
  });

  it('records an HTTP failure without reading or recording its arbitrary error body', async () => {
    vi.stubEnv('DEV', true);
    vi.stubEnv('VITE_REAL_MODEL_AUDIT', '1');
    const auditFetch = vi.fn().mockResolvedValue(new Response(null, { status: 204 }));
    vi.stubGlobal('fetch', auditFetch);
    const modelFetch = vi.fn().mockResolvedValue(new Response(`private ${key}`, { status: 401 }));
    await expect(complete(profile, messages, undefined, modelFetch)).rejects.toMatchObject({
      code: 'auth',
    });
    const record = JSON.parse(auditFetch.mock.calls[0][1].body);
    expect(record).toMatchObject({ httpStatus: 401, outcome: 'auth', responseContent: null });
    expect(JSON.stringify(record)).not.toContain('private');
  });

  it('does not change the successful model result when local auditing fails', async () => {
    vi.stubEnv('DEV', true);
    vi.stubEnv('VITE_REAL_MODEL_AUDIT', '1');
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error(`local ${key}`)));
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {});
    expect(
      await complete(profile, messages, undefined, vi.fn().mockResolvedValue(response())),
    ).toBe(`结果；意外回显 ${key}`);
    expect(warning).toHaveBeenCalledWith('Real model audit was not persisted.');
  });

  it('normalizes endpoint credentials/query/fragment away and rejects invalid record shapes', () => {
    const record = {
      version: 1,
      auditId: 'test-1',
      startedAt: new Date().toISOString(),
      elapsedMs: 12,
      endpoint: `https://user:${key}@example.test/v1/chat/completions?key=${key}#${key}`,
      requestedModel: 'test',
      messages,
      outcome: 'success',
      Authorization: key,
      headers: { Authorization: key },
    };
    const normalized = normalizeAuditRecord(record);
    expect(normalized?.endpoint).toBe('https://example.test/v1/chat/completions');
    expect(JSON.stringify(normalized)).not.toContain(key);
    expect(
      normalizeAuditRecord({ ...record, messages: [{ role: 'user', content: { key } }] }),
    ).toBeNull();
    expect(normalizeAuditRecord({ ...record, outcome: key })).toBeNull();
  });
});
