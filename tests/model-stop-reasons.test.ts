import { afterEach, describe, expect, it, vi } from 'vitest';
import { complete, testConnection } from '../src/core/model-client';
import { emptyProfile } from '../src/core/profile-store';
import * as capability from '../src/core/local-capability';
import type { ModelProfile, SharedModelProfile } from '../src/core/types';
import { normalizeModelResponse } from '../scripts/model-protocol.mjs';

const profile: ModelProfile = {
  ...emptyProfile(),
  name: 'Stop reason test',
  model: 'synthetic',
  baseUrl: 'https://example.test/v1',
  apiFormat: 'anthropic',
};
const messages = [{ role: 'user' as const, content: '改写这段合成文字。' }];
const claude = (reason: string, text: string) => ({
  stop_reason: reason,
  content: text ? [{ type: 'text', text }] : [{ type: 'thinking', thinking: 'private thinking' }],
});

describe('model stop reasons', () => {
  afterEach(() => vi.unstubAllGlobals());

  it.each([
    ['max_tokens', 'length'],
    ['model_context_window_exceeded', 'context_length'],
    ['refusal', 'content_filter'],
  ])(
    'preserves the distinct Claude stop reason %s without exposing thinking',
    (reason, normalized) => {
      expect(normalizeModelResponse('anthropic', claude(reason, ''))).toMatchObject({
        choices: [{ finish_reason: normalized, message: { content: '' } }],
      });
      expect(JSON.stringify(normalizeModelResponse('anthropic', claude(reason, '')))).not.toContain(
        'private thinking',
      );
    },
  );

  it.each(['', '未完成的改写片段'])(
    'rejects context exhaustion before returning any writable text: %s',
    async (text) => {
      const response = claude('model_context_window_exceeded', text);
      const fetcher = vi.fn().mockImplementation(async () => Response.json(response));
      await expect(complete(profile, messages, undefined, fetcher)).rejects.toMatchObject({
        code: 'response',
        message: expect.stringContaining('上下文空间已用尽'),
      });
      expect(fetcher).toHaveBeenCalledTimes(1);
    },
  );

  it('reports context exhaustion as a connection warning, separately from output-token exhaustion', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockImplementation(async () => Response.json(claude('model_context_window_exceeded', ''))),
    );
    const result = await testConnection(profile);
    expect(result.warning).toContain('上下文空间用尽');
    expect(result.warning).not.toContain('输出长度上限');
  });

  it.each([
    ['length', '长度限制截断'],
    ['context_length', '上下文空间已用尽'],
    ['content_filter', '模型未完成此请求'],
  ])('retains the shared/native diagnosis for %s with null content', async (reason, diagnostic) => {
    vi.spyOn(capability, 'getLocalCapability').mockReturnValue('test-token');
    const shared: SharedModelProfile = {
      id: profile.id,
      name: profile.name,
      model: profile.model,
      baseUrl: profile.baseUrl,
      apiFormat: 'anthropic',
      timeoutSeconds: profile.timeoutSeconds,
      kind: 'shared',
      revision: 1,
      hasKey: false,
    };
    const fetcher = vi.fn().mockImplementation(async () =>
      Response.json({
        choices: [{ finish_reason: reason, message: { content: null } }],
      }),
    );
    await expect(complete(shared, messages, undefined, fetcher)).rejects.toMatchObject({
      code: 'response',
      message: expect.stringContaining(diagnostic),
    });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
});
