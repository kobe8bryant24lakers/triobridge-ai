import type { ChatMessage } from './types';

export const MODEL_AUDIT_PATH = '/__triobridge-real-model-audit';

type AuditUsage = Record<string, number | Record<string, number>>;
export interface ModelAuditRecord {
  version: 1;
  auditId: string;
  startedAt: string;
  officeHost: string | null;
  elapsedMs: number;
  endpoint: string;
  requestedModel: string;
  returnedModel: string | null;
  completionId: string | null;
  providerRequestId: string | null;
  httpStatus: number | null;
  finishReason: string | null;
  usage: AuditUsage | null;
  messages: ChatMessage[];
  responseContent: string | null;
  responseContentTruncated: boolean;
  outcome: string;
}

const text = (value: unknown, limit: number): string | null =>
  typeof value === 'string' ? value.slice(0, limit) : null;
const object = (value: unknown): Record<string, unknown> | undefined =>
  value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
const count = (value: unknown): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;

function usageFields(value: unknown): AuditUsage | null {
  const source = object(value);
  if (!source) return null;
  const result: AuditUsage = {};
  for (const name of ['prompt_tokens', 'completion_tokens', 'total_tokens'])
    if (count(source[name])) result[name] = source[name];
  for (const name of ['prompt_tokens_details', 'completion_tokens_details']) {
    const details = object(source[name]);
    if (!details) continue;
    const selected: Record<string, number> = {};
    for (const key of [
      'cached_tokens',
      'audio_tokens',
      'text_tokens',
      'reasoning_tokens',
      'accepted_prediction_tokens',
      'rejected_prediction_tokens',
    ])
      if (count(details[key])) selected[key] = details[key];
    if (Object.keys(selected).length) result[name] = selected;
  }
  return Object.keys(result).length ? result : null;
}

/** Select fields again at the local server boundary; never persist arbitrary objects. */
export function normalizeAuditRecord(value: unknown): ModelAuditRecord | null {
  const source = object(value);
  if (
    !source ||
    source.version !== 1 ||
    typeof source.auditId !== 'string' ||
    !/^[\w-]{1,80}$/.test(source.auditId) ||
    typeof source.startedAt !== 'string' ||
    !Number.isFinite(Date.parse(source.startedAt)) ||
    !count(source.elapsedMs) ||
    typeof source.endpoint !== 'string' ||
    typeof source.requestedModel !== 'string' ||
    !Array.isArray(source.messages) ||
    source.messages.length > 16 ||
    typeof source.outcome !== 'string' ||
    !['success', 'auth', 'rate', 'network', 'timeout', 'cancelled', 'response', 'config'].includes(
      source.outcome,
    )
  )
    return null;
  let endpoint: string;
  try {
    const url = new URL(source.endpoint, 'http://localhost');
    if (!['http:', 'https:'].includes(url.protocol)) return null;
    endpoint = source.endpoint.startsWith('/') ? url.pathname : `${url.origin}${url.pathname}`;
  } catch {
    return null;
  }
  const messages: ChatMessage[] = [];
  for (const item of source.messages) {
    const message = object(item);
    if (
      !message ||
      !['system', 'user', 'assistant'].includes(String(message.role)) ||
      typeof message.content !== 'string' ||
      message.content.length > 100_000
    )
      return null;
    messages.push({ role: message.role as ChatMessage['role'], content: message.content });
  }
  return {
    version: 1,
    auditId: source.auditId,
    startedAt: source.startedAt,
    officeHost: ['Word', 'Excel', 'PowerPoint'].includes(String(source.officeHost))
      ? String(source.officeHost)
      : null,
    elapsedMs: source.elapsedMs,
    endpoint,
    requestedModel: source.requestedModel.slice(0, 200),
    returnedModel: text(source.returnedModel, 200),
    completionId: text(source.completionId, 300),
    providerRequestId: text(source.providerRequestId, 300),
    httpStatus:
      count(source.httpStatus) && source.httpStatus >= 100 && source.httpStatus <= 599
        ? source.httpStatus
        : null,
    finishReason: text(source.finishReason, 80),
    usage: usageFields(source.usage),
    messages,
    responseContent: text(source.responseContent, 100_000),
    responseContentTruncated:
      source.responseContentTruncated === true ||
      (typeof source.responseContent === 'string' && source.responseContent.length > 100_000),
    outcome: source.outcome,
  };
}

/** Explicit development-only diagnostics for synthetic acceptance data. */
export function startModelAudit(
  endpoint: string,
  requestedModel: string,
  messages: ChatMessage[],
  secretToRedact: string,
) {
  if (!import.meta.env.DEV || import.meta.env.VITE_REAL_MODEL_AUDIT !== '1') return undefined;
  const started = performance.now();
  let officeHost: string | null = null;
  try {
    if (typeof Office !== 'undefined' && Office.context?.host)
      officeHost = String(Office.context.host);
  } catch {
    // Office may not have initialized yet, e.g. a standalone connection test.
  }
  const record: ModelAuditRecord = {
    version: 1,
    auditId: crypto.randomUUID(),
    startedAt: new Date().toISOString(),
    officeHost,
    elapsedMs: 0,
    endpoint,
    requestedModel,
    returnedModel: null,
    completionId: null,
    providerRequestId: null,
    httpStatus: null,
    finishReason: null,
    usage: null,
    messages: messages.map(({ role, content }) => ({
      role,
      content:
        typeof content === 'string'
          ? content
          : content
              .map((part) => (part.type === 'text' ? part.text : '[图片内容未记录]'))
              .join('\n'),
    })),
    responseContent: null,
    responseContentTruncated: false,
    outcome: 'network',
  };
  return {
    response(response: Response) {
      record.httpStatus = response.status;
      // CORS may hide this header; null means unavailable, not absent at the provider.
      record.providerRequestId = text(response.headers.get('x-request-id'), 300);
    },
    payload(value: unknown) {
      const data = object(value);
      const localModel = object(data?.localModel);
      if (typeof localModel?.endpoint === 'string') record.endpoint = localModel.endpoint;
      const choice = Array.isArray(data?.choices) ? object(data.choices[0]) : undefined;
      const content = object(choice?.message)?.content;
      record.returnedModel = text(data?.model, 200);
      record.completionId = text(data?.id, 300);
      record.finishReason = text(choice?.finish_reason, 80);
      record.usage = usageFields(data?.usage);
      record.responseContent = text(content, 100_000);
      record.responseContentTruncated = typeof content === 'string' && content.length > 100_000;
    },
    async finish(outcome: string): Promise<void> {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 2000);
      try {
        record.elapsedMs = Math.max(0, Math.round(performance.now() - started));
        record.outcome = outcome;
        const normalized = normalizeAuditRecord(record);
        if (!normalized) return;
        // A provider could echo a key in content or an ID. Redact it in every selected string.
        const redact = (value: unknown): unknown => {
          if (typeof value === 'string')
            return secretToRedact ? value.split(secretToRedact).join('[REDACTED]') : value;
          if (Array.isArray(value)) return value.map(redact);
          if (value && typeof value === 'object')
            return Object.fromEntries(
              Object.entries(value).map(([key, item]) => [key, redact(item)]),
            );
          return value;
        };
        const response = await fetch(MODEL_AUDIT_PATH, {
          method: 'POST',
          credentials: 'omit',
          redirect: 'error',
          cache: 'no-store',
          signal: controller.signal,
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(redact(normalized)),
        });
        if (!response.ok) throw new Error('Audit rejected');
      } catch {
        // Diagnostics must not replace a successful model result or disclose error bodies.
        console.warn('Real model audit was not persisted.');
      } finally {
        clearTimeout(timer);
      }
    },
  };
}
