import type { ChatMessage, ModelApiFormat } from '../src/core/types';

export function isModelApiFormat(value: unknown): value is ModelApiFormat;
export function modelEndpoint(baseUrl: string, apiFormat?: ModelApiFormat): string;
export class ModelProtocolError extends Error {}
export function modelHeaders(
  apiFormat: ModelApiFormat | undefined,
  key?: string,
): Record<string, string>;
export function modelRequest(
  apiFormat: ModelApiFormat | undefined,
  model: string,
  messages: ChatMessage[],
  key: string,
): { headers: Record<string, string>; body: Record<string, unknown> };
export function normalizeModelResponse(
  apiFormat: ModelApiFormat | undefined,
  data: unknown,
): unknown;
