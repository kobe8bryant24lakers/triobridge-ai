// Used by both the native model process and the direct client.
export function isModelApiFormat(value) {
  return value === 'openai' || value === 'anthropic';
}

export function modelEndpoint(baseUrl, apiFormat = 'openai') {
  if (!isModelApiFormat(apiFormat)) throw new Error('请选择支持的接口格式（OpenAI 或 Claude）。');
  if (typeof baseUrl !== 'string' || baseUrl.length > 2048) throw new Error('模型服务地址无效。');
  const base = baseUrl.trim().replace(/\/+$/, '');
  const proxy = /^\/model-proxy(?:\/[\w-]+)*$/.test(base);
  let url;
  try {
    url = proxy ? new URL(base, 'http://localhost') : new URL(base);
  } catch {
    throw new Error('请填写完整的 API Base URL，例如 http://127.0.0.1:8001/v1。');
  }
  if (
    !['https:', 'http:'].includes(url.protocol) ||
    url.username ||
    url.password ||
    url.search ||
    url.hash
  )
    throw new Error('服务地址须使用 HTTP(S)，且不能包含账号、密码、查询参数或片段。');
  if (url.protocol === 'http:' && !['localhost', '127.0.0.1', '[::1]'].includes(url.hostname))
    throw new Error('远程模型服务需要 HTTPS，HTTP 仅限本机回环地址。');
  const path = url.pathname.replace(/\/+$/, '');
  // Existing OpenAI gateways may mount Chat Completions under any base path.
  if (apiFormat === 'openai')
    return path.endsWith('/chat/completions') ? base : base + '/chat/completions';
  if (path.endsWith('/chat/completions'))
    throw new Error(
      '该地址是 OpenAI Chat Completions 接口。请选择 OpenAI 格式，或填写 Claude 接口地址。',
    );
  if (path.endsWith('/messages')) return base;
  return base + (path.endsWith('/v1') ? '/messages' : '/v1/messages');
}

export class ModelProtocolError extends Error {}

// A connection probe verifies the endpoint and credentials, not whether a document
// generation finished. Local reasoning models may exhaust the server's output
// budget before producing final text, while still returning a valid response.
export function connectionTestWarning(data) {
  const choice = Array.isArray(data?.choices) ? data.choices[0] : undefined;
  const message = choice?.message;
  if (
    !message ||
    typeof message !== 'object' ||
    Array.isArray(message) ||
    (typeof message.content !== 'string' && message.content !== null)
  )
    throw new ModelProtocolError('连接测试未返回兼容的文本响应，请检查接口格式。');
  if (message.refusal || choice.finish_reason === 'content_filter')
    throw new ModelProtocolError('模型未完成连接测试，请检查模型权限或服务策略。');
  if (typeof message.content === 'string' && message.content.length > 100_000)
    throw new ModelProtocolError('模型返回内容过长，请检查服务端配置。');
  if (choice.finish_reason === 'length')
    return '测试回复达到服务端输出长度上限；正式生成前请提高服务端输出上限。';
  if (choice.finish_reason === 'context_length')
    return '测试回复因模型上下文空间用尽而停止；正式生成前请减少输入内容，或在服务端增大上下文窗口。';
  if (!message.content?.trim() || (choice.finish_reason != null && choice.finish_reason !== 'stop'))
    throw new ModelProtocolError('连接测试未返回可用文本，请检查模型输出和接口格式。');
  return undefined;
}

function anthropicContent(content) {
  if (typeof content === 'string') return content;
  return content.map((part) => {
    if (part.type === 'text') return { type: 'text', text: part.text };
    const image = /^data:(image\/(?:jpeg|png));base64,([A-Za-z0-9+/]+={0,2})$/.exec(
      part.image_url?.url,
    );
    if (!image) throw new ModelProtocolError('Claude 图片输入需要本地 PNG 或 JPEG 图片。');
    return { type: 'image', source: { type: 'base64', media_type: image[1], data: image[2] } };
  });
}

export function modelHeaders(apiFormat, key) {
  const headers = { 'Content-Type': 'application/json' };
  if (apiFormat === 'anthropic') {
    headers['anthropic-version'] = '2023-06-01';
    if (key) headers['x-api-key'] = key;
  } else if (key) headers.Authorization = 'Bearer ' + key;
  return headers;
}

export function modelRequest(apiFormat, model, messages, key) {
  const headers = modelHeaders(apiFormat, key);
  if (apiFormat !== 'anthropic') return { headers, body: { model, messages, stream: false } };
  const system = [];
  const turns = [];
  for (const message of messages) {
    if (message.role === 'system') {
      if (typeof message.content !== 'string')
        throw new ModelProtocolError('Claude 系统指令需要文本。');
      system.push(message.content);
    } else {
      turns.push({ role: message.role, content: anthropicContent(message.content) });
    }
  }
  if (!turns.length) throw new ModelProtocolError('Claude 请求至少需要一条用户消息。');
  return {
    headers,
    body: {
      model,
      ...(system.length ? { system: system.join('\n\n') } : {}),
      messages: turns,
      stream: false,
    },
  };
}

async function protocolJson(response) {
  const reader = response.body?.getReader();
  if (!reader) throw new ModelProtocolError('模型信息接口未返回 JSON。');
  const decoder = new TextDecoder();
  let text = '';
  let bytes = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      bytes += value.byteLength;
      if (bytes > 65_536) {
        void reader.cancel().catch(() => {});
        throw new ModelProtocolError('模型信息接口返回的数据过大。');
      }
      text += decoder.decode(value, { stream: true });
    }
    return JSON.parse(text + decoder.decode());
  } finally {
    reader.releaseLock();
  }
}

function requiresMaxTokens(data) {
  if (
    Array.isArray(data?.detail) &&
    data.detail.length > 0 &&
    data.detail.every(
      (entry) =>
        Array.isArray(entry?.loc) &&
        entry.loc.length === 2 &&
        entry.loc[0] === 'body' &&
        entry.loc[1] === 'max_tokens' &&
        ['missing', 'value_error.missing'].includes(entry.type),
    )
  )
    return true;
  const message = data?.error?.message ?? data?.message;
  if (typeof message !== 'string') return false;
  return [
    /^max_tokens\s+is\s+required\s+and\s+must\s+be\s+positive[.!]?$/i,
    /^missing\s+(?:required\s+)?(?:parameter|field|argument)\s*[:=]?\s*["'`]?max_tokens["'`]?[.!]?$/i,
    /^["'`]?max_tokens["'`]?\s*[:=]?\s*(?:(?:is|was)\s+)?(?:a\s+)?(?:required(?:\s+(?:field|parameter|argument))?|(?:field|parameter|argument)\s+(?:is\s+)?required|missing(?:\s+required\s+(?:field|parameter|argument))?)[.!]?$/i,
    /^(?:field|parameter|argument)\s+(?:required|missing)\s*[:=]?\s*["'`]?max_tokens["'`]?[.!]?$/i,
  ].some((pattern) => pattern.test(message.trim()));
}

function checkAborted(signal) {
  if (signal?.aborted) throw signal.reason ?? new DOMException('Aborted', 'AbortError');
}

function modelInfoEndpoint(endpoint, model) {
  if (!endpoint.endsWith('/messages') || model === '.' || model === '..')
    throw new ModelProtocolError('无法从当前 Claude 接口地址查询模型输出上限。');
  // Derive the lookup from the configured endpoint, never from response URLs.
  return endpoint.slice(0, -'/messages'.length) + '/models/' + encodeURIComponent(model);
}

async function modelOutputLimit(fetcher, endpoint, model, headers, signal) {
  const url = modelInfoEndpoint(endpoint, model);
  try {
    checkAborted(signal);
    const response = await fetcher(url, {
      // The fixed local proxy validates same-origin POSTs and translates this
      // allowlisted model lookup to GET at its configured upstream.
      method: url.startsWith('/model-proxy/') ? 'POST' : 'GET',
      headers,
      signal,
      redirect: 'error',
      credentials: 'omit',
      cache: 'no-store',
    });
    if (!response.ok) {
      await response.body?.cancel();
      throw new ModelProtocolError(
        `Claude 接口要求 max_tokens，但模型信息查询返回 HTTP ${response.status}。请确认服务支持查询该模型的输出上限；若同一服务提供 OpenAI 兼容接口，也可在模型设置中手动选择该格式与地址。`,
      );
    }
    const data = await protocolJson(response);
    if (
      !data ||
      typeof data !== 'object' ||
      !Number.isSafeInteger(data.max_tokens) ||
      data.max_tokens <= 0
    )
      throw new ModelProtocolError(
        'Claude 接口要求 max_tokens，但模型信息未提供有效的 max_tokens 输出上限。请让服务提供该模型的真实上限后重试；若同一服务提供 OpenAI 兼容接口，也可在模型设置中手动选择该格式与地址。',
      );
    checkAborted(signal);
    return data.max_tokens;
  } catch (error) {
    checkAborted(signal);
    if (error instanceof ModelProtocolError) throw error;
    throw new ModelProtocolError(
      'Claude 接口要求 max_tokens，但无法读取同一服务的模型输出上限。请检查模型信息接口、网络和访问权限。',
    );
  }
}

/** Share capability negotiation between browser-direct and native/shared calls. */
export async function sendModelRequest(fetcher, endpoint, apiFormat, model, messages, key, signal) {
  const { headers, body } = modelRequest(apiFormat, model, messages, key);
  const send = (maxTokens) => {
    checkAborted(signal);
    return fetcher(endpoint, {
      method: 'POST',
      signal,
      redirect: 'error',
      credentials: 'omit',
      cache: 'no-store',
      headers,
      body: JSON.stringify({
        ...body,
        ...(maxTokens === undefined ? {} : { max_tokens: maxTokens }),
      }),
    });
  };
  if (apiFormat !== 'anthropic') return send();
  if (!endpoint.startsWith('/') && new URL(endpoint).origin === 'https://api.anthropic.com')
    return send(await modelOutputLimit(fetcher, endpoint, model, headers, signal));
  const response = await send();
  // Never replay a successful (including partially generated) response, or a
  // rejection unrelated to the one required field that we can negotiate.
  if (![400, 422].includes(response.status)) return response;
  let missing = false;
  try {
    missing = requiresMaxTokens(await protocolJson(response));
  } catch {
    checkAborted(signal);
  }
  // Both callers deliberately expose only status/headers for upstream errors.
  // Return an empty error response after inspecting its bounded body; cloning
  // would retain an unread tee when a gateway sends a large or slow error page.
  if (!missing)
    return new Response(null, {
      status: response.status,
      statusText: response.statusText,
      headers: response.headers,
    });
  return send(await modelOutputLimit(fetcher, endpoint, model, headers, signal));
}

const tokenCount = (value) => Number.isSafeInteger(value) && value >= 0;
function anthropicUsage(usage) {
  if (!usage) return undefined;
  const result = {};
  if (tokenCount(usage.input_tokens)) {
    const cached = tokenCount(usage.cache_read_input_tokens) ? usage.cache_read_input_tokens : 0;
    const created = tokenCount(usage.cache_creation_input_tokens)
      ? usage.cache_creation_input_tokens
      : 0;
    result.prompt_tokens = usage.input_tokens + cached + created;
    if (cached) result.prompt_tokens_details = { cached_tokens: cached };
  }
  if (tokenCount(usage.output_tokens)) result.completion_tokens = usage.output_tokens;
  if (tokenCount(result.prompt_tokens) && tokenCount(result.completion_tokens))
    result.total_tokens = result.prompt_tokens + result.completion_tokens;
  return result;
}

export function normalizeModelResponse(apiFormat, data) {
  if (apiFormat !== 'anthropic') return data;
  if (!Array.isArray(data?.content))
    throw new ModelProtocolError('服务未返回 Claude Messages 响应，请检查接口格式和地址。');
  let reason;
  switch (data.stop_reason) {
    case 'end_turn':
    case 'stop_sequence':
      reason = 'stop';
      break;
    case 'max_tokens':
      reason = 'length';
      break;
    case 'model_context_window_exceeded':
      reason = 'context_length';
      break;
    case 'refusal':
      reason = 'content_filter';
      break;
    default:
      throw new ModelProtocolError('Claude 接口未返回完整的最终回复，请检查模型输出和接口格式。');
  }
  const textBlocks = data.content.filter((block) => block?.type === 'text');
  if (textBlocks.some((block) => typeof block.text !== 'string'))
    throw new ModelProtocolError('Claude 接口返回了无效的文本片段。');
  const content = textBlocks.map((block) => block.text).join('');
  if (reason === 'stop' && !content.trim())
    throw new ModelProtocolError('Claude 接口未返回可用的文本内容。');
  return {
    id: data.id,
    model: data.model,
    choices: [{ message: { content }, finish_reason: reason }],
    usage: anthropicUsage(data.usage),
  };
}
