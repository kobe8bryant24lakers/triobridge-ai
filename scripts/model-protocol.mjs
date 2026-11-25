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
      max_tokens: 4096,
      ...(system.length ? { system: system.join('\n\n') } : {}),
      messages: turns,
      stream: false,
    },
  };
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
    case 'model_context_window_exceeded':
      reason = 'length';
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
