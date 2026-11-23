/** Explain transport failures without exposing raw errors, headers, or response bodies. */
export function modelConnectionError(error, endpoint) {
  const url = new URL(endpoint);
  const local = ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
  const codes = new Set();
  const seen = new Set();
  const pending = [error];
  while (pending.length && seen.size < 20) {
    const item = pending.shift();
    if (!item || typeof item !== 'object' || seen.has(item)) continue;
    seen.add(item);
    if (typeof item.code === 'string') codes.add(item.code);
    if (item.cause) pending.push(item.cause);
    if (Array.isArray(item.errors)) pending.push(...item.errors.slice(0, 20));
  }
  const has = (...values) => values.some((value) => codes.has(value));
  const target = url.host;
  const localHelp = local
    ? ' 回环地址指运行 TrioBridge 本机服务的电脑；API 在容器或虚拟机中时，请确认端口已映射到这台电脑。'
    : '';
  if (has('ECONNREFUSED'))
    return `模型服务 ${target} 拒绝连接（ECONNREFUSED）。请确认 API 已启动，且监听地址和端口与配置一致。${localHelp}`;
  if (has('ENOTFOUND', 'EAI_AGAIN'))
    return `无法解析模型服务 ${target} 的地址。请检查主机名和 DNS 设置。`;
  if (
    has('ETIMEDOUT', 'UND_ERR_CONNECT_TIMEOUT', 'UND_ERR_HEADERS_TIMEOUT', 'UND_ERR_BODY_TIMEOUT')
  )
    return `连接模型服务 ${target} 超时。请检查 API 是否正常响应，以及网络或防火墙是否允许此连接。${localHelp}`;
  if (has('ECONNRESET', 'EPIPE', 'UND_ERR_SOCKET'))
    return `模型服务 ${target} 的连接被中断。请检查 API 日志，并确认地址使用的 HTTP 或 HTTPS 与服务一致。`;
  if (
    has(
      'CERT_HAS_EXPIRED',
      'DEPTH_ZERO_SELF_SIGNED_CERT',
      'SELF_SIGNED_CERT_IN_CHAIN',
      'UNABLE_TO_VERIFY_LEAF_SIGNATURE',
      'UNABLE_TO_GET_ISSUER_CERT_LOCALLY',
      'ERR_TLS_CERT_ALTNAME_INVALID',
    )
  )
    return `无法验证模型服务 ${target} 的 HTTPS 证书。请检查证书有效期、主机名和 Node.js 的受信任证书配置。`;
  if (has('ERR_SSL_WRONG_VERSION_NUMBER', 'HPE_INVALID_CONSTANT'))
    return `模型服务 ${target} 未返回有效的 HTTP 响应。请确认协议、端口以及 OpenAI 兼容 API 地址正确。`;
  if (has('ENETUNREACH', 'EHOSTUNREACH'))
    return `本机服务无法访问 ${target} 所在的网络。请检查网络连接、路由和防火墙。${localHelp}`;
  return `本机服务无法连接模型服务 ${target}。请检查 API 是否运行、服务地址及 API 日志。${localHelp}`;
}
