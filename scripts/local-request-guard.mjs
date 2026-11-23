/** Reject browser CSRF and DNS rebinding. This is not authentication of other local processes. */
export function acceptsLocalOrigin(req, allowedOrigin) {
  const allowed = new URL(allowedOrigin);
  if (
    allowed.protocol !== 'https:' ||
    !['localhost', '127.0.0.1'].includes(allowed.hostname) ||
    allowed.username ||
    allowed.password ||
    allowed.search ||
    allowed.hash ||
    allowed.pathname !== '/'
  )
    return false;
  const authority = req.headers[':authority'];
  const host = authority ?? req.headers.host;
  const hosts = [`localhost:${allowed.port}`, `127.0.0.1:${allowed.port}`];
  return (
    !(authority && req.headers.host && authority !== req.headers.host) &&
    hosts.includes(host) &&
    hosts.map((item) => `https://${item}`).includes(req.headers.origin)
  );
}
