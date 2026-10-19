// Small, bounded CDP probes. Playwright attaches to every page: one stalled renderer
// can otherwise prevent collection from using the healthy tabs in the same browser.
export function cdpCall(url, method, params = {}, { timeout = 3000, Socket = WebSocket } = {}) {
  return new Promise((resolve, reject) => {
    const socket = new Socket(url);
    const finish = (error, value) => {
      clearTimeout(timer);
      socket.onopen = socket.onmessage = socket.onerror = socket.onclose = null;
      try { socket.close(); } catch {}
      error ? reject(error) : resolve(value);
    };
    const timer = setTimeout(() => finish(new Error(`CDP ${method} timed out`)), timeout);
    socket.onopen = () => socket.send(JSON.stringify({ id: 1, method, params }));
    socket.onmessage = event => {
      const reply = JSON.parse(event.data);
      if (reply.id !== 1) return;
      finish(reply.error ? new Error(`CDP ${method} failed`) : null, reply.result);
    };
    socket.onerror = () => finish(new Error(`CDP ${method} connection failed`));
    socket.onclose = () => finish(new Error(`CDP ${method} disconnected`));
  });
}
export async function responsiveTarget(target, call = cdpCall) {
  if (target.type !== 'page' || !target.webSocketDebuggerUrl) return true;
  try {
    await call(target.webSocketDebuggerUrl, 'Page.getFrameTree');
    await call(target.webSocketDebuggerUrl, 'Runtime.evaluate', { expression: '1', returnByValue: true });
    return true;
  } catch { return false; }
}
export async function recoverTargets(endpoint, { request = fetch, call = cdpCall, application = () => null } = {}) {
  const targets = await (await request(endpoint + '/json/list', { signal: AbortSignal.timeout(3000) })).json();
  const checked = await Promise.all(targets.filter(t => t.type === 'page').map(async target => ({ target, ok: await responsiveTarget(target, call) })));
  const seen = new Set();
  const duplicates = checked.filter(({target,ok})=>{const app=ok&&application(target.url);if(!app)return false;if(seen.has(app))return true;seen.add(app);return false;});
  const stalled = [...checked.filter(t => !t.ok), ...duplicates];
  // Close only the frozen targets in this dedicated browser. No whole-browser kill,
  // cookie/profile deletion, or interference with the user's regular Chrome.
  for (const { target } of stalled) {
    const response = await request(endpoint + '/json/close/' + encodeURIComponent(target.id), { signal: AbortSignal.timeout(3000) });
    if (!response.ok) throw new Error('A frozen Meta-Nav tab could not be closed. Close that tab and click SIGN IN.');
  }
  return stalled.map(({ target }) => target.id);
}
