// Optional companion exporter. Coalesce clicks while a send is running, then
// retry once with the latest state. Never hold up the local service or a run.
import { spawn } from 'node:child_process';

export function completionHook(script, { launch = spawn, timeout = 180_000 } = {}) {
  let child = null, pending = false;
  const trigger = () => {
    if (!script) return;
    if (child) { pending = true; return; }
    try {
      child = launch(process.execPath, [script], { stdio: 'ignore', windowsHide: true });
      const current = child;
      const timer = setTimeout(() => current.kill(), timeout);
      timer.unref?.();
      let finished = false;
      const finish = () => {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        child = null;
        if (pending) { pending = false; trigger(); }
      };
      current.once('error', finish);
      current.once('exit', finish);
      current.unref?.();
    } catch { child = null; } // periodic exporter retries; collection remains usable
  };
  return trigger;
}
