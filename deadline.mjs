export async function withDeadline(promise, ms, label) {
  let timer;
  try { return await Promise.race([promise, new Promise((_,reject)=>{ timer=setTimeout(()=>reject(new Error(`${label} timed out; collection is incomplete`)),ms); })]); }
  finally { clearTimeout(timer); }
}
