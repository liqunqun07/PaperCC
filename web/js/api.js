// 简易 fetch 封装 + SSE 流式读取
export async function getJSON(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error((await errText(r)) || r.statusText);
  return r.json();
}

export async function postJSON(url, body, method = "POST") {
  const r = await fetch(url, {
    method, headers: { "content-type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (!r.ok) throw new Error((await errText(r)) || r.statusText);
  return r.json();
}

export async function putJSON(url, body) { return postJSON(url, body, "PUT"); }
export async function delJSON(url) {
  const r = await fetch(url, { method: "DELETE" });
  if (!r.ok) throw new Error(r.statusText);
  return r.json();
}

async function errText(r) {
  try { const j = await r.json(); return j.error || j.detail || ""; } catch { return ""; }
}

// POST + SSE。onEvent(obj) 逐事件回调；返回最终是否正常结束
export async function ssePost(url, body, onEvent, signal) {
  const r = await fetch(url, {
    method: "POST", headers: { "content-type": "application/json" },
    body: JSON.stringify(body), signal,
  });
  if (!r.ok) throw new Error((await errText(r)) || r.statusText);
  const reader = r.body.getReader();
  const dec = new TextDecoder();
  let buf = "";
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += dec.decode(value, { stream: true });
    let idx;
    while ((idx = buf.indexOf("\n\n")) >= 0) {
      const chunk = buf.slice(0, idx); buf = buf.slice(idx + 2);
      for (const line of chunk.split("\n")) {
        if (!line.startsWith("data:")) continue;
        try { onEvent(JSON.parse(line.slice(5).trim())); } catch { /* ignore */ }
      }
    }
  }
  return true;
}
