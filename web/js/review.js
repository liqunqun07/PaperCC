// 审阅模式：批注版 .docx 阅读 + 划词/批注/段落翻译 + WPS 联动
import { getJSON, postJSON, ssePost } from "./api.js";
import { renderMD, escapeHtml, toast } from "./utils.js";

const $ = id => document.getElementById(id);
const S = {
  doc: null,            // 当前渲染结果
  view: "markup",
  fs: 15,
  watchTimer: null,
  cmOut: new Map(),     // `${mode}:${批注文本}` → 结果（文档刷新后恢复）
  results: [],
  popAbort: null,
  busy: false,          // 面板写入进行中：暂停外部修改监听
  epOk: null,
  jobs: new Map(),      // jid → 队列任务（含已完成，页面会话内回填）
  jobPoll: null,        // 轮询 timer
  adopted: 0,           // 本次已采用改写数
  lastCids: null,       // {docId, ids:Set} 上一轮批注集合（新批注自动响应用）
  kbIdx: -1,            // 键盘流当前聚焦的批注序号
};

const lsGet = (k, d) => { try { return localStorage.getItem(k) ?? d; } catch { return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, v); } catch { /* ignore */ } };

// ---------- 文件 ----------

async function loadRecent() {
  try {
    const r = await getJSON("/api/review/recent");
    const sel = $("historySel");
    sel.innerHTML = '<option value="">最近…</option>';
    for (const f of r.folders) sel.insertAdjacentHTML("beforeend",
      `<option value="${escapeHtml(f)}">📁 ${escapeHtml(f.split("/").slice(-2).join("/"))}</option>`);
    for (const d of r.docs.filter(d => d.exists)) sel.insertAdjacentHTML("beforeend",
      `<option value="${escapeHtml(d.path)}">📄 ${escapeHtml(d.name)}</option>`);
    return r;
  } catch (e) { toast(e.message, "err"); return { docs: [], folders: [] }; }
}

async function openPath(raw) {
  const p = raw.trim().replace(/^['"]|['"]$/g, "");
  if (!p) return;
  if (/\.docx$/i.test(p)) {
    const folder = p.replace(/\/[^/]+$/, "");
    if (!S.root || !folder.startsWith(S.root)) await setRoot(folder, true);
    expandTo(folder);
    return openDoc(p);
  }
  await setRoot(p);
}

// ---------- 左栏：目录树（映射访达里的文件夹，每 5 秒同步） ----------

const fmtTime = t => new Date(t * 1000).toLocaleString("zh-CN", { hour12: false, month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });

function loadExpanded() {
  try { return new Set(JSON.parse(lsGet("rv-expanded", "[]"))); } catch { return new Set(); }
}
S.expanded = loadExpanded();
S.dirs = new Map();      // 路径 → 该层内容

function saveExpanded() { lsSet("rv-expanded", JSON.stringify([...S.expanded])); }

async function fetchDir(path) {
  const r = await getJSON("/api/review/dir?path=" + encodeURIComponent(path));
  S.dirs.set(r.path, r);
  return r;
}

async function setRoot(folder, quiet = false) {
  const list = $("fileList");
  try {
    const r = await fetchDir(folder);
    S.root = r.path;
    lsSet("rv-root", r.path);
    $("folderInput").value = r.path;
    $("rootName").textContent = r.name;
    $("rootName").title = r.path;
    $("upBtn").disabled = !r.parent;
    await syncTree();
    renderTree();   // 换根后必须重绘（内容已预先读入，syncTree 判断为无变化）
    loadRecent();
  } catch (e) {
    list.innerHTML = `<div class="hint">⚠️ ${escapeHtml(e.message)}</div>`;
    if (!quiet) toast(e.message, "err");
  }
}

function expandTo(folder) {
  if (!S.root) return;
  let cur = folder;
  while (cur.startsWith(S.root) && cur.length > S.root.length) {
    S.expanded.add(cur);
    cur = cur.replace(/\/[^/]+$/, "");
  }
  saveExpanded();
  syncTree();
}

// 重新读取根目录与所有已展开的子目录，有变化才重绘
async function syncTree() {
  if (!S.root) return;
  const before = JSON.stringify([...S.dirs.entries()]);
  const need = [S.root, ...[...S.expanded].filter(p => p.startsWith(S.root + "/"))];
  await Promise.all(need.map(p => fetchDir(p).catch(() => { S.expanded.delete(p); S.dirs.delete(p); })));
  if (JSON.stringify([...S.dirs.entries()]) !== before || !$("fileList").querySelector(".tnode")) renderTree();
}

function renderTree() {
  const out = [];
  const walk = (path, lv) => {
    const d = S.dirs.get(path);
    if (!d) return;
    for (const e of d.entries) {
      if (e.type === "dir") {
        const open = S.expanded.has(e.path);
        out.push(`<div class="tnode dir" style="--lv:${lv}" data-dir="${escapeHtml(e.path)}" title="${escapeHtml(e.path)}">
          <span class="tw">${open ? "▼" : "▶"}</span><span class="tn">${escapeHtml(e.name)}</span></div>`);
        if (open) walk(e.path, lv + 1);
      } else {
        out.push(`<div class="tnode file${S.doc?.id === e.id ? " on" : ""}" style="--lv:${lv}" data-path="${escapeHtml(e.path)}" data-id="${e.id}" title="${escapeHtml(e.path)}">
          <span class="tw"></span><span class="tn">${escapeHtml(e.name)}</span><span class="tm">${fmtTime(e.mtime)}</span></div>`);
      }
    }
  };
  walk(S.root, 0);
  const list = $("fileList");
  const top = list.scrollTop;
  list.innerHTML = out.join("") || '<div class="hint">这个文件夹是空的</div>';
  list.scrollTop = top;
}

function wireTree() {
  $("fileList").addEventListener("click", async e => {
    const node = e.target.closest(".tnode");
    if (!node) return;
    if (node.dataset.dir) {
      const p = node.dataset.dir;
      if (S.expanded.has(p)) S.expanded.delete(p); else S.expanded.add(p);
      saveExpanded();
      await syncTree();
      renderTree();
    } else openDoc(node.dataset.path);
  });
  $("upBtn").addEventListener("click", () => {
    const d = S.dirs.get(S.root);
    if (d?.parent) { S.expanded.add(S.root); saveExpanded(); setRoot(d.parent); }
  });
  $("treeRefresh").addEventListener("click", () => syncTree());
  setInterval(() => { if (!document.hidden) syncTree(); }, 5000);
}

async function pickPath(kind) {
  const btn = $("pickBtn");
  btn.disabled = true;
  try {
    const r = await postJSON("/api/review/pick", { kind });
    if (r.path) await openPath(r.path);
  } catch (e) { toast(e.message, "err"); }
  finally { btn.disabled = false; }
}

async function openDoc(path) {
  try {
    const entry = await postJSON("/api/review/open", { path });
    lsSet("rv-doc", path);
    await loadDoc(entry.id);
    loadRecent();
  } catch (e) { toast(e.message, "err"); }
}

async function loadDoc(id, keepScroll = false) {
  const scroller = $("docScroll");
  const ratio = keepScroll ? scroller.scrollTop / Math.max(1, scroller.scrollHeight) : 0;
  const doc = await getJSON(`/api/review/doc/${id}`);
  S.doc = doc;
  $("docEmpty").classList.add("hidden");
  const page = $("docPage");
  page.innerHTML = doc.html;
  page.classList.remove("hidden");
  $("docBar").classList.remove("hidden");
  $("docTitle").textContent = doc.name;
  $("docTitle").title = doc.path;
  const st = doc.stats;
  $("docStats").textContent = `批注 ${st.comments} · 插入 ${st.ins} · 删除 ${st.del} · 表 ${st.tables} · 图 ${st.images}`;
  $("wordBtn").disabled = false;
  $("trAllBtn").disabled = !doc.comments.length;
  $("queueAllBtn").disabled = !doc.comments.length;
  $("exportRespBtn").disabled = !doc.comments.length;
  document.querySelectorAll("#fileList .tnode.file").forEach(f => f.classList.toggle("on", f.dataset.id === id));
  renderComments();
  diffComments();
  renderQbar();
  scroller.scrollTop = keepScroll ? ratio * scroller.scrollHeight : 0;
  startWatch();
}

// ---------- 文件变动监听（Word 保存后自动刷新） ----------

function startWatch() {
  clearInterval(S.watchTimer);
  S.watchTimer = setInterval(async () => {
    if (!S.doc || document.hidden || S.busy) return;
    try {
      const r = await getJSON(`/api/review/doc/${S.doc.id}/mtime`);
      if (r.mtime !== S.doc.mtime) {
        await loadDoc(S.doc.id, true);
        toast("文档已在外部保存，已刷新", "ok");
      }
    } catch { /* 文件暂时被 Word 锁写时忽略 */ }
  }, 3000);
}

// ---------- 批注 ----------

function renderComments() {
  const list = $("commentList");
  const cms = S.doc?.comments || [];
  $("cmCount").textContent = cms.length ? `(${cms.length})` : "";
  if (!cms.length) { list.innerHTML = '<div class="hint">这份文档没有批注。</div>'; return; }
  // 回复紧跟在父批注之后
  // 按正文中首次出现的位置排序
  const pos = new Map();
  document.querySelectorAll("#docPage .cm, #docPage .cm-ref").forEach((el, i) => {
    for (const id of (el.dataset.c || el.dataset.ref || "").split(" ")) if (id && !pos.has(id)) pos.set(id, i);
  });
  const roots = cms.filter(c => !c.parent || !cms.some(x => x.id === c.parent))
    .sort((a, b) => (pos.get(a.id) ?? 1e9) - (pos.get(b.id) ?? 1e9));
  const ordered = [];
  for (const r of roots) {
    ordered.push(r);
    for (const c of cms) if (c.parent === r.id) ordered.push(c);
  }
  const hideDone = $("hideDone").checked;
  list.innerHTML = ordered.filter(c => !(hideDone && c.done)).map((c, i) => `
    <div class="cmt${c.parent ? " reply" : ""}${c.done ? " done" : ""}" data-cid="${c.id}">
      <div class="cmt-head"><span><b>${escapeHtml(c.author || "匿名")}</b>${c.parent ? " · 回复" : ""}${c.done ? " · 已解决" : ""}</span>
        <span>${escapeHtml(fmtDate(c.date))}</span></div>
      ${c.anchor && !c.parent ? `<div class="cmt-anchor" title="定位到正文">${escapeHtml(c.anchor)}</div>` : ""}
      <div class="cmt-text">${escapeHtml(c.text)}</div>
      <div class="cmt-acts">
        <button class="btn small" data-m="translate">翻译</button>
        <button class="btn small" data-m="explain">解释要改什么</button>
        ${c.parent ? "" : `<button class="btn small" data-r="1" title="按批注意见 AI 改写锚定的文字，预览后可一键写入">AI 改写</button>
        <button class="btn small" data-q="rewrite" title="加入后台改写队列，继续看其他批注">排队改写</button>
`}
        <button class="btn small" data-q="reply" title="起草这条批注的审稿回复">回复草稿</button>
        <button class="btn small ghost" data-m="copy">复制</button>
      </div>
      <div class="cmt-acts">
        <button class="btn small" data-e="reply">回复</button>
        <button class="btn small" data-e="resolve">${c.done ? "重新打开" : "✓ 已解决"}</button>
        ${c.parent ? "" : `<button class="btn small ghost" data-e="acceptRange" title="接受该批注范围内的修订">接受范围修订</button>
        <button class="btn small ghost" data-e="rejectRange" title="拒绝该批注范围内的修订">拒绝</button>`}
        <button class="btn small ghost danger" data-e="delete">删除</button>
      </div>
      <div class="cmt-outs"></div>
    </div>`).join("");
  list.querySelectorAll(".cmt").forEach(card => {
    const c = cms.find(x => x.id === card.dataset.cid);
    card.querySelector(".cmt-anchor")?.addEventListener("click", () => jumpToComment(c.id));
    card.addEventListener("click", e => { if (!e.target.closest("button") && !window.getSelection().toString()) focusComment(c.id, false); });
    card.querySelectorAll(".cmt-acts button[data-m]").forEach(b => b.addEventListener("click", () => {
      if (b.dataset.m === "copy") copy(c.text);
      else runComment(c, b.dataset.m, card);
    }));
    card.querySelectorAll(".cmt-acts button[data-q]").forEach(b => b.addEventListener("click", () => queueOp(b.dataset.q, c)));
    card.querySelector("button[data-r]")?.addEventListener("click", () => aiRewrite(c, card));
    card.querySelectorAll(".cmt-acts button[data-e]").forEach(b => b.addEventListener("click", () => commentAction(c, b.dataset.e, b)));
    for (const m of ["translate", "explain"]) {
      const saved = S.cmOut.get(`${m}:${c.text}`);
      if (saved) showCommentOut(card, m, saved, false);
    }
    const savedRw = S.cmOut.get(`rewrite:${c.text}`);
    if (savedRw) restoreRewrite(c, card, savedRw);
    // 队列任务盒（排队/处理中/完成）在卡片重建后恢复
    for (const j of S.jobs.values()) if (j.cid === c.id && j.doc_id === S.doc.id) jobBoxRender(j);
  });
}

// 带 Z 的 UTC 时间转本地显示（Word / WPS 同样按本地时区显示）
function fmtDate(d) {
  if (!d) return "";
  if (/Z$/.test(d)) {
    const t = new Date(d);
    if (!isNaN(t)) return t.toLocaleString("zh-CN", { hour12: false }).slice(0, 16).replace(/\//g, "-");
  }
  return d.slice(0, 16).replace("T", " ");
}

function showCommentOut(card, mode, text, typing) {
  const outs = card.querySelector(".cmt-outs");
  let box = outs.querySelector(`[data-m="${mode}"]`);
  if (!box) {
    box = document.createElement("div");
    box.className = "cmt-out"; box.dataset.m = mode;
    box.innerHTML = `<div class="lbl">${mode === "translate" ? "译文" : "解读"}</div><div class="md"></div>`;
    outs.append(box);
  }
  const md = box.querySelector(".md");
  md.innerHTML = renderMD(text);
  md.classList.toggle("typing", typing);
  return md;
}

// 首字返回前显示等待计时（推理模型会先思考十几秒，避免看起来像卡住）
function waitHint(el, label) {
  const t0 = Date.now();
  const draw = () => { el.innerHTML = `<span class="wait-hint">${label}… 模型思考 ${Math.round((Date.now() - t0) / 1000)} 秒</span>`; };
  draw();
  const timer = setInterval(draw, 1000);
  return () => clearInterval(timer);
}

async function runComment(c, mode, card) {
  const key = `${mode}:${c.text}`;
  const anchorEl = firstSpan(c.id);
  const context = anchorEl ? paraText(anchorEl.closest("[data-p]")) : "";
  let acc = "", err = null;
  const stop = waitHint(showCommentOut(card, mode, "", true), mode === "translate" ? "翻译中" : "解读中");
  try {
    await ssePost("/api/review/comment", { doc_id: S.doc.id, cid: c.id, mode, context }, ev => {
      if (ev.delta) { stop(); acc += ev.delta; showCommentOut(card, mode, acc, true); }
      if (ev.error) err = ev.error;
    });
    stop();
    if (err) throw new Error(err);
    S.cmOut.set(key, acc);
    showCommentOut(card, mode, acc, false);
  } catch (e) {
    stop();
    showCommentOut(card, mode, `⚠️ ${e.message}`, false);
  }
}

// ---------- 批注 AI 改写（按批注意见改写锚定文字，采用后以修订写入） ----------

// 批注 → 锚定区间的写回参数（偏移与后端文字模型一致：含删除文字，不含 💬 / 脚注号）
function commentAnchorRange(c) {
  const spans = [...document.querySelectorAll("#docPage .cm")]
    .filter(s => s.dataset.c.split(" ").includes(c.id));
  // 零宽锚点（批注只打点没选字，正文只剩 💬）：退化为改所在整段
  if (!spans.length) {
    const ref = document.querySelector(`#docPage .cm-ref[data-ref="${c.id}"]`);
    const pEl = ref?.closest("[data-p]");
    if (!pEl) return { error: "正文中找不到该批注的锚定文字" };
    const ptext = rawParaText(pEl);
    return { pEl, ptext, so: 0, eo: ptext.length, target: paraText(pEl) };
  }
  const pEl = spans[0].closest("[data-p]");
  if (!pEl) return { error: "正文中找不到该批注的锚定文字" };
  if (spans.some(s => s.closest("[data-p]") !== pEl)) {
    const range = document.createRange();
    range.setStartBefore(spans[0]);
    range.setEndAfter(spans[spans.length - 1]);
    const anchor = rangeAnchor(range);
    if (!anchor) return { error: "无法定位跨段批注" };
    const holder = document.createElement('div');
    holder.append(range.cloneContents());
    holder.querySelectorAll('p,h1,h2,h3,h4,h5,h6,td').forEach(b => b.append('\n'));
    return { pEl, anchor, target: cleanFragment(holder) };
  }
  const ptext = rawParaText(pEl);
  const inP = spans.filter(s => s.closest("[data-p]") === pEl);
  let so = textUpTo(pEl, inP[0], 0).length;
  const last = inP[inP.length - 1], lp = last.parentNode;
  let eo = textUpTo(pEl, lp, [...lp.childNodes].indexOf(last) + 1).length;
  // 零宽锚点（批注只打点没选字）：改整段
  if (!(eo > so && eo <= ptext.length)) { so = 0; eo = ptext.length; }
  const r = document.createRange();
  r.setStartBefore(inP[0]);
  r.setEndAfter(last);
  const h = document.createElement("div");
  h.append(r.cloneContents());
  // 给模型的文字按当前视图口径（默认不含删除线文字）；写回仍按含删除文字的原始偏移
  const target = cleanFragment(h) || paraText(pEl);
  return { pEl, ptext, so, eo, target };
}

function aiRewriteBox(card, key, text, payload, actions) {
  const outs = card.querySelector(".cmt-outs");
  let box = outs.querySelector(".cmt-out[data-m=rewrite]");
  if (!box) {
    box = document.createElement("div");
    box.className = "cmt-out"; box.dataset.m = "rewrite";
    box.innerHTML = `<div class="lbl">AI 改写（采用后以修订痕迹写入：原文删除线 + 新文）</div><div class="md"></div>
      <div class="cmt-acts rw-acts hidden">
        <button class="btn small" data-rw="apply">✓ 采用</button>
        <button class="btn small ghost" data-rw="cmp">对照</button>
        <button class="btn small ghost" data-rw="discard">放弃</button>
        <button class="btn small ghost" data-rw="copy">复制</button>
      </div>`;
    outs.append(box);
    const md = box.querySelector(".md");
    box.querySelectorAll("[data-rw]").forEach(b => b.addEventListener("click", async () => {
      const act = b.dataset.rw;
      if (act === "discard") { S.cmOut.delete(`rewrite:${box._ckey}`); box.remove(); return; }
      if (act === "copy") return copy(md._text || box._payload?.text || md.textContent);
      if (act === "cmp") {
        const p = box._payload;
        const src = p ? (p.source ?? p.ptext?.slice(p.s, p.e) ?? "") : "";
        return toggleCompare(box, src, md._text || "");
      }
      b.disabled = true;
      S.cmOut.delete(`rewrite:${box._ckey}`);   // 先删缓存：写入成功后刷新重建时不再恢复
      const p = box._payload;
      const res = await edit("rewrite", p.anchor ? { ...p.anchor, text: p.text } : { p: p.p, s: p.s, e: p.e, ptext: p.ptext, text: p.text },
        "已采用 AI 改写（以修订写入，可撤销）");
      if (!res) b.disabled = false;
      else { S.adopted++; renderQbar(); }
    }));
  }
  const md = box.querySelector(".md");
  box._ckey = key;
  box._payload = payload;
  md._text = text;
  md.innerHTML = renderMD(text);
  box.querySelector(".rw-acts").classList.toggle("hidden", !actions);
  return box;
}

// 文档刷新后恢复预览；原文段落已变（ptext 对不上）就丢弃，防止采用时错位
function restoreRewrite(c, card, raw) {
  try {
    const p = JSON.parse(raw);
    const cur = commentAnchorRange(c);
    if (cur.error || (p.anchor ? JSON.stringify(cur.anchor) !== JSON.stringify(p.anchor) : cur.ptext !== p.ptext)) { S.cmOut.delete(`rewrite:${c.text}`); return; }
    aiRewriteBox(card, `rewrite:${c.text}`, p.text, p, true);
  } catch { S.cmOut.delete(`rewrite:${c.text}`); }
}

async function aiRewrite(c, card) {
  const a = commentAnchorRange(c);
  if (a.error) { toast(a.error, "err"); return; }
  const key = `rewrite:${c.text}`;
  const old = card.querySelector(".cmt-out[data-m=rewrite]");
  old?._ctrl?.abort();          // 再点一次 = 重新生成，停掉上一轮流
  const box = aiRewriteBox(card, key, "", null, false);
  const md = box.querySelector(".md");
  const stop = waitHint(md, "AI 改写中");
  const ctrl = new AbortController();
  box._ctrl = ctrl;
  let acc = "", err = null;
  try {
    await ssePost("/api/review/comment/rewrite",
      { doc_id: S.doc.id, cid: c.id, text: a.target, context: rewriteContext(a) },
      ev => {
        if (ev.delta) { stop(); acc += ev.delta; md.innerHTML = renderMD(acc); }
        if (ev.error) err = ev.error;
      }, ctrl.signal);
    stop();
    if (err) throw new Error(err);
    const text = acc.trim();
    if (!text) throw new Error("模型未返回内容");
    const payload = a.anchor ? { anchor: a.anchor, source: a.target, text } : { p: +a.pEl.dataset.p, s: a.so, e: a.eo, ptext: a.ptext, text };
    S.cmOut.set(key, JSON.stringify(payload));
    aiRewriteBox(card, key, text, payload, true);
  } catch (e) {
    stop();
    md.innerHTML = e.name === "AbortError" ? "（已取消）" : `⚠️ ${escapeHtml(e.message)}`;
  }
}

function rewriteContext(a) {
  const context = paraText(a.pEl);
  return a.anchor ? `${context}\n\n跨段改写必须保留原段落数，每段用一个换行分隔，不得合并或新增段落。` : context;
}

// ---------- 批注任务队列（后台处理：排队改写 / 回复草稿 / 深度改写，边看边出结果） ----------

const OP_LABEL = { explain: "解读", translate: "译文", rewrite: "AI 改写", reply: "回复草稿" };
const JOB_ACT = new Set(["queued", "running", "cancelling"]);

async function enqueueJobs(items) {
  if (!S.doc || !items.length) return;
  try {
    const r = await postJSON("/api/review/jobs", { doc_id: S.doc.id, items });
    for (const j of r.jobs) { S.jobs.set(j.id, j); jobBoxRender(j); }
    for (const s of r.skipped || []) toast(`${OP_LABEL[s.op] || s.op}：${s.reason}`, "err");
    if (r.jobs.length) { startJobPoll(); renderQbar(); }
  } catch (e) { toast(e.message, "err"); }
}

async function queueOp(op, c) {
  if (!c) return;
  const item = { cid: c.id, op };
  if (op === "rewrite" || op === "reply") {
    const a = commentAnchorRange(c);
    if (a.error) { toast(a.error, "err"); return; }
    item.text = a.target;
    item.context = rewriteContext(a);
  } else {
    const anchorEl = firstSpan(c.id);
    item.text = c.anchor || "";
    item.context = anchorEl ? paraText(anchorEl.closest("[data-p]")) : "";
  }
  await enqueueJobs([item]);
}

function startJobPoll() {
  if (S.jobPoll) return;
  S.jobPoll = setInterval(async () => {
    clearInterval(S.jobPoll); S.jobPoll = null;      // 防重入：每轮重新挂
    if (!S.doc || document.hidden) { renderQbar(); return; }
    try {
      const r = await getJSON(`/api/review/jobs?doc_id=${S.doc.id}`);
      let active = false, changed = false;
      for (const j of r.jobs) {
        const old = S.jobs.get(j.id);
        if (!old || old.status !== j.status || old.result !== j.result) changed = true;
        S.jobs.set(j.id, j);
        if (JOB_ACT.has(j.status)) active = true;
      }
      if (changed) for (const j of r.jobs) jobBoxRender(j);
      if (active) startJobPoll(); else renderQbar();
    } catch { renderQbar(); }
  }, 1500);
}

function jobBoxRender(j) {
  const card = document.querySelector(`.cmt[data-cid="${j.cid}"]`);
  if (!card) return;
  // explain/translate 的完成结果并入缓存恢复路径，避免同一内容出现两份
  if ((j.op === "explain" || j.op === "translate") && j.status === "done") {
    S.cmOut.set(`${j.op}:${j.comment_text}`, j.result);
    if (!card.querySelector(`.cmt-out[data-m="${j.op}"]`)) showCommentOut(card, j.op, j.result, false);
    renderQbar(); return;
  }
  const outs = card.querySelector(".cmt-outs");
  let box = outs.querySelector(`.cmt-out[data-jid="${j.id}"]`);
  if (!box) {
    box = document.createElement("div");
    box.className = "cmt-out job"; box.dataset.jid = j.id;
    outs.append(box);
  }
  const label = OP_LABEL[j.op] || j.op;
  const hedgeNote = j.hedges && Object.keys(j.hedges).length
    ? `<div class="hedge-warn">⚠️ 改写新增限定词：${Object.entries(j.hedges).map(([k, v]) => `${escapeHtml(k)}×${v}`).join("、")}——按你的红线默认不应新增，可重试或自行核对后采用</div>` : "";
  if (j.status === "queued" || j.status === "cancelling") {
    box.innerHTML = `<div class="lbl">⏳ ${label}排队中…</div>`;
  } else if (j.status === "running") {
    box.innerHTML = `<div class="lbl">⚙️ ${label}处理中…</div>`;
  } else if (j.status === "error") {
    box.innerHTML = `<div class="lbl">⚠️ ${label}失败：${escapeHtml(j.error || "")}</div>
      <div class="cmt-acts"><button class="btn small" data-jr="retry">重试</button><button class="btn small ghost" data-jr="dismiss">移除</button></div>`;
  } else if (j.status === "done") {
    const md = `<div class="md">${renderMD(j.result)}</div>`;
    if (j.op === "reply") {
      box.innerHTML = `<div class="lbl">✍️ 回复草稿${j.cached ? "（缓存）" : ""}</div>${md}
        <div class="cmt-acts"><button class="btn small" data-jr="copy">复制</button><button class="btn small ghost" data-jr="cmp">对照</button><button class="btn small ghost" data-jr="retry">重新生成</button><button class="btn small ghost" data-jr="dismiss">丢弃</button></div>`;
    } else {
      box.innerHTML = `<div class="lbl">✨ ${label}${j.cached ? "（缓存）" : ""}（采用后以修订痕迹写入）</div>${md}${hedgeNote}
        <div class="cmt-acts"><button class="btn small" data-jr="apply">✓ 采用</button><button class="btn small ghost" data-jr="cmp">对照</button><button class="btn small ghost" data-jr="discard">放弃</button><button class="btn small ghost" data-jr="copy">复制</button></div>`;
    }
  }
  box.querySelectorAll("[data-jr]").forEach(b => b.addEventListener("click", async () => {
    const act = b.dataset.jr;
    if (act === "copy") return copy(j.result);
    if (act === "dismiss") { S.jobs.delete(j.id); box.remove(); renderQbar(); return; }
    if (act === "cmp") return toggleCompare(box, j.target, j.result);
    if (act === "retry") {
      S.jobs.delete(j.id); box.remove();
      await queueOp(j.op, S.doc?.comments.find(x => x.id === j.cid));
      return;
    }
    if (act === "apply") {
      const c = S.doc?.comments.find(x => x.id === j.cid);
      if (!c) return;
      const a = commentAnchorRange(c);
      if (a.error) { toast(a.error, "err"); return; }
      b.disabled = true;
      if (a.target !== j.target) { toast("批注原文已变化，请重新生成改写", "err"); b.disabled = false; return; }
      const res = await edit("rewrite", a.anchor ? { ...a.anchor, text: j.result } : { p: +a.pEl.dataset.p, s: a.so, e: a.eo, ptext: a.ptext, text: j.result },
        `已采用${OP_LABEL[j.op] || "改写"}（以修订写入，可撤销）`);
      if (res) { S.adopted++; S.jobs.delete(j.id); box.remove(); renderQbar(); }
      else b.disabled = false;
    }
  }));
  renderQbar();
}

function renderQbar() {
  const bar = $("qbar");
  if (!bar) return;
  let queued = 0, running = 0, done = 0, err = 0;
  for (const j of S.jobs.values()) {
    if (j.doc_id !== S.doc?.id) continue;
    if (JOB_ACT.has(j.status)) j.status === "running" ? running++ : queued++;
    else if (j.status === "done") done++;
    else if (j.status === "error") err++;
  }
  const cms = S.doc?.comments || [];
  if (!cms.length && !done && !queued && !running && !err) { bar.classList.add("hidden"); return; }
  bar.classList.remove("hidden");
  bar.innerHTML = `<span>批注 ${cms.length}${cms.length ? ` · 已解决 ${cms.filter(c => c.done).length}` : ""}</span>`
    + (running || queued ? `<span> · ⚙️ ${running} ⏳ ${queued}</span><button id="qbarCancel" class="btn small ghost">取消队列</button>` : "")
    + (done ? `<span> · 已完成 ${done}</span>` : "")
    + (err ? `<span class="qbar-err"> · 失败 ${err}</span>` : "")
    + (S.adopted ? `<span> · 本次已采用 ${S.adopted}</span>` : "");
  $("qbarCancel")?.addEventListener("click", async () => {
    const ids = [...S.jobs.values()].filter(j => j.doc_id === S.doc.id && JOB_ACT.has(j.status)).map(j => j.id);
    try { await postJSON("/api/review/jobs/cancel", { ids }); } catch (e) { toast(e.message, "err"); }
  });
}

// 新批注自动响应：文档刷新后 diff 批注集合，新增的按开关自动入队
function diffComments() {
  const cms = S.doc?.comments || [];
  const ids = new Set(cms.map(c => c.id));
  const last = S.lastCids;
  const sameDoc = last && last.docId === S.doc.id;
  S.lastCids = { docId: S.doc.id, ids };
  if (!sameDoc || S.busy) return;                  // 首开定基线；面板写入流程不触发
  const fresh = [...ids].filter(x => !last.ids.has(x))
    .map(id => cms.find(c => c.id === id)).filter(c => c && !c.parent);
  if (!fresh.length) return;
  const explainOn = lsGet("rv-auto-explain", "1") === "1";
  const rewriteOn = lsGet("rv-auto-rewrite", "0") === "1";
  if (!explainOn && !rewriteOn) return;
  const cap = 15;
  const list = fresh.slice(0, cap);
  const items = [];
  for (const c of list) {
    if (explainOn) {
      const anchorEl = firstSpan(c.id);
      items.push({ cid: c.id, op: "explain", text: c.anchor || "",
        context: anchorEl ? paraText(anchorEl.closest("[data-p]")) : "" });
    }
    if (rewriteOn) {
      const a = commentAnchorRange(c);
      if (!a.error && a.target) items.push({ cid: c.id, op: "rewrite", text: a.target, context: rewriteContext(a) });
    }
  }
  if (!items.length) return;
  enqueueJobs(items).then(() =>
    toast(`新批注 ${list.length} 条已自动入队${fresh.length > cap ? `（上限截断 ${fresh.length - cap} 条）` : ""}`, "ok"));
}

// 批注回复稿汇总导出（md + docx 到 PaperCC 数据目录）
async function exportResponses() {
  if (!S.doc) return;
  const items = [];
  for (const c of S.doc.comments.filter(x => !x.parent)) {
    let response = "";
    for (const j of S.jobs.values())
      if (j.cid === c.id && j.op === "reply" && j.status === "done") response = j.result;
    if (response) items.push({ comment: c.text, anchor: c.anchor || "", response });
  }
  if (!items.length) { toast("还没有回复草稿：先在批注上点「回复草稿」生成", "err"); return; }
  try {
    const r = await postJSON("/api/review/responses/export", { doc_name: S.doc.name, items });
    toast(`已导出 ${r.count} 条回复：${r.docx}`, "ok");
    copy(r.docx);
  } catch (e) { toast(e.message, "err"); }
}

// ---------- 键盘流（J/K 切换批注，E/W/Q/R 触发动作；输入框内不生效） ----------

function wireKeyboard() {
  document.addEventListener("keydown", e => {
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    const t = e.target;
    if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.isContentEditable)) return;
    if (!$("editPop").classList.contains("hidden") || !$("selPop").classList.contains("hidden")) return;
    const cards = [...document.querySelectorAll("#commentList .cmt")];
    if (!cards.length) return;
    const k = e.key.toLowerCase();
    if (k === "j" || k === "k") {
      e.preventDefault();
      S.kbIdx = k === "j" ? Math.min(cards.length - 1, S.kbIdx + 1) : Math.max(0, S.kbIdx - 1);
      const card = cards[S.kbIdx];
      const c = S.doc?.comments.find(x => x.id === card.dataset.cid);
      if (c) { focusComment(c.id); card.scrollIntoView({ block: "nearest", behavior: "smooth" }); }
      return;
    }
    if (S.kbIdx < 0 || S.kbIdx >= cards.length) return;
    const card = cards[S.kbIdx];
    const c = S.doc?.comments.find(x => x.id === card.dataset.cid);
    if (!c) return;
    if (k === "e") { e.preventDefault(); runComment(c, "explain", card); }
    else if (k === "w") { e.preventDefault(); aiRewrite(c, card); }
    else if (k === "q") { e.preventDefault(); queueOp("rewrite", c); }
    else if (k === "r") { e.preventDefault(); queueOp("reply", c); }
  });
}

// ---------- 采用前对照视图（原文 / 改写 并排 + 词级 diff 高亮，超长自动降级） ----------

function diffTokens(A, B) {
  const n = A.length, m = B.length, w = m + 1;
  const dp = new Uint32Array((n + 1) * w);
  for (let i = n - 1; i >= 0; i--)
    for (let j = m - 1; j >= 0; j--)
      dp[i * w + j] = A[i] === B[j] ? dp[(i + 1) * w + j + 1] + 1 : Math.max(dp[(i + 1) * w + j], dp[i * w + j + 1]);
  const runs = [];
  let i = 0, j = 0;
  const push = (t, c) => {
    if (!c) return;
    if (runs.length && runs[runs.length - 1].t === t) runs[runs.length - 1].c += c;
    else runs.push({ t, c });
  };
  while (i < n && j < m) {
    if (A[i] === B[j]) { push(" ", A[i]); i++; j++; }
    else if (dp[(i + 1) * w + j] >= dp[i * w + j + 1]) { push("-", A[i]); i++; }
    else { push("+", B[j]); j++; }
  }
  while (i < n) push("-", A[i++]);
  while (j < m) push("+", B[j++]);
  const html = runs.map(r => r.t === " " ? escapeHtml(r.c)
    : r.t === "-" ? `<del class="df">${escapeHtml(r.c)}</del>` : `<ins class="df">${escapeHtml(r.c)}</ins>`).join("");
  return {
    aHtml: html.replace(/<ins class="df">[\s\S]*?<\/ins>/g, ""),
    bHtml: html.replace(/<del class="df">[\s\S]*?<\/del>/g, ""),
  };
}

function toggleCompare(container, src, out) {
  const old = container.querySelector(".rw-cmp");
  if (old) { old.remove(); return; }
  if (!src || src === out) { toast(src ? "原文与改写相同" : "该结果没有原文可对照", "err"); return; }
  const tok = s => s.match(/[\u4e00-\u9fff]|[A-Za-z0-9']+|\s+|[^\s]/g) || [];
  const A = tok(src), B = tok(out);
  let aHtml = escapeHtml(src), bHtml = escapeHtml(out);
  if (A.length * B.length <= 360000) ({ aHtml, bHtml } = diffTokens(A, B));
  container.insertAdjacentHTML("beforeend",
    `<div class="rw-cmp"><div class="rw-col"><div class="lbl">原文（删除线=改写删掉的字）</div><div class="rw-txt">${aHtml}</div></div>
     <div class="rw-col"><div class="lbl">改写（下划线=新增的字）</div><div class="rw-txt">${bHtml}</div></div></div>`);
}

async function translateAll() {
  const btn = $("trAllBtn");
  btn.disabled = true;
  const cards = [...document.querySelectorAll("#commentList .cmt")];
  const todo = cards.filter(card => {
    const c = S.doc.comments.find(x => x.id === card.dataset.cid);
    return c && !S.cmOut.has(`translate:${c.text}`);
  });
  let done = 0;
  const worker = async () => {
    while (todo.length) {
      const card = todo.shift();
      const c = S.doc.comments.find(x => x.id === card.dataset.cid);
      await runComment(c, "translate", card);
      btn.textContent = `翻译中 ${++done}/${done + todo.length}`;
    }
  };
  await Promise.all([worker(), worker()]);
  btn.textContent = "全部翻译";
  btn.disabled = false;
}

const firstSpan = cid => [...document.querySelectorAll("#docPage .cm")].find(s => s.dataset.c.split(" ").includes(cid));

function focusComment(cid, scrollCard = true) {
  document.querySelectorAll("#docPage .cm.on").forEach(s => s.classList.remove("on"));
  document.querySelectorAll("#docPage .cm").forEach(s => { if (s.dataset.c.split(" ").includes(cid)) s.classList.add("on"); });
  document.querySelectorAll(".cmt").forEach(c => c.classList.toggle("on", c.dataset.cid === cid));
  if (scrollCard) {
    switchTab("comments");
    document.querySelector(`.cmt[data-cid="${cid}"]`)?.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }
}

function jumpToComment(cid) {
  focusComment(cid, false);
  const span = firstSpan(cid) || document.querySelector(`#docPage .cm-ref[data-ref="${cid}"]`);
  if (!span) { toast("正文中找不到该批注的锚点", "err"); return; }
  span.scrollIntoView({ block: "center", behavior: "smooth" });
  const p = span.closest("[data-p]") || span;
  p.classList.remove("flash"); void p.offsetWidth; p.classList.add("flash");
}

// ---------- 文本提取（按当前修订视图口径） ----------

function cleanFragment(frag, keepDel = false) {
  const drop = [".cm-ref", ".p-tr", ".fn-ref"];
  if (!keepDel && S.view !== "original") drop.push("del.rev-del", ".para-del");
  if (S.view === "original") drop.push("ins.rev-ins");
  frag.querySelectorAll(drop.join(",")).forEach(n => n.remove());
  return frag.textContent.replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim();
}

function paraText(el) {
  if (!el) return "";
  return cleanFragment(el.cloneNode(true));
}

function selectionText(sel) {
  const range = sel.getRangeAt(0);
  const holder = document.createElement("div");
  holder.append(range.cloneContents());
  // 跨段选区：块级元素之间补换行
  holder.querySelectorAll("p,h1,h2,h3,h4,h5,h6,td").forEach(b => b.append("\n"));
  // 所见即所得：划选范围内的删除文字（删除线）也纳入，可翻译 / 解释
  return cleanFragment(holder, true);
}

// ---------- 划词 ----------

function wireSelection() {
  const bar = $("selToolbar");
  const zones = () => [$("docPage"), $("commentList"), $("resultList")];
  document.addEventListener("mouseup", e => {
    if (bar.contains(e.target) || $("selPop").contains(e.target)) return;
    setTimeout(() => {
      const sel = window.getSelection();
      if (!sel || sel.isCollapsed || !sel.rangeCount) { bar.classList.add("hidden"); return; }
      const node = sel.anchorNode;
      const host = node?.nodeType === 1 ? node : node?.parentElement;
      if (!host || !zones().some(z => z.contains(host))) { bar.classList.add("hidden"); return; }
      const text = $("docPage").contains(host) ? selectionText(sel) : sel.toString().trim();
      if (text.length < 2) { bar.classList.add("hidden"); return; }
      const rect = sel.getRangeAt(0).getBoundingClientRect();
      const context = $("docPage").contains(host) ? paraText(host.closest("[data-p]")) : "";
      const inDoc = $("docPage").contains(host);
      const range = sel.getRangeAt(0).cloneRange();
      const revs = inDoc ? revsInRange(range) : [];
      bar._payload = { text, context, rect, range, inDoc, revs };
      bar.querySelectorAll(".edit-act, .sep").forEach(el => el.classList.toggle("hidden", !inDoc));
      for (const b of bar.querySelectorAll(".rev-act")) {
        b.classList.toggle("hidden", !revs.length);
        b.textContent = `${b.dataset.act === "acceptSel" ? "✓ 接受" : "✕ 拒绝"}修订(${revs.length})`;
      }
      if ($("instantChk").checked) { bar.classList.add("hidden"); runQuick("translate", bar._payload); return; }
      bar.classList.remove("hidden");
      const bw = bar.offsetWidth || 280;
      bar.style.left = Math.max(8, Math.min(innerWidth - bw - 8, rect.left + rect.width / 2 - bw / 2)) + "px";
      bar.style.top = (rect.top > 50 ? rect.top - 42 : rect.bottom + 8) + "px";
    }, 10);
  });
  document.addEventListener("mousedown", e => { if (!bar.contains(e.target)) bar.classList.add("hidden"); }, true);
  $("docScroll").addEventListener("scroll", () => bar.classList.add("hidden"));
  bar.querySelectorAll("button").forEach(b => b.addEventListener("click", () => {
    bar.classList.add("hidden");
    const p = bar._payload;
    if (!p) return;
    if (b.dataset.act === "copy") copy(p.text);
    else if (b.dataset.act === "comment") newComment(p);
    else if (b.dataset.act === "rewrite") rewriteSel(p);
    else if (b.dataset.act === "delSel") delSel(p);
    else if (b.dataset.act === "acceptSel" || b.dataset.act === "rejectSel") applySelRevs(p.revs, b.dataset.act === "acceptSel");
    else runQuick(b.dataset.act, p);
  }));
}

const ACT_LABEL = { translate: "翻译", explain: "解释", term: "术语", summarize: "概括" };

function placePop(rect) {
  const pop = $("selPop");
  pop.classList.remove("hidden");
  const w = pop.offsetWidth, h = Math.min(pop.offsetHeight || 200, innerHeight * .46);
  let left = Math.max(12, Math.min(innerWidth - w - 12, rect.left + rect.width / 2 - w / 2));
  let top = rect.bottom + 10;
  if (top + h > innerHeight - 12) top = Math.max(12, rect.top - h - 10);
  pop.style.left = left + "px";
  pop.style.top = top + "px";
}

async function runQuick(action, payload) {
  const { text, context, rect, revs } = payload;
  S.popAbort?.abort();
  const ctrl = new AbortController();
  S.popAbort = ctrl;
  $("spLabel").textContent = ACT_LABEL[action] || action;
  S.popRevs = revs || [];
  S.popPayload = payload;
  for (const id of ["spComment", "spRewrite", "spDel"]) $(id).classList.toggle("hidden", !payload.inDoc);
  for (const [id, word] of [["spAccept", "✓ 接受"], ["spReject", "✕ 拒绝"]]) {
    $(id).classList.toggle("hidden", !S.popRevs.length);
    $(id).textContent = `${word}修订(${S.popRevs.length})`;
  }
  $("spSrc").textContent = text;
  const body = $("spBody");
  body.classList.add("typing");
  const stop = waitHint(body, `${ACT_LABEL[action] || action}中`);
  placePop(rect);
  let acc = "", err = null;
  try {
    await ssePost("/api/review/quick", { text, action, context }, ev => {
      if (ev.delta) { stop(); acc += ev.delta; body.innerHTML = renderMD(acc); }
      if (ev.error) err = ev.error;
    }, ctrl.signal);
    stop();
    if (err) throw new Error(err);
    body.classList.remove("typing");
    $("spBody")._text = acc;
    addResult(action, text, acc);
  } catch (e) {
    stop();
    body.classList.remove("typing");
    if (e.name !== "AbortError") body.innerHTML = `<p>⚠️ ${escapeHtml(e.message)}</p>`;
  }
}

function closePop() {
  S.popAbort?.abort();
  $("selPop").classList.add("hidden");
}

// ---------- 段落翻译 ----------

function wireParagraphs() {
  const btn = $("paraBtn");
  const page = $("docPage");
  let cur = null;
  page.addEventListener("mouseover", e => {
    const p = e.target.closest("[data-p]");
    if (!p || p === cur || p.closest("td") || !paraText(p)) return;
    cur = p;
    const pr = p.getBoundingClientRect(), panel = $("docPanel").getBoundingClientRect();
    const pageR = page.getBoundingClientRect();
    btn.style.left = (pageR.right - panel.left - 34) + "px";
    btn.style.top = (pr.top - panel.top + 2) + "px";
    btn.classList.remove("hidden");
  });
  $("docScroll").addEventListener("scroll", () => { btn.classList.add("hidden"); cur = null; });
  btn.addEventListener("click", async () => {
    if (!cur) return;
    const p = cur;
    const next = p.nextElementSibling;
    if (next?.classList.contains("p-tr")) { next.remove(); return; }  // 再点一次收起
    const box = document.createElement("div");
    box.className = "p-tr md typing";
    p.after(box);
    let acc = "", err = null;
    const stop = waitHint(box, "翻译中");
    try {
      await ssePost("/api/review/quick", { text: paraText(p), action: "translate" }, ev => {
        if (ev.delta) { stop(); acc += ev.delta; box.innerHTML = renderMD(acc); }
        if (ev.error) err = ev.error;
      });
      stop();
      if (err) throw new Error(err);
    } catch (e) { stop(); box.innerHTML = `⚠️ ${escapeHtml(e.message)}`; }
    box.classList.remove("typing");
  });
}

// ---------- 结果 ----------

function addResult(action, src, text) {
  S.results.unshift({ action, src, text, ts: Date.now() });
  S.results = S.results.slice(0, 100);
  renderResults();
}

function renderResults() {
  $("resCount").textContent = S.results.length ? `(${S.results.length})` : "";
  const list = $("resultList");
  if (!S.results.length) { list.innerHTML = '<div class="hint">划词翻译、解释的结果会留存在这里。</div>'; return; }
  list.innerHTML = S.results.map((r, i) => `
    <div class="res"><div class="lbl">${ACT_LABEL[r.action] || r.action} · ${new Date(r.ts).toLocaleTimeString("zh-CN", { hour12: false })}
      <button class="btn small ghost" data-i="${i}" style="float:right">复制</button></div>
      <div class="src">${escapeHtml(r.src)}</div><div class="md">${renderMD(r.text)}</div></div>`).join("");
  list.querySelectorAll("button[data-i]").forEach(b => b.addEventListener("click", () => copy(S.results[+b.dataset.i].text)));
}

// ---------- 原生写回（批注 / 修订 / 改写，直接改原文件） ----------

function setAuthor(a) {
  lsSet("rv-author", a);
  $("authorBtn").textContent = "署名：" + (a || "未设置");
}

function ensureAuthor() {
  let a = lsGet("rv-author", "");
  if (!a) {
    a = (prompt("批注与修订的署名（会显示在 WPS / Word 里）：", "") || "").trim();
    if (a) setAuthor(a);
  }
  return a;
}

async function postRaw(url, body) {
  const r = await fetch(url, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) });
  let data = {};
  try { data = await r.json(); } catch { /* 非 JSON 响应 */ }
  return { ok: r.ok, status: r.status, data };
}

const LOCK_HINT = "已写入。WPS 正开着此文件：关闭时请选「不保存」，否则会覆盖面板的修改";

async function edit(op, payload, okMsg) {
  if (!S.doc) return false;
  const author = ensureAuthor();
  if (!author) return false;
  // 不弹确认：WPS 占用时照常写入，事后提示
  const body = { id: S.doc.id, mtime: S.doc.mtime, author, op, ...payload, force: true };
  if (op === "rev" && payload.action === "accept" && $("autoCleanChk").checked) body.auto_clean = true;
  S.busy = true;
  try {
    const r = await postRaw("/api/review/edit", body);
    if (r.ok) {
      await loadDoc(S.doc.id, true);
      const busy = r.data.holders?.length;
      const cleaned = r.data.cleaned ? `；该段已通过，删除对应批注 ${r.data.cleaned} 条（可撤销）` : "";
      toast(busy ? LOCK_HINT : (okMsg || "已写入原文件") + cleaned, busy ? "err" : "ok");
      return r.data;
    }
    if (r.data.code === "stale") await loadDoc(S.doc.id, true);
    toast(r.data.error || `写入失败（${r.status}）`, "err");
    return false;
  } catch (e) {
    toast("写入失败：" + e.message, "err");
    return false;
  } finally {
    S.busy = false;
  }
}

async function undoLast(redo = false) {
  if (!S.doc) return;
  S.busy = true;
  try {
    const r = await postRaw(redo ? "/api/review/redo" : "/api/review/undo", { id: S.doc.id });
    if (!r.ok) { toast(r.data.error || `${redo ? "重做" : "撤销"}失败（${r.status}）`, "err"); return; }
    await loadDoc(S.doc.id, true);
    toast(redo ? `已重做，还可重做 ${r.data.left} 步` : `已撤销（点「↷ 重做」可恢复），还可撤销 ${r.data.left} 步`, "ok");
  } finally {
    S.busy = false;
  }
}

// 选区端点 → 段落内字符偏移（与后端文字模型一致：含删除文字，不含 💬 / 脚注号）
const closestP = n => (n?.nodeType === 1 ? n : n?.parentElement)?.closest("#docPage [data-p]");

function stripMarks(holder) {
  holder.querySelectorAll(".cm-ref, .fn-ref, .p-tr").forEach(n => n.remove());
  return holder.textContent;
}

function rawParaText(pEl) {
  return stripMarks(pEl.cloneNode(true));
}

function textUpTo(pEl, node, off) {
  const r = document.createRange();
  r.selectNodeContents(pEl);
  r.setEnd(node, off);
  const holder = document.createElement("div");
  holder.append(r.cloneContents());
  return stripMarks(holder);
}

function rangeAnchor(range) {
  const sp = closestP(range.startContainer);
  let ep = closestP(range.endContainer);
  if (!sp || !ep) return null;
  const so = textUpTo(sp, range.startContainer, range.startOffset).length;
  let eo = textUpTo(ep, range.endContainer, range.endOffset).length;
  // 三击选整段时终点落在下一段开头：收回到本段末尾
  if (ep !== sp && eo === 0) {
    const all = [...document.querySelectorAll("#docPage [data-p]")];
    const prev = all[all.indexOf(ep) - 1];
    if (prev) { ep = prev; eo = rawParaText(prev).length; }
  }
  const all = [...document.querySelectorAll('#docPage [data-p]')];
  const i = all.indexOf(sp), j = all.indexOf(ep);
  if (i < 0 || j < i) return null;
  return {
    start: { p: +sp.dataset.p, o: so }, end: { p: +ep.dataset.p, o: eo },
    ptexts: Object.fromEntries(all.slice(i, j + 1).map(p => [p.dataset.p, rawParaText(p)])),
  };
}

function openEditPop({ title, src = "", initial = "", rect, onOk }) {
  closePop();
  hideRevBar();
  const pop = $("editPop");
  $("epTitle").textContent = title;
  $("epSrc").textContent = src;
  $("epSrc").classList.toggle("hidden", !src);
  $("epText").value = initial;
  S.epOk = onOk;
  pop.classList.remove("hidden");
  const w = pop.offsetWidth, h = pop.offsetHeight;
  const r = rect || { left: innerWidth / 2, right: innerWidth / 2, top: innerHeight / 3, bottom: innerHeight / 3, width: 0 };
  let top = r.bottom + 10;
  if (top + h > innerHeight - 12) top = Math.max(12, r.top - h - 10);
  pop.style.left = Math.max(12, Math.min(innerWidth - w - 12, r.left + (r.width || 0) / 2 - w / 2)) + "px";
  pop.style.top = top + "px";
  const ta = $("epText");
  ta.focus();
  ta.setSelectionRange(ta.value.length, ta.value.length);
}

function closeEditPop() {
  $("editPop").classList.add("hidden");
  S.epOk = null;
}

async function submitEdit() {
  const fn = S.epOk;
  if (!fn) return;
  const ok = $("epOk");
  ok.disabled = true;
  try {
    if ((await fn($("epText").value)) !== false) closeEditPop();
  } finally {
    ok.disabled = false;
  }
}

function newComment(p) {
  const a = rangeAnchor(p.range);
  if (!a) { toast("请在正文段落内选择", "err"); return; }
  openEditPop({
    title: "新建批注", src: p.text, rect: p.rect,
    onOk: text => text.trim() ? edit("add_comment", { ...a, text }, "批注已写入") : (toast("批注内容不能为空", "err"), false),
  });
}

function rewriteSel(p) {
  const a = rangeAnchor(p.range);
  if (!a) { toast("请在正文段落内选择", "err"); return; }
  if (a.start.p === a.end.p && a.end.o <= a.start.o) { toast("请选中要改写的文字", "err"); return; }
  const multi = a.start.p !== a.end.p;
  openEditPop({
    title: multi ? "跨段改写（请保留原段落数，每段用换行分隔）" : "改写（以修订痕迹写入：原文标删除，新文标插入）", src: p.text, initial: p.text, rect: p.rect,
    onOk: text => text === p.text ? (toast("内容没有变化", "err"), false)
      : edit("rewrite", multi ? { ...a, text } : { p: a.start.p, s: a.start.o, e: a.end.o, ptext: a.ptexts[a.start.p], text }, "已改写（修订）"),
  });
}

// 删减：把选中文字以「删除修订」写入（原文标删除线，不插入新文），可撤销
function delSel(p) {
  const a = rangeAnchor(p.range);
  if (!a) { toast("请在正文段落内选择", "err"); return; }
  if (a.start.p !== a.end.p) { toast("删减一次只能在一个段落内进行", "err"); return; }
  if (a.end.o <= a.start.o) { toast("请选中要删减的文字", "err"); return; }
  edit("rewrite", { p: a.start.p, s: a.start.o, e: a.end.o, ptext: a.ptexts[a.start.p], text: "" },
    "已删减（原文标删除线写入，可撤销）");
}

// 与选区有交集的修订（整条处理）
function revsInRange(range) {
  return [...document.querySelectorAll("#docPage [data-rev]")]
    .filter(el => range.intersectsNode(el)).map(el => +el.dataset.rev);
}

function applySelRevs(ids, accept) {
  if (!ids?.length) { toast("选区内没有修订", "err"); return; }
  closePop();
  window.getSelection().removeAllRanges();
  edit("rev", { indices: ids, action: accept ? "accept" : "reject" }, `已${accept ? "接受" : "拒绝"}选中的 ${ids.length} 处修订（可撤销）`);
}

function revsInComment(cid) {
  return [...document.querySelectorAll("#docPage [data-rev]")]
    .filter(el => el.querySelector(`.cm[data-c~="${cid}"]`) || el.closest(`.cm[data-c~="${cid}"]`))
    .map(el => +el.dataset.rev);
}

function commentAction(c, act, btn) {
  const rect = btn.getBoundingClientRect();
  if (act === "reply") {
    openEditPop({ title: "回复批注", src: c.text, rect,
      onOk: text => text.trim() ? edit("reply", { cid: c.id, text }, "回复已写入") : (toast("回复内容不能为空", "err"), false) });
  } else if (act === "resolve") {
    edit("resolve", { cid: c.id, done: !c.done }, c.done ? "已重新打开" : "已标记为已解决");
  } else if (act === "delete") {
    edit("delete_comment", { cid: c.id }, "批注已删除（可撤销）");
  } else {
    const ids = revsInComment(c.id);
    const accept = act === "acceptRange";
    if (!ids.length) { toast("该批注范围内没有修订", "err"); return; }
    edit("rev", { indices: ids, action: accept ? "accept" : "reject" }, `已${accept ? "接受" : "拒绝"} ${ids.length} 处修订`);
  }
}

function showRevBar(el) {
  const bar = $("revBar");
  document.querySelectorAll("#docPage .rev-on").forEach(n => n.classList.remove("rev-on"));
  el.classList.add("rev-on");
  bar._rev = +el.dataset.rev;
  $("revInfo").textContent = el.title || "";
  bar.classList.remove("hidden");
  const r = el.getBoundingClientRect(), w = bar.offsetWidth;
  bar.style.left = Math.max(8, Math.min(innerWidth - w - 8, r.left)) + "px";
  bar.style.top = (r.top > 46 ? r.top - 40 : r.bottom + 6) + "px";
}

function hideRevBar() {
  $("revBar").classList.add("hidden");
  document.querySelectorAll("#docPage .rev-on").forEach(n => n.classList.remove("rev-on"));
}

function wireEditing() {
  $("revBar").querySelectorAll("button").forEach(b => b.addEventListener("click", () => {
    const accept = b.dataset.a === "accept";
    const idx = $("revBar")._rev;
    hideRevBar();
    edit("rev", { indices: [idx], action: b.dataset.a }, accept ? "已接受修订" : "已拒绝修订");
  }));
  $("docScroll").addEventListener("scroll", hideRevBar);
  document.addEventListener("mousedown", e => {
    if (!$("revBar").contains(e.target) && !e.target.closest?.("#docPage [data-rev]")) hideRevBar();
  });
  $("epOk").addEventListener("click", submitEdit);
  $("spAccept").addEventListener("click", () => applySelRevs(S.popRevs, true));
  $("spComment").addEventListener("click", () => { const p = S.popPayload; closePop(); if (p) newComment(p); });
  $("spRewrite").addEventListener("click", () => { const p = S.popPayload; closePop(); if (p) rewriteSel(p); });
  $("spDel").addEventListener("click", () => { const p = S.popPayload; closePop(); if (p) delSel(p); });
  $("spReject").addEventListener("click", () => applySelRevs(S.popRevs, false));
  $("epClose").addEventListener("click", closeEditPop);
  $("undoBtn").addEventListener("click", () => undoLast(false));
  $("redoBtn").addEventListener("click", () => undoLast(true));
  $("authorBtn").addEventListener("click", () => {
    const a = prompt("批注与修订的署名：", lsGet("rv-author", ""));
    if (a !== null && a.trim()) setAuthor(a.trim());
  });
  for (const [id, accept] of [["acceptAllBtn", true], ["rejectAllBtn", false]]) {
    $(id).addEventListener("click", () => {
      if (!S.doc) return;
      const st = S.doc.stats;
      edit("rev_all", { action: accept ? "accept" : "reject" },
        `已${accept ? "接受" : "拒绝"}全部修订（插入 ${st.ins}、删除 ${st.del}，可撤销）`);
    });
  }
}

// ---------- 杂项 ----------

async function copy(text) {
  try { await navigator.clipboard.writeText(text); toast("已复制", "ok"); }
  catch (e) { toast("复制失败：" + e.message, "err"); }
}

function switchTab(name) {
  document.querySelectorAll(".tabs button").forEach(b => b.classList.toggle("on", b.dataset.tab === name));
  $("tab-comments").classList.toggle("hidden", name !== "comments");
  $("tab-results").classList.toggle("hidden", name !== "results");
}

function setView(v) {
  S.view = v;
  lsSet("rv-view", v);
  const page = $("docPage");
  page.classList.remove("view-markup", "view-final", "view-original");
  page.classList.add("view-" + v);
  document.querySelectorAll("#viewSeg button").forEach(b => b.classList.toggle("on", b.dataset.v === v));
}

function setZoom(fs) {
  S.fs = Math.max(11, Math.min(24, fs));
  lsSet("rv-fs", String(S.fs));
  $("docPage").style.setProperty("--fs", S.fs + "px");
}

function wireUI() {
  $("openBtn").addEventListener("click", () => openPath($("folderInput").value));
  $("pickBtn").addEventListener("click", e => pickPath(e.altKey ? "file" : "folder"));
  $("folderInput").addEventListener("keydown", e => { if (e.key === "Enter") openPath($("folderInput").value); });
  $("historySel").addEventListener("change", e => { if (e.target.value) openPath(e.target.value); e.target.value = ""; });
  document.querySelectorAll("#viewSeg button").forEach(b => b.addEventListener("click", () => setView(b.dataset.v)));
  $("zoomIn").addEventListener("click", () => setZoom(S.fs + 1));
  $("zoomOut").addEventListener("click", () => setZoom(S.fs - 1));
  $("instantChk").checked = lsGet("rv-instant", "0") === "1";
  $("instantChk").addEventListener("change", e => lsSet("rv-instant", e.target.checked ? "1" : "0"));
  $("hideDone").addEventListener("change", renderComments);
  $("autoCleanChk").checked = lsGet("rv-autoclean", "1") === "1";
  $("autoCleanChk").addEventListener("change", e => lsSet("rv-autoclean", e.target.checked ? "1" : "0"));
  $("trAllBtn").addEventListener("click", translateAll);
  $("queueAllBtn").addEventListener("click", async () => {
    if (!S.doc) return;
    const items = [];
    for (const c of S.doc.comments.filter(x => !x.parent)) {
      const a = commentAnchorRange(c);
      if (a.error) continue;
      items.push({ cid: c.id, op: "rewrite", text: a.target, context: rewriteContext(a) });
    }
    if (!items.length) { toast("没有可改写的批注", "err"); return; }
    await enqueueJobs(items);
    toast(`已入队 ${items.length} 条改写`, "ok");
  });
  $("exportRespBtn").addEventListener("click", exportResponses);
  $("autoExplainChk").checked = lsGet("rv-auto-explain", "1") === "1";
  $("autoExplainChk").addEventListener("change", e => lsSet("rv-auto-explain", e.target.checked ? "1" : "0"));
  $("autoRewriteChk").checked = lsGet("rv-auto-rewrite", "0") === "1";
  $("autoRewriteChk").addEventListener("change", e => lsSet("rv-auto-rewrite", e.target.checked ? "1" : "0"));
  $("wordBtn").addEventListener("click", async () => {
    if (!S.doc) return;
    try { await postJSON("/api/review/open-word", { id: S.doc.id }); toast("已在 WPS 中打开；保存后本页自动刷新", "ok"); }
    catch (e) { toast(e.message, "err"); }
  });
  document.querySelectorAll(".tabs button").forEach(b => b.addEventListener("click", () => switchTab(b.dataset.tab)));
  $("resClear").addEventListener("click", () => { S.results = []; renderResults(); });
  $("spClose").addEventListener("click", closePop);
  $("spCopy").addEventListener("click", () => copy($("spBody")._text || $("spBody").textContent));
  document.addEventListener("keydown", e => {
    if (e.key === "Escape") { closePop(); closeEditPop(); hideRevBar(); }
    if (e.key === "Enter" && (e.metaKey || e.ctrlKey) && !$("editPop").classList.contains("hidden")) submitEdit();
  });
  document.addEventListener("mousedown", e => {
    const pop = $("selPop");
    if (!pop.classList.contains("hidden") && !pop.contains(e.target) && !$("selToolbar").contains(e.target)) closePop();
  });
  // 正文批注高亮 / 💬 → 右栏对应批注
  $("docPage").addEventListener("click", e => {
    if (!window.getSelection().isCollapsed) return;
    const ref = e.target.closest(".cm-ref");
    const span = e.target.closest(".cm");
    const cid = ref ? ref.dataset.ref : span ? span.dataset.c.split(" ")[0] : null;
    if (cid) focusComment(cid);
    const rv = e.target.closest("[data-rev]");
    if (rv) showRevBar(rv); else hideRevBar();
  });
  wireEditing();
  const theme = lsGet("pp-theme", "dark");
  document.documentElement.dataset.theme = theme;
  $("themeBtn").textContent = theme === "dark" ? "🌙" : "☀️";
  $("themeBtn").addEventListener("click", () => {
    const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
    document.documentElement.dataset.theme = next;
    lsSet("pp-theme", next);
    $("themeBtn").textContent = next === "dark" ? "🌙" : "☀️";
  });
}

// ---------- 布局：侧栏拖拽调宽 / 折叠（记忆在 localStorage，只动本页布局） ----------

const LAYOUT_KEY = "rv-layout";
const LAYOUT = { lw: 300, rw: 380, lc: false, rc: false };
const LAYOUT_MIN = { lw: 180, rw: 240 };

function applyLayout() {
  const lay = $("rvLayout");
  lay.classList.toggle("no-left", LAYOUT.lc);
  lay.classList.toggle("no-right", LAYOUT.rc);
  if (!LAYOUT.lc) lay.style.setProperty("--lw", `${LAYOUT.lw}px`);
  if (!LAYOUT.rc) lay.style.setProperty("--rw", `${LAYOUT.rw}px`);
  $("splitL").querySelector(".rv-tab").textContent = LAYOUT.lc ? "›" : "‹";
  $("splitR").querySelector(".rv-tab").textContent = LAYOUT.rc ? "‹" : "›";
}

function saveLayout() { lsSet(LAYOUT_KEY, JSON.stringify(LAYOUT)); }

function clampWidth(side) {
  const key = side === "left" ? "lw" : "rw";
  // 窄屏（≤1180px）左栏被媒体查询隐藏，宽度预算里按 0 计
  const leftHidden = matchMedia("(max-width: 1180px)").matches;
  const other = side === "left" ? (LAYOUT.rc ? 0 : LAYOUT.rw)
                                : ((LAYOUT.lc || leftHidden) ? 0 : LAYOUT.lw);
  // 正文至少留 380px；窗口不够宽时以最小宽度兜底
  const max = Math.max(LAYOUT_MIN[key], $("rvLayout").clientWidth - 380 - other);
  LAYOUT[key] = Math.min(Math.max(LAYOUT[key], LAYOUT_MIN[key]), max);
}

function togglePanel(side) {
  if (side === "left") LAYOUT.lc = !LAYOUT.lc; else LAYOUT.rc = !LAYOUT.rc;
  if (!(side === "left" ? LAYOUT.lc : LAYOUT.rc)) clampWidth(side);
  applyLayout();
  saveLayout();
}

function wireLayout() {
  try {
    const saved = JSON.parse(lsGet(LAYOUT_KEY, "") || "{}");
    for (const k of ["lw", "rw"]) if (Number.isFinite(saved[k])) LAYOUT[k] = +saved[k];
    for (const k of ["lc", "rc"]) if (typeof saved[k] === "boolean") LAYOUT[k] = saved[k];
  } catch { /* 忽略坏数据，用默认布局 */ }
  for (const side of ["left", "right"]) {
    if (!(side === "left" ? LAYOUT.lc : LAYOUT.rc)) clampWidth(side);
    const strip = $(side === "left" ? "splitL" : "splitR");
    strip.addEventListener("mousedown", ev => {
      if (ev.button !== 0 || ev.target.closest(".rv-tab")) return;
      const panel = side === "left" ? $("filePanel") : $("sidePanel");
      const cs = getComputedStyle(panel);
      if ((side === "left" ? LAYOUT.lc : LAYOUT.rc) || cs.display === "none" || cs.visibility === "hidden") return;
      ev.preventDefault();
      strip.classList.add("drag");
      document.body.classList.add("layout-drag");
      const onMove = e2 => {
        const rect = $("rvLayout").getBoundingClientRect();
        LAYOUT[side === "left" ? "lw" : "rw"] = Math.round(side === "left"
          ? e2.clientX - rect.left : rect.right - e2.clientX);
        clampWidth(side);
        applyLayout();
      };
      const onUp = () => {
        strip.classList.remove("drag");
        document.body.classList.remove("layout-drag");
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
        saveLayout();
      };
      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseup", onUp);
    });
    strip.addEventListener("dblclick", ev => { ev.preventDefault(); togglePanel(side); });
    strip.querySelector(".rv-tab").addEventListener("click", () => togglePanel(side));
  }
  window.addEventListener("resize", () => {
    if (!LAYOUT.lc) clampWidth("left");
    if (!LAYOUT.rc) clampWidth("right");
    applyLayout();
  });
  applyLayout();
}

async function init() {
  wireUI();
  wireLayout();
  wireSelection();
  wireKeyboard();
  wireParagraphs();
  wireTree();
  setView(lsGet("rv-view", "markup"));
  setZoom(+lsGet("rv-fs", "15"));
  setAuthor(lsGet("rv-author", ""));
  const qs = new URLSearchParams(location.search);
  const target = qs.get("path") || qs.get("doc");
  const recentInfo = await loadRecent();
  if (target) return openPath(target);
  const last = lsGet("rv-doc", "");
  // 树的根：上次的根；没有则取最近文档所在文件夹的上一级（与访达里看到的层级一致）
  const lastFolder = last ? last.replace(/\/[^/]+$/, "") : (lsGet("rv-folder", "") || recentInfo.folders[0] || "");
  const root = lsGet("rv-root", "") || (lastFolder ? lastFolder.replace(/\/[^/]+$/, "") : "");
  if (root) await setRoot(root, true);
  if (last) { expandTo(last.replace(/\/[^/]+$/, "")); openDoc(last); }
}

init();
