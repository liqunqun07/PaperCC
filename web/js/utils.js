// Shared Markdown + LaTeX renderer for desktop, mobile and local note previews.
const moduleURL = new URL(import.meta.url);
const assetToken = moduleURL.searchParams.get('token') || new URLSearchParams(location.search).get('token') || localStorage.getItem('panel_token') || '';
function assetURL(relative, base = moduleURL) {
  const u = new URL(relative, base);
  if (assetToken) u.searchParams.set('token', assetToken);
  return u.href;
}
function loadScript(relative, globalName) {
  if (window[globalName]) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    const timer = setTimeout(() => { s.remove(); reject(Error(`${globalName} 加载超时`)); }, 8000);
    s.src = assetURL(relative);
    s.onload = () => { clearTimeout(timer); window[globalName] ? resolve() : reject(Error(`${globalName} 加载失败`)); };
    s.onerror = () => { clearTimeout(timer); s.remove(); reject(Error(`${globalName} 加载失败`)); };
    document.head.append(s);
  });
}
async function loadMathCSS() {
  if (document.getElementById('papersuite-math-css')) return;
  const href = assetURL('../vendor/katex/katex.min.css');
  const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 8000);
  try {
    const r = await fetch(href, {signal: controller.signal});
    if (!r.ok) throw Error('公式样式加载失败');
    // The authenticated mobile proxy requires a token on font requests as well.
    const css = (await r.text()).replace(/url\(([^)]+)\)/g, (_, value) => {
      const path = value.trim().replace(/^['"]|['"]$/g, '');
      return `url("${assetURL(path, href)}")`;
    });
    const style = document.createElement('style'); style.id = 'papersuite-math-css';
    style.textContent = css; document.head.append(style);
  } finally { clearTimeout(timer); }
}
const richStyle = document.createElement('style');
richStyle.textContent = `
.math-inline { display:inline-block; max-width:100%; vertical-align:middle; overflow-x:auto; overflow-y:hidden; padding:2px 1px; }
.math-display { display:block; max-width:100%; overflow-x:auto; overflow-y:hidden; padding:6px 2px; margin:8px 0; }
.math-display .katex-display { margin:0; text-align:left; }
.math-inline .katex, .math-display .katex { font-size:1.05em; }
.math-pending,.math-fallback { white-space:pre-wrap; overflow-wrap:anywhere; font:inherit; }
.md-table-scroll { max-width:100%; overflow-x:auto; }
.md-table-scroll table { border-collapse:collapse; margin:8px 0; }
.md-table-scroll th,.md-table-scroll td { border:1px solid var(--line,#64748b55); padding:6px 10px; }
#notePreview { flex:1 1 auto; min-height:0; overflow:auto; padding:12px 16px; overflow-wrap:anywhere; }
#notePreview[hidden],#noteEditor[hidden],#preview[hidden],#editor[hidden] { display:none; }
#preview { min-height:52vh; max-width:100%; overflow-wrap:anywhere; padding:16px; border:1px solid var(--line); border-radius:7px; background:white; }
`;
document.head.append(richStyle);
// Load once before any consumer paints history or starts a streamed response.
const ready = await Promise.allSettled([
  loadScript('../vendor/marked/marked.umd.js', 'marked'),
  loadScript('../vendor/katex/katex.min.js', 'katex'), loadMathCSS(),
]);
const mathReady = !!window.katex && ready[2].status === 'fulfilled';
let renderSerial = 0;
function escapedAt(text, at) {
  let slashes = 0; while (at > 0 && text[--at] === '\\') slashes++;
  return slashes % 2 === 1;
}
function protectMath(text) {
  const prefix = `PSMATH${++renderSerial}Q`, items = [];
  let out = '', i = 0;
  const codeRanges = [];
  let offset = 0;
  for (const token of (window.marked?.lexer(text) || [])) {
    if (token.type === 'code') codeRanges.push([offset, offset + token.raw.length]);
    offset += token.raw?.length || 0;
  }
  function insert(source, display, raw, pending = false) {
    const id = `${prefix}${items.length}END`;
    let html;
    if (pending) html = `<span class="math-pending">${escapeHtml(raw)}</span>`;
    else if (mathReady) {
      try {
        html = window.katex.renderToString(source, {displayMode:display, throwOnError:true,
          trust:false, strict:'ignore', maxExpand:1000, maxSize:20, output:'htmlAndMathml'});
        html = `<span class="${display ? 'math-display' : 'math-inline'}">${html}</span>`;
      } catch { html = `<span class="math-fallback" title="公式语法未能解析，保留原文">${escapeHtml(raw)}</span>`; }
    } else html = `<span class="math-fallback" title="公式组件暂不可用，保留原文">${escapeHtml(raw)}</span>`;
    items.push([id, html, raw]); out += id;
  }
  while (i < text.length) {
    const codeRange = codeRanges.find(([start,end]) => i >= start && i < end);
    if (codeRange) { out += text.slice(i,codeRange[1]); i = codeRange[1]; continue; }
    if (text[i] === '<') {
      const tag = text.slice(i).match(/^<\/?[A-Za-z][^>]*>/);
      if (tag) { out += tag[0]; i += tag[0].length; continue; }
    }
    // Markdown code blocks and inline code must retain their literal source.
    if (i === 0 || text[i-1] === '\n') {
      const fence = text.slice(i).match(/^ {0,3}(`{3,}|~{3,})[^\n]*\n/);
      if (fence) {
        const close = new RegExp(`^ {0,3}${fence[1][0]}{${fence[1].length},}[^\\S\\n]*(?:\\n|$)`, 'gm');
        close.lastIndex = i + fence[0].length;
        const m = close.exec(text), end = m ? close.lastIndex : text.length;
        out += text.slice(i,end); i = end; continue;
      }
    }
    if (text[i] === '`' && !escapedAt(text,i)) {
      const ticks = text.slice(i).match(/^`+/)[0], end = text.indexOf(ticks, i+ticks.length);
      if (end >= 0) { out += text.slice(i,end+ticks.length); i = end+ticks.length; continue; }
    }
    let open = '', close = '', display = false, env = false;
    if (!escapedAt(text,i)) {
      if (text.startsWith('$$',i)) { open = close = '$$'; display = true; }
      else if (text[i] === '$' && !/\s/.test(text[i+1] || ' ')) open = close = '$';
      else if (text.startsWith('\\[',i)) { open = '\\['; close = '\\]'; display = true; }
      else if (text.startsWith('\\(',i)) { open = '\\('; close = '\\)'; }
      else {
        const m = text.slice(i).match(/^\\begin\{(equation\*?|align\*?|aligned|gather\*?|gathered|multline\*?|cases|[bpBvV]?matrix|array)\}/);
        if (m) { open = m[0]; close = `\\end{${m[1]}}`; display = env = true; }
      }
    }
    if (open) {
      let end = text.indexOf(close,i+open.length);
      while (end >= 0 && (escapedAt(text,end) || (open === '$' && (/\s/.test(text[end-1]) || /\d/.test(text[end+1] || '')))))
        end = text.indexOf(close,end+close.length);
      if (end >= 0) {
        const source = text.slice(i+open.length,end);
        if (source.trim() && (display || !source.includes('\n\n')) &&
            !(open === '$' && /^\d/.test(source) && /[^\x00-\x7f]|\s+[A-Za-z]{2,}/.test(source))) {
          const raw = text.slice(i,end+close.length);
          insert(env ? raw : source,display,raw); i = end+close.length; continue;
        }
      } else if (open !== '$' || !/^\d/.test(text[i+1] || '')) {
        // Keep an unfinished streaming formula intact until its closing delimiter arrives.
        insert('',false,text.slice(i),true); break;
      }
    }
    out += text[i++];
  }
  return {out,items};
}
function cleanMarkdown(html) {
  const template = document.createElement('template'); template.innerHTML = html;
  template.content.querySelectorAll('script,style,iframe,object,embed,link,meta,base,form').forEach(el => el.remove());
  for (const el of template.content.querySelectorAll('*')) {
    for (const attr of [...el.attributes]) {
      if (/^on/i.test(attr.name) || ['srcdoc','style'].includes(attr.name) ||
          (['href','src','xlink:href'].includes(attr.name) && !/^(https?:|mailto:|tel:|\/|#|\.\.?\/|[^:]+$)/i.test(attr.value.replace(/[\s\u0000-\u001f]/g,'')))) el.removeAttribute(attr.name);
    }
  }
  for (const table of template.content.querySelectorAll('table')) {
    const wrapper = document.createElement('div'); wrapper.className = 'md-table-scroll';
    table.replaceWith(wrapper); wrapper.append(table);
  }
  return template.innerHTML;
}
export function renderMD(text) {
  const {out,items} = protectMath(String(text ?? ''));
  let html;
  try { html = window.marked ? cleanMarkdown(window.marked.parse(out,{breaks:true,gfm:true})) : escapeHtml(out).replace(/\n/g,'<br>'); }
  catch { html = escapeHtml(out).replace(/\n/g,'<br>'); }
  const template = document.createElement('template'); template.innerHTML = html;
  const lookup = new Map(items.map(item => [item[0],item]));
  const pattern = new RegExp(items.map(([id]) => id).join('|') || '(?!)','g');
  // Insert generated math only into text nodes, never into HTML attributes or code.
  const walker = document.createTreeWalker(template.content,NodeFilter.SHOW_TEXT), nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  for (const node of nodes) {
    const source = node.textContent; pattern.lastIndex = 0;
    if (!pattern.test(source)) continue;
    const fragment = document.createElement('template');
    if (node.parentElement?.closest('pre,code')) {
      pattern.lastIndex = 0; node.textContent = source.replace(pattern,id => lookup.get(id)[2]); continue;
    }
    pattern.lastIndex = 0;
    let cursor = 0, rendered = '';
    for (const match of source.matchAll(pattern)) {
      let start = match.index, end = start + match[0].length, formula = lookup.get(match[0])[1];
      if (formula.startsWith('<span class="math-inline">')) {
        const before = start > cursor && /[（(【\[]/.test(source[start-1]) ? source[--start] : '';
        const after = source.slice(end).match(/^[，。；：！？、,.;:!?）)】\]]+/)?.[0] || '';
        end += after.length;
        formula = formula.replace('<span class="math-inline">','<span class="math-inline">'+escapeHtml(before));
        formula = formula.slice(0,-7) + escapeHtml(after) + '</span>';
      }
      rendered += escapeHtml(source.slice(cursor,start)) + formula;
      cursor = end;
    }
    fragment.innerHTML = rendered + escapeHtml(source.slice(cursor));
    node.replaceWith(fragment.content);
  }
  for (const el of template.content.querySelectorAll('*')) for (const attr of [...el.attributes]) {
    if (pattern.test(attr.value)) el.setAttribute(attr.name,attr.value.replace(pattern,id => lookup.get(id)[2]));
    pattern.lastIndex = 0;
  }
  return template.innerHTML;
}

export function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g,
    c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

let toastTimer;
export function toast(msg, type = "") {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.className = type;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add("hidden"), 3200);
}

export function fmtSize(n) {
  if (!n) return "";
  if (n < 1024 * 1024) return (n / 1024).toFixed(0) + " KB";
  return (n / 1024 / 1024).toFixed(1) + " MB";
}

export function debounce(fn, ms) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

export function download(url) {
  const a = document.createElement("a");
  a.href = url; a.download = "";
  document.body.appendChild(a); a.click(); a.remove();
}
