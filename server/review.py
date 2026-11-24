"""审阅模式：把带批注/修订的 .docx 渲染成 HTML（只读，原文件永不改动）。

渲染要点：修订（w:ins / w:del）保留为 <ins>/<del>；批注区间可跨段落，
用活动批注集合给每段文字打 data-c；表格、图片、脚注、标题样式一并保留。
"""
from __future__ import annotations

import hashlib
import html
import os
import re
import threading
import time
import zipfile
from pathlib import Path

from lxml import etree

from . import store

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _attr(el, name: str, ns: str = W) -> str | None:
    return el.get(f"{{{ns}}}{name}") if el is not None else None


_lock = threading.Lock()
_render_cache: dict[str, tuple[float, dict]] = {}


# ---------- 文档登记（按绝对路径生成稳定 id） ----------

def _index_file() -> Path:
    return store.DATA / "review_index.json"


def _load_index() -> dict:
    return store._read_json(_index_file(), {"docs": {}, "folders": []})


def _save_index(idx: dict) -> None:
    store._write_json(_index_file(), idx)


def doc_id_for(path: Path) -> str:
    return hashlib.sha1(str(path).encode("utf-8")).hexdigest()[:12]


def register(path_str: str) -> dict:
    path = Path(path_str).expanduser().resolve()
    if not path.is_file() or path.suffix.lower() != ".docx":
        raise ValueError("请提供存在的 .docx 文件路径")
    did = doc_id_for(path)
    with _lock:
        idx = _load_index()
        # 保留已有字段（如 last_write_mtime），只更新打开信息
        idx["docs"].setdefault(did, {}).update({"id": did, "path": str(path), "name": path.name, "opened": time.time()})
        _save_index(idx)
    return idx["docs"][did]


def remember_folder(folder: str) -> None:
    with _lock:
        idx = _load_index()
        folders = [f for f in idx.get("folders", []) if f != folder]
        idx["folders"] = [folder] + folders[:9]
        _save_index(idx)


def get_path(did: str) -> Path | None:
    entry = _load_index()["docs"].get(did)
    return Path(entry["path"]) if entry else None


def recent() -> dict:
    idx = _load_index()
    docs = sorted(idx["docs"].values(), key=lambda d: -d.get("opened", 0))[:20]
    for d in docs:
        d["exists"] = os.path.exists(d["path"])
    return {"docs": docs, "folders": idx.get("folders", [])}


_SEARCH_ROOTS = ("~/Desktop", "~/Documents", "~/Downloads", "~")
_SKIP_DIRS = {"Library", "Applications", "node_modules", "__pycache__"}


def resolve_folder(raw: str) -> Path:
    """接受完整路径、~ 路径、file:// 链接，或只写文件夹名（在桌面/文稿/下载等处查找）。"""
    text = raw.strip().strip("'\"")
    if text.startswith("file://"):
        from urllib.parse import unquote, urlparse
        text = unquote(urlparse(text).path)
    path = Path(text).expanduser()
    if path.is_dir():
        return path.resolve()
    if "/" in text.rstrip("/"):
        raise ValueError(f"找不到文件夹：{text}")
    name = text.rstrip("/")
    hits: list[Path] = []
    for root in _SEARCH_ROOTS:
        base = Path(root).expanduser()
        for cur, dirs, _ in os.walk(base):
            depth = len(Path(cur).relative_to(base).parts)
            dirs[:] = [d for d in dirs if not d.startswith(".") and d not in _SKIP_DIRS] if depth < 3 else []
            if name in dirs:
                hit = (Path(cur) / name).resolve()
                if hit not in hits:
                    hits.append(hit)
        if hits:
            break
    if len(hits) == 1:
        return hits[0]
    if hits:
        raise ValueError("有多个同名文件夹，请输入完整路径：" + "；".join(str(h) for h in hits[:4]))
    raise ValueError(f"在桌面、文稿、下载中都没找到文件夹「{name}」，请输入完整路径或点「选择…」")


def pick(kind: str) -> str | None:
    """弹出 macOS 原生选择框，返回所选路径；取消返回 None。"""
    import subprocess
    if kind == "file":
        script = 'POSIX path of (choose file with prompt "选择要审阅的 Word 文档" of type {"docx", "org.openxmlformats.wordprocessingml.document"})'
    else:
        script = 'POSIX path of (choose folder with prompt "选择存放批注稿 / 修订稿的文件夹")'
    # 由 System Events 弹出，保证对话框在最前面（后台服务自身没有窗口焦点）
    full = f'tell application "System Events"\nactivate\nreturn {script}\nend tell'
    try:
        r = subprocess.run(["osascript", "-e", full], capture_output=True, text=True, timeout=600)
    except subprocess.TimeoutExpired:
        return None
    if r.returncode != 0:  # 用户点了取消
        return None
    return r.stdout.strip() or None


def _natural(name: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def list_dir(folder: str) -> dict:
    """目录的一层内容（与访达一致：全部子文件夹 + .docx），供左栏树状映射按需展开。"""
    base = resolve_folder(folder)
    entries = []
    with os.scandir(base) as it:
        for e in it:
            if e.name.startswith((".", "~$")) or e.name in _SKIP_DIRS:
                continue
            try:
                st = e.stat()
            except OSError:
                continue
            if e.is_dir():
                entries.append({"type": "dir", "name": e.name, "path": e.path, "mtime": st.st_mtime})
            elif e.name.lower().endswith(".docx"):
                entries.append({"type": "file", "name": e.name, "path": e.path, "mtime": st.st_mtime,
                                "size": st.st_size, "id": doc_id_for(Path(e.path))})
    entries.sort(key=lambda x: _natural(x["name"]))
    parent = str(base.parent) if base.parent != base else None
    return {"path": str(base), "name": base.name, "parent": parent, "entries": entries}


def list_folder(folder: str) -> tuple[Path, list[dict]]:
    base = resolve_folder(folder)
    out = []
    for p in sorted(base.rglob("*.docx")):
        rel = p.relative_to(base)
        # 跳过 Word 锁文件、隐藏目录与过深层级
        if p.name.startswith("~$") or any(s.startswith(".") for s in rel.parts) or len(rel.parts) > 3:
            continue
        st = p.stat()
        out.append({"path": str(p), "name": p.name, "rel": str(rel), "mtime": st.st_mtime,
                    "size": st.st_size, "id": doc_id_for(p)})
    return base, out


# ---------- 渲染 ----------

def _read_xml(z: zipfile.ZipFile, name: str):
    try:
        return etree.fromstring(z.read(name), parser=etree.XMLParser(resolve_entities=False, no_network=True))
    except KeyError:
        return None


def _style_names(z: zipfile.ZipFile) -> dict[str, str]:
    root = _read_xml(z, "word/styles.xml")
    names = {}
    if root is None:
        return names
    for st in root.iter(_w("style")):
        sid = _attr(st, "styleId")
        nm = st.find(_w("name"))
        if sid:
            names[sid] = (_attr(nm, "val") or sid).lower()
    return names


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    root = _read_xml(z, part)
    out = {}
    if root is None:
        return out
    for rel in root.iter(f"{{{PKG_REL}}}Relationship"):
        out[rel.get("Id")] = rel.get("Target", "")
    return out


def _para_text(el) -> str:
    """段落纯文本（修订后口径：含插入、不含删除）。"""
    parts = []
    for node in el.iter(_w("t"), _w("tab"), _w("br")):
        if node.tag == _w("t"):
            parts.append(node.text or "")
        else:
            parts.append("\t" if node.tag == _w("tab") else "\n")
    return "".join(parts)


# 正文流之外的段落/修订（文本框、批注、脚注内）不参与编号
_OUTSIDE = {"txbxContent", "drawing", "pict", "AlternateContent", "comment", "footnote", "endnote"}
REV_TAGS = ("ins", "del", "moveTo", "moveFrom")


def _in_flow(el) -> bool:
    return not any(etree.QName(a).localname in _OUTSIDE for a in el.iterancestors())


def body_paragraphs(body) -> list:
    """正文段落，按文档顺序；渲染的 data-p 与编辑定位共用这一编号。"""
    return [p for p in body.iter(_w("p")) if _in_flow(p)]


def body_revisions(body) -> list:
    """内容型修订（不含段落标记/行属性上的修订标记），渲染的 data-rev 与编辑共用这一编号。"""
    return [e for e in body.iter(*[_w(t) for t in REV_TAGS])
            if etree.QName(e.getparent()).localname not in ("rPr", "trPr", "pPr") and _in_flow(e)]


class _Renderer:
    def __init__(self, did: str, styles: dict, rels: dict, pmap: dict | None = None, rmap: dict | None = None):
        self.did = did
        self.styles = styles
        self.rels = rels
        self.active: list[str] = []       # 当前处于其区间内的批注 id
        self.anchor: dict[str, list[str]] = {}
        self.anchor_del: dict[str, list[str]] = {}   # 锚定文字全被删除时的回退
        self.pidx = 0
        self.pmap = pmap or {}
        self.rmap = rmap or {}
        self.stats = {"ins": 0, "del": 0, "tables": 0, "images": 0}
        self.footnote_refs: list[str] = []

    # --- 文字片段（统一处理批注包裹） ---
    def _wrap(self, inner: str, plain: str, deleted: bool) -> str:
        if not self.active or not inner:
            return inner
        target = self.anchor_del if deleted else self.anchor
        for cid in self.active:
            target.setdefault(cid, []).append(plain)
        return f'<span class="cm" data-c="{html.escape(" ".join(self.active))}">{inner}</span>'

    def run(self, r, rev: str | None) -> str:
        rpr = r.find(_w("rPr"))
        out = []
        for ch in r:
            tag = etree.QName(ch).localname
            if tag in ("t", "delText"):
                txt = ch.text or ""
                out.append(("t", txt))
            elif tag == "tab":
                out.append(("t", "\t"))
            elif tag in ("br", "cr"):
                out.append(("h", "<br>"))
            elif tag == "noBreakHyphen":
                out.append(("t", "‑"))
            elif tag == "sym":
                code = _attr(ch, "char") or ""
                try:
                    out.append(("t", chr(int(code, 16))))
                except ValueError:
                    pass
            elif tag == "drawing":
                out.append(("h", self.drawing(ch)))
            elif tag == "commentReference":
                cid = _attr(ch, "id")
                out.append(("h", f'<sup class="cm-ref" data-ref="{html.escape(cid)}" title="批注">💬</sup>'))
            elif tag == "footnoteReference":
                fid = _attr(ch, "id")
                self.footnote_refs.append(fid)
                out.append(("h", f'<sup class="fn-ref"><a href="#fn-{html.escape(fid)}">[{len(self.footnote_refs)}]</a></sup>'))
        if not out:
            return ""
        plain = "".join(v for k, v in out if k == "t")
        inner = "".join(html.escape(v) if k == "t" else v for k, v in out)
        inner = self._fmt(inner, rpr)
        return self._wrap(inner, plain, rev == "del")

    def _fmt(self, inner: str, rpr) -> str:
        if rpr is None:
            return inner

        def on(tag):
            el = rpr.find(_w(tag))
            return el is not None and (_attr(el, "val") or "true") not in ("0", "false", "none")

        va = _attr(rpr.find(_w("vertAlign")), "val")
        if va == "superscript":
            inner = f"<sup>{inner}</sup>"
        elif va == "subscript":
            inner = f"<sub>{inner}</sub>"
        if on("b"):
            inner = f"<b>{inner}</b>"
        if on("i"):
            inner = f"<i>{inner}</i>"
        if on("u"):
            inner = f"<u>{inner}</u>"
        if on("strike") or on("dstrike"):
            inner = f"<s>{inner}</s>"
        hl = _attr(rpr.find(_w("highlight")), "val")
        if hl and hl != "none":
            inner = f'<mark class="hl hl-{html.escape(hl)}">{inner}</mark>'
        return inner

    def drawing(self, d) -> str:
        blip = next(d.iter(f"{{{A}}}blip"), None)
        rid = _attr(blip, "embed", R)
        target = self.rels.get(rid or "", "")
        if not target:
            return ""
        self.stats["images"] += 1
        ext = next(d.iter(f"{{{WP}}}extent"), None)
        width = ""
        if ext is not None and ext.get("cx"):
            width = f' style="width:{int(ext.get("cx")) / 9525:.0f}px"'
        name = html.escape(target.split("/")[-1])
        return f'<img class="doc-img" src="/api/review/doc/{self.did}/media/{name}"{width} alt="">'

    # --- 段落内容（递归处理 ins/del/hyperlink/sdt 等容器） ---
    def inline(self, el, rev: str | None = None) -> str:
        out = []
        for ch in el:
            tag = etree.QName(ch).localname
            if tag == "r":
                out.append(self.run(ch, rev))
            elif tag in ("ins", "moveTo"):
                inner = self.inline(ch, "ins")
                if inner:
                    self.stats["ins"] += 1
                    out.append(f'<ins class="rev-ins"{self._rev_attr(ch)} title="{self._rev_title(ch, "插入")}">{inner}</ins>')
            elif tag in ("del", "moveFrom"):
                inner = self.inline(ch, "del")
                if inner:
                    self.stats["del"] += 1
                    out.append(f'<del class="rev-del"{self._rev_attr(ch)} title="{self._rev_title(ch, "删除")}">{inner}</del>')
            elif tag in ("hyperlink", "smartTag", "customXml", "fldSimple", "sdtContent", "bdo", "dir"):
                out.append(self.inline(ch, rev))
            elif tag == "sdt":
                c = ch.find(_w("sdtContent"))
                if c is not None:
                    out.append(self.inline(c, rev))
            elif tag == "commentRangeStart":
                cid = _attr(ch, "id")
                if cid and cid not in self.active:
                    self.active.append(cid)
            elif tag == "commentRangeEnd":
                cid = _attr(ch, "id")
                if cid in self.active:
                    self.active.remove(cid)
            elif tag in ("oMath", "oMathPara"):
                txt = "".join(t.text or "" for t in ch.iter(f"{{{M}}}t"))
                out.append(self._wrap(f'<span class="math">{html.escape(txt)}</span>', txt, rev == "del"))
        return "".join(out)

    def _rev_attr(self, el) -> str:
        rid = self.rmap.get(el)
        return f' data-rev="{rid}"' if rid is not None else ""

    @staticmethod
    def _rev_title(el, kind: str) -> str:
        author = _attr(el, "author") or ""
        date = (_attr(el, "date") or "")[:10]
        return html.escape(f"{kind} · {author} {date}".strip())

    def paragraph(self, p) -> str:
        ppr = p.find(_w("pPr"))
        sid = _attr(ppr.find(_w("pStyle")) if ppr is not None else None, "val") or ""
        sname = self.styles.get(sid, sid.lower())
        jc = _attr(ppr.find(_w("jc")) if ppr is not None else None, "val") or ""
        inner = self.inline(p)
        tag = "p"
        m = re.match(r"heading\s*(\d)", sname)
        if m:
            tag = f"h{min(int(m.group(1)) + 1, 6)}"
        elif sname == "title":
            tag = "h1"
        cls = []
        if jc in ("center", "right", "both", "distribute"):
            cls.append({"center": "al-c", "right": "al-r"}.get(jc, "al-j"))
        if ppr is not None and ppr.find(_w("numPr")) is not None:
            cls.append("li")
        # 段落标记本身被删除（整段删除）
        if ppr is not None and ppr.find(f"{_w('rPr')}/{_w('del')}") is not None and not _para_text(p).strip():
            cls.append("para-del")
        self.pidx += 1
        c = f' class="{" ".join(cls)}"' if cls else ""
        pid = self.pmap.get(p)
        dp = f' data-p="{pid}"' if pid is not None else ""
        return f'<{tag}{dp}{c}>{inner or "<br>"}</{tag}>'

    def table(self, tbl) -> str:
        self.stats["tables"] += 1
        rows = tbl.findall(_w("tr"))
        grid = []  # 每行：[(tc, gridSpan, vmerge)]
        for tr in rows:
            cells = []
            for tc in tr.findall(_w("tc")):
                tcpr = tc.find(_w("tcPr"))
                span = int(_attr(tcpr.find(_w("gridSpan")) if tcpr is not None else None, "val") or 1)
                vm_el = tcpr.find(_w("vMerge")) if tcpr is not None else None
                vm = None if vm_el is None else (_attr(vm_el, "val") or "continue")
                cells.append([tc, span, vm, 1])
            grid.append(cells)
        # 计算纵向合并：restart 单元格向下吞并 continue
        for ri, cells in enumerate(grid):
            col = 0
            for cell in cells:
                if cell[2] == "restart":
                    rr = ri + 1
                    while rr < len(grid):
                        below = _cell_at(grid[rr], col)
                        if below is None or below[2] != "continue":
                            break
                        cell[3] += 1
                        rr += 1
                col += cell[1]
        parts = ['<div class="tbl-wrap"><table class="doc-tbl">']
        for cells in grid:
            parts.append("<tr>")
            for tc, span, vm, rowspan in cells:
                if vm == "continue":
                    # 区间仍需走一遍，保持批注开闭状态正确
                    self.block_children(tc)
                    continue
                attrs = (f' colspan="{span}"' if span > 1 else "") + (f' rowspan="{rowspan}"' if rowspan > 1 else "")
                parts.append(f"<td{attrs}>{self.block_children(tc)}</td>")
            parts.append("</tr>")
        parts.append("</table></div>")
        return "".join(parts)

    def block_children(self, el) -> str:
        out = []
        for ch in el:
            tag = etree.QName(ch).localname
            if tag == "p":
                out.append(self.paragraph(ch))
            elif tag == "tbl":
                out.append(self.table(ch))
            elif tag == "sdt":
                c = ch.find(_w("sdtContent"))
                if c is not None:
                    out.append(self.block_children(c))
            elif tag in ("ins", "del", "customXml"):
                out.append(self.block_children(ch))
            elif tag == "commentRangeStart":
                cid = _attr(ch, "id")
                if cid and cid not in self.active:
                    self.active.append(cid)
            elif tag == "commentRangeEnd":
                cid = _attr(ch, "id")
                if cid in self.active:
                    self.active.remove(cid)
        return "".join(out)


def _cell_at(cells, col):
    c = 0
    for cell in cells:
        if c == col:
            return cell
        c += cell[1]
        if c > col:
            return None
    return None


def _comments(z: zipfile.ZipFile) -> list[dict]:
    root = _read_xml(z, "word/comments.xml")
    if root is None:
        return []
    ext = _read_xml(z, "word/commentsExtended.xml")
    ex = {}
    if ext is not None:
        for ce in ext.iter(f"{{{W15}}}commentEx"):
            ex[ce.get(f"{{{W15}}}paraId")] = (ce.get(f"{{{W15}}}paraIdParent"), ce.get(f"{{{W15}}}done") == "1")
    out, para_to_cid = [], {}
    for c in root.iter(_w("comment")):
        cid = _attr(c, "id")
        paras = c.findall(f".//{_w('p')}")
        text = "\n".join(_para_text(p) for p in paras).strip()
        last_pid = _attr(paras[-1], "paraId", W14) if paras else None
        if last_pid:
            para_to_cid[last_pid] = cid
        out.append({"id": cid, "author": _attr(c, "author") or "", "date": _attr(c, "date") or "",
                    "text": text, "_pid": last_pid})
    for item in out:
        parent_pid, done = ex.get(item.pop("_pid"), (None, False))
        item["parent"] = para_to_cid.get(parent_pid) if parent_pid else None
        item["done"] = done
    return out


def render(did: str) -> dict:
    path = get_path(did)
    if path is None or not path.exists():
        raise FileNotFoundError("文档不存在或已移动")
    mtime = path.stat().st_mtime
    cached = _render_cache.get(did)
    if cached and cached[0] == mtime:
        return cached[1]
    with zipfile.ZipFile(path) as z:
        doc = _read_xml(z, "word/document.xml")
        if doc is None:
            raise ValueError("不是有效的 .docx（缺少 word/document.xml）")
        body = doc.find(_w("body"))
        pmap = {p: i for i, p in enumerate(body_paragraphs(body))}
        rmap = {e: i for i, e in enumerate(body_revisions(body))}
        rnd = _Renderer(did, _style_names(z), _rels(z, "word/_rels/document.xml.rels"), pmap, rmap)
        body_html = rnd.block_children(body)
        footnotes_html = ""
        if rnd.footnote_refs:
            fn_root = _read_xml(z, "word/footnotes.xml")
            fns = {_attr(f, "id"): f for f in fn_root.iter(_w("footnote"))} if fn_root is not None else {}
            items = []
            for n, fid in enumerate(rnd.footnote_refs, 1):
                f = fns.get(fid)
                if f is not None:
                    items.append(f'<li id="fn-{html.escape(fid)}" value="{n}">{rnd.block_children(f)}</li>')
            footnotes_html = f'<ol class="footnotes">{"".join(items)}</ol>' if items else ""
        comments = _comments(z)
    for c in comments:
        c["anchor"] = "".join(rnd.anchor.get(c["id"], [])).strip()
        if not c["anchor"] and rnd.anchor_del.get(c["id"]):
            c["anchor"] = "（已删除）" + "".join(rnd.anchor_del[c["id"]]).strip()
    result = {
        "id": did, "name": path.name, "path": str(path), "mtime": mtime,
        "html": body_html + footnotes_html, "comments": comments,
        "stats": {**rnd.stats, "comments": len(comments), "paragraphs": len(pmap)},
    }
    _render_cache[did] = (mtime, result)
    return result


def media(did: str, name: str) -> tuple[bytes, str] | None:
    path = get_path(did)
    if path is None or "/" in name or ".." in name:
        return None
    with zipfile.ZipFile(path) as z:
        try:
            data = z.read(f"word/media/{name}")
        except KeyError:
            return None
    ext = name.rsplit(".", 1)[-1].lower()
    mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif",
            "bmp": "image/bmp", "tif": "image/tiff", "tiff": "image/tiff",
            "emf": "image/emf", "wmf": "image/wmf"}.get(ext, "application/octet-stream")
    return data, mime


# ---------- 翻译缓存 ----------

def cache_key(*parts: str) -> str:
    return hashlib.sha256("\x00".join(parts).encode("utf-8")).hexdigest()[:24]


def _cache_dir() -> Path:
    d = store.CACHE / "review"
    d.mkdir(parents=True, exist_ok=True)
    return d


def cache_get(key: str) -> str | None:
    f = _cache_dir() / f"{key}.md"
    return f.read_text(encoding="utf-8") if f.exists() else None


def cache_put(key: str, text: str) -> None:
    try:
        (_cache_dir() / f"{key}.md").write_text(text, encoding="utf-8")
    except OSError:
        pass


def mostly_cjk(text: str) -> bool:
    cjk = len(re.findall(r"[一-鿿]", text))
    latin = len(re.findall(r"[A-Za-z]", text))
    return cjk > latin / 2
