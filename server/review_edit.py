"""审阅模式的原生写回：批注（新建/回复/解决/删除）、修订（接受/拒绝）、选中改写（带修订痕迹）。

直接改 .docx 内部 XML，WPS / Word 打开看到的是原生批注与修订。
每次写入前自动备份到 data/review_backups/<doc_id>/，支持逐步撤销。
"""
from __future__ import annotations

import copy
import random
import shutil
import subprocess
import time
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

from . import review, store
from .review import W, W14, W15, _w, _attr, body_paragraphs, body_revisions

M = review.M
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL = review.PKG_REL
PARTS = {
    "comments": ("word/comments.xml", "comments",
                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                 "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", W),
    "ext": ("word/commentsExtended.xml", "commentsEx",
            "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml", W15),
}
KEEP_BACKUPS = 10


class EditError(Exception):
    def __init__(self, msg: str, code: str = "bad", status: int = 400):
        super().__init__(msg)
        self.code, self.status = code, status


# ---------- 文件占用检测 / 备份 / 写入 ----------

def holders(path: Path) -> list[str]:
    """正在打开该文件的 WPS / Word 进程名。"""
    names = set()
    try:
        out = subprocess.run(["lsof", "-Fc", str(path)], capture_output=True, text=True, timeout=5).stdout
        for line in out.splitlines():
            if line.startswith("c"):
                names.add(line[1:])
    except (OSError, subprocess.TimeoutExpired):
        pass
    found = sorted({"WPS" if "wps" in n.lower() else "Microsoft Word"
                    for n in names if "wps" in n.lower() or "word" in n.lower()})
    lock = path.with_name("~$" + path.name[2:])
    if lock.exists() and "Microsoft Word" not in found:
        found.append("Microsoft Word")
    return found


def _backup_dir(did: str) -> Path:
    d = store.DATA / "review_backups" / did
    d.mkdir(parents=True, exist_ok=True)
    return d


def _by_created(files) -> list[Path]:
    """按备份文件的创建时间排序（不按文件名：时区变化会打乱名字里的时间）。"""
    return sorted(files, key=lambda f: f.stat().st_birthtime)


def _redo_dir(did: str) -> Path:
    d = _backup_dir(did) / "redo"
    d.mkdir(exist_ok=True)
    return d


def backup(did: str, path: Path, folder: Path | None = None) -> Path:
    d = folder or _backup_dir(did)
    # 名字以毫秒时间戳开头：与时区无关、按名排序即按时间排序；后接本地时间便于人工辨认
    dst = d / f"{int(time.time() * 1000)}_{time.strftime('%Y%m%d-%H%M%S')}_{path.name}"
    shutil.copy2(path, dst)
    for old in _by_created(d.glob("*.docx"))[:-KEEP_BACKUPS]:
        old.unlink(missing_ok=True)
    return dst


def _set_last_write(did: str, mtime: float) -> None:
    with review._lock:
        idx = review._load_index()
        if did in idx["docs"]:
            idx["docs"][did]["last_write_mtime"] = mtime
            review._save_index(idx)


class Package:
    """docx 读改写：只重写改动过的部件，其余条目原样拷贝。"""

    def __init__(self, path: Path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.infos = z.infolist()
            self.raw = {i.filename: z.read(i.filename) for i in self.infos}
        self.trees: dict[str, etree._Element] = {}
        self.dirty: set[str] = set()

    def xml(self, name: str):
        if name not in self.trees:
            if name not in self.raw:
                return None
            self.trees[name] = etree.fromstring(self.raw[name], parser=etree.XMLParser(resolve_entities=False, no_network=True))
        return self.trees[name]

    def touch(self, name: str) -> None:
        self.dirty.add(name)

    def ensure_part(self, key: str):
        name, root_tag, rel_type, ctype, ns = PARTS[key]
        root = self.xml(name)
        if root is not None:
            return root
        root = etree.Element(f"{{{ns}}}{root_tag}", nsmap={"w": W, "w14": W14, "w15": W15})
        self.trees[name] = root
        self.touch(name)
        rels = self.xml("word/_rels/document.xml.rels")
        ids = {r.get("Id") for r in rels}
        n = 1
        while f"rId{n}" in ids:
            n += 1
        etree.SubElement(rels, f"{{{PKG_REL}}}Relationship", Id=f"rId{n}", Type=rel_type,
                         Target=name.split("/", 1)[1])
        self.touch("word/_rels/document.xml.rels")
        ct = self.xml("[Content_Types].xml")
        etree.SubElement(ct, f"{{{CT_NS}}}Override", PartName="/" + name, ContentType=ctype)
        self.touch("[Content_Types].xml")
        return root

    def save(self, did: str) -> float:
        for name in self.dirty:
            self.raw[name] = etree.tostring(self.trees[name], xml_declaration=True,
                                            encoding="UTF-8", standalone=True)
        names = {i.filename for i in self.infos}
        tmp = store.DATA / "review_backups" / f".tmp_{did}.docx"
        tmp.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(tmp, "w") as z:
            for info in self.infos:
                z.writestr(info, self.raw[info.filename], compress_type=info.compress_type)
            for name in sorted(self.dirty - names):
                z.writestr(name, self.raw[name], compress_type=zipfile.ZIP_DEFLATED)
        with zipfile.ZipFile(tmp) as check:
            bad = check.testzip()
        if bad:
            tmp.unlink(missing_ok=True)
            raise EditError(f"写入校验失败（{bad}），原文件未改动", "io", 500)
        backup(did, self.path)
        for f in _redo_dir(did).glob("*.docx"):   # 新的写入使「重做」失效
            f.unlink(missing_ok=True)
        shutil.copyfile(tmp, self.path)  # 覆盖内容但保留原文件 inode，WPS/访达里位置不变
        tmp.unlink(missing_ok=True)
        mtime = self.path.stat().st_mtime
        _set_last_write(did, mtime)
        return mtime


# ---------- 通用工具 ----------

class Ctx:
    def __init__(self, pkg: Package, author: str):
        self.pkg = pkg
        self.doc = pkg.xml("word/document.xml")
        self.body = self.doc.find(_w("body"))
        self.author = author
        self.date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        ids = [0]
        for name in ("word/document.xml", "word/comments.xml", "word/footnotes.xml", "word/endnotes.xml"):
            root = pkg.xml(name)
            if root is not None:
                ids += [int(v) for v in root.xpath("//@w:id", namespaces={"w": W}) if v.lstrip("-").isdigit()]
        self._next = max(ids) + 1

    def new_id(self) -> str:
        self._next += 1
        return str(self._next - 1)

    def rev(self, tag: str):
        return etree.Element(_w(tag), {_w("id"): self.new_id(), _w("author"): self.author, _w("date"): self.date})


def _para_id() -> str:
    return f"{random.randint(0x10000000, 0x7FFFFFFE):08X}"


def _paragraph(index: int, body) -> etree._Element:
    paras = body_paragraphs(body)
    if not 0 <= index < len(paras):
        raise EditError("段落定位失败，请刷新后重试", "stale", 409)
    return paras[index]


def _run_chunks(r) -> list[tuple[etree._Element, str]]:
    out = []
    for ch in r:
        tag = etree.QName(ch).localname
        if tag in ("t", "delText"):
            out.append((ch, ch.text or ""))
        elif tag == "tab":
            out.append((ch, "\t"))
        elif tag == "noBreakHyphen":
            out.append((ch, "‑"))
        elif tag == "sym":
            try:
                out.append((ch, chr(int(_attr(ch, "char") or "", 16))))
            except ValueError:
                out.append((ch, ""))
        else:
            out.append((ch, ""))
    return out


def _segments(p) -> list[tuple[etree._Element, str, str]]:
    """段落的文字模型，与前端渲染文本逐字一致：[(节点, 文本, run|atom)]。"""
    segs = []
    for el in p.iter(_w("r"), f"{{{M}}}oMath"):
        if next((a for a in el.iterancestors(_w("p"))), None) is not p or not review._in_flow(el):
            continue
        if el.tag == _w("r"):
            if any(a.tag == f"{{{M}}}oMath" for a in el.iterancestors()):
                continue
            segs.append((el, "".join(t for _, t in _run_chunks(el)), "run"))
        else:
            segs.append((el, "".join(t.text or "" for t in el.iter(f"{{{M}}}t")), "atom"))
    return segs


def para_text(p) -> str:
    return "".join(t for _, t, _ in _segments(p))


def _split_run(r, k: int):
    """把 run 在第 k 个字符处一分为二，返回右半部分（新 run，紧跟其后）。"""
    rpr = r.find(_w("rPr"))
    right = etree.Element(_w("r"), dict(r.attrib))
    if rpr is not None:
        right.append(copy.deepcopy(rpr))
    pos, moving = 0, False
    for ch, text in _run_chunks(r):
        if ch.tag == _w("rPr"):
            continue
        if moving:
            right.append(ch)
            continue
        L = len(text)
        if pos + L <= k:
            pos += L
            if pos == k and L:
                moving = True
            continue
        cut = k - pos
        if ch.tag in (_w("t"), _w("delText")) and 0 < cut < L:
            tail = etree.SubElement(right, ch.tag)
            tail.text = text[cut:]
            tail.set(XML_SPACE, "preserve")
            ch.text = text[:cut]
            ch.set(XML_SPACE, "preserve")
        else:
            right.append(ch)
        moving = True
    r.addnext(right)
    return right


_NO_SPLIT = {"sdt", "sdtContent", "fldSimple", "oMath", "oMathPara"}


def _clone_container(parent, ctx: Ctx):
    clone = etree.Element(parent.tag, dict(parent.attrib))
    if etree.QName(parent).localname in review.REV_TAGS and parent.get(_w("id")):
        clone.set(_w("id"), ctx.new_id())
    return clone


def _lift_before(p, node, ctx: Ctx) -> int:
    """把 node 之前的位置提升到段落顶层（必要时拆分外层容器），返回顶层插入下标。"""
    while node.getparent() is not p:
        parent = node.getparent()
        if etree.QName(parent).localname in _NO_SPLIT:
            raise EditError("选区端点位于内容控件或域中，暂不支持，请换个位置")
        idx = parent.index(node)
        first_content = next((i for i, c in enumerate(parent) if not etree.QName(c).localname.endswith("Pr")), 0)
        if idx > first_content:
            clone = _clone_container(parent, ctx)
            for ch in list(parent)[idx:]:
                clone.append(ch)
            parent.addnext(clone)
            node = clone
        else:
            node = parent
    return p.index(node)


def _lift_after(p, node, ctx: Ctx) -> int:
    while node.getparent() is not p:
        parent = node.getparent()
        if etree.QName(parent).localname in _NO_SPLIT:
            raise EditError("选区端点位于内容控件或域中，暂不支持，请换个位置")
        idx = parent.index(node)
        if idx < len(parent) - 1:
            clone = _clone_container(parent, ctx)
            for ch in list(parent)[idx + 1:]:
                clone.append(ch)
            parent.addnext(clone)
        node = parent
    return p.index(node) + 1


def split_point(p, offset: int, ctx: Ctx) -> int:
    segs = [s for s in _segments(p) if s[1]]
    pos = 0
    for node, text, kind in segs:
        L = len(text)
        if offset <= pos:
            return _lift_before(p, node, ctx)
        if offset < pos + L:
            if kind == "atom":
                return _lift_before(p, node, ctx)
            return _lift_before(p, _split_run(node, offset - pos), ctx)
        pos += L
    if segs:
        return _lift_after(p, segs[-1][0], ctx)
    ppr = p.find(_w("pPr"))
    return 1 if ppr is not None else 0


def _check_text(p, expected: str | None) -> None:
    if expected is not None and para_text(p) != expected:
        raise EditError("正文与面板显示不一致（文件可能已在外部修改），已刷新，请重新选择", "stale", 409)


# ---------- 批注 ----------

def _comments_root(ctx: Ctx):
    return ctx.pkg.ensure_part("comments")


def _find_comment(ctx: Ctx, cid: str):
    root = ctx.pkg.xml("word/comments.xml")
    c = None if root is None else next((c for c in root.iter(_w("comment")) if _attr(c, "id") == cid), None)
    if c is None:
        raise EditError("批注不存在（可能已被删除），请刷新", "stale", 409)
    return c


def _last_para_id(c, ctx: Ctx, create: bool = True) -> str | None:
    paras = c.findall(_w("p"))
    if not paras:
        if not create:
            return None
        paras = [etree.SubElement(c, _w("p"))]
    pid = _attr(paras[-1], "paraId", W14)
    if not pid and create:
        pid = _para_id()
        paras[-1].set(f"{{{W14}}}paraId", pid)
        ctx.pkg.touch("word/comments.xml")
    return pid


def _comment_el(ctx: Ctx, cid: str, text: str):
    c = etree.Element(_w("comment"), {_w("id"): cid, _w("author"): ctx.author, _w("date"): ctx.date,
                                      _w("initials"): ctx.author[:2]})
    lines = text.split("\n") or [""]
    for i, line in enumerate(lines):
        p = etree.SubElement(c, _w("p"))
        p.set(f"{{{W14}}}paraId", _para_id())
        p.set(f"{{{W14}}}textId", "77777777")
        if i == 0:
            etree.SubElement(etree.SubElement(p, _w("r")), _w("annotationRef"))
        t = etree.SubElement(etree.SubElement(p, _w("r")), _w("t"))
        t.text = line
        t.set(XML_SPACE, "preserve")
    return c


def _ref_run(cid: str):
    r = etree.Element(_w("r"))
    etree.SubElement(r, _w("commentReference"), {_w("id"): cid})
    return r


def _commentex(ctx: Ctx, para_id: str, parent_pid: str | None = None, done: bool | None = None):
    ext = ctx.pkg.ensure_part("ext")
    ce = next((e for e in ext.iter(f"{{{W15}}}commentEx") if e.get(f"{{{W15}}}paraId") == para_id), None)
    if ce is None:
        ce = etree.SubElement(ext, f"{{{W15}}}commentEx", {f"{{{W15}}}paraId": para_id})
    if parent_pid:
        ce.set(f"{{{W15}}}paraIdParent", parent_pid)
    if done is not None:
        ce.set(f"{{{W15}}}done", "1" if done else "0")
    ctx.pkg.touch("word/commentsExtended.xml")
    return ce


def add_comment(ctx: Ctx, start: dict, end: dict, text: str, ptexts: dict) -> str:
    p1, p2 = _paragraph(int(start["p"]), ctx.body), _paragraph(int(end["p"]), ctx.body)
    _check_text(p1, ptexts.get(str(start["p"])))
    _check_text(p2, ptexts.get(str(end["p"])))
    cid = ctx.new_id()
    # 先放结束标记：标记零宽，不影响起点的字符偏移
    i = split_point(p2, int(end["o"]), ctx)
    p2.insert(i, etree.Element(_w("commentRangeEnd"), {_w("id"): cid}))
    p2.insert(i + 1, _ref_run(cid))
    i = split_point(p1, int(start["o"]), ctx)
    p1.insert(i, etree.Element(_w("commentRangeStart"), {_w("id"): cid}))
    root = _comments_root(ctx)
    c = _comment_el(ctx, cid, text)
    root.append(c)
    ctx.pkg.touch("word/comments.xml")
    _commentex(ctx, _last_para_id(c, ctx), done=False)
    ctx.pkg.touch("word/document.xml")
    return cid


def reply(ctx: Ctx, parent_cid: str, text: str) -> str:
    parent = _find_comment(ctx, parent_cid)
    parent_pid = _last_para_id(parent, ctx)
    cid = ctx.new_id()
    c = _comment_el(ctx, cid, text)
    parent.getparent().append(c)
    ctx.pkg.touch("word/comments.xml")
    _commentex(ctx, parent_pid)
    _commentex(ctx, _last_para_id(c, ctx), parent_pid=parent_pid, done=False)
    # 回复与父批注共用同一锚定区间
    for tag in ("commentRangeStart", "commentRangeEnd"):
        mark = next((m for m in ctx.doc.iter(_w(tag)) if _attr(m, "id") == parent_cid), None)
        if mark is not None:
            mark.addnext(etree.Element(_w(tag), {_w("id"): cid}))
    ref = next((r for r in ctx.doc.iter(_w("commentReference")) if _attr(r, "id") == parent_cid), None)
    if ref is not None:
        ref.getparent().addnext(_ref_run(cid))
    ctx.pkg.touch("word/document.xml")
    return cid


def resolve(ctx: Ctx, cid: str, done: bool) -> None:
    c = _find_comment(ctx, cid)
    _commentex(ctx, _last_para_id(c, ctx), done=done)


def delete_comment(ctx: Ctx, cid: str) -> int:
    root = ctx.pkg.xml("word/comments.xml")
    target = _find_comment(ctx, cid)
    ext = ctx.pkg.xml("word/commentsExtended.xml")
    pid_of = {}
    for c in root.iter(_w("comment")):
        pid_of[_attr(c, "id")] = _last_para_id(c, ctx, create=False)
    parent_of = {}
    if ext is not None:
        for ce in ext.iter(f"{{{W15}}}commentEx"):
            parent_of[ce.get(f"{{{W15}}}paraId")] = ce.get(f"{{{W15}}}paraIdParent")
    # 连同所有回复（递归）一起删
    doomed = {_attr(target, "id")}
    changed = True
    while changed:
        changed = False
        doomed_now = {pid_of[d] for d in doomed if pid_of.get(d)}
        for c_id, pid in pid_of.items():
            if c_id not in doomed and pid and parent_of.get(pid) in doomed_now:
                doomed.add(c_id)
                changed = True
    doomed_pids = {pid_of[d] for d in doomed if pid_of.get(d)}
    for c in list(root.iter(_w("comment"))):
        if _attr(c, "id") in doomed:
            c.getparent().remove(c)
    if ext is not None:
        for ce in list(ext.iter(f"{{{W15}}}commentEx")):
            if ce.get(f"{{{W15}}}paraId") in doomed_pids:
                ext.remove(ce)
        ctx.pkg.touch("word/commentsExtended.xml")
    _drop_durable(ctx, doomed_pids)
    for tag in ("commentRangeStart", "commentRangeEnd"):
        for m in list(ctx.doc.iter(_w(tag))):
            if _attr(m, "id") in doomed:
                m.getparent().remove(m)
    for ref in list(ctx.doc.iter(_w("commentReference"))):
        if _attr(ref, "id") in doomed:
            r = ref.getparent()
            r.remove(ref)
            if all(etree.QName(c).localname == "rPr" for c in r):
                r.getparent().remove(r)
    ctx.pkg.touch("word/comments.xml")
    ctx.pkg.touch("word/document.xml")
    return len(doomed)


def _drop_durable(ctx: Ctx, para_ids: set) -> None:
    """清理 commentsIds / commentsExtensible 中对应条目，避免孤儿引用。"""
    ids_root = ctx.pkg.xml("word/commentsIds.xml")
    if ids_root is None:
        return
    durable = set()
    for e in list(ids_root.iter(f"{{{W16CID}}}commentId")):
        if e.get(f"{{{W16CID}}}paraId") in para_ids:
            durable.add(e.get(f"{{{W16CID}}}durableId"))
            ids_root.remove(e)
    ctx.pkg.touch("word/commentsIds.xml")
    cex = ctx.pkg.xml("word/commentsExtensible.xml")
    if cex is not None and durable:
        for e in list(cex.iter(f"{{{W16CEX}}}commentExtensible")):
            if e.get(f"{{{W16CEX}}}durableId") in durable:
                cex.remove(e)
        ctx.pkg.touch("word/commentsExtensible.xml")


# ---------- 修订 ----------

def _unwrap(el) -> None:
    parent = el.getparent()
    i = parent.index(el)
    for ch in list(el):
        parent.insert(i, ch)
        i += 1
    parent.remove(el)


def _undelete(el) -> None:
    for t in list(el.iter(_w("delText"))):
        t.tag = _w("t")
    for t in list(el.iter(_w("delInstrText"))):
        t.tag = _w("instrText")


def _attached(el, root) -> bool:
    top = el
    while top.getparent() is not None:
        top = top.getparent()
    return top is root


def apply_rev(el, accept: bool) -> None:
    kind = etree.QName(el).localname
    inserted = kind in ("ins", "moveTo")
    if inserted == accept:      # 接受插入 / 拒绝删除：保留内容
        if not inserted:
            _undelete(el)
        _unwrap(el)
    else:                        # 拒绝插入 / 接受删除：去掉内容
        el.getparent().remove(el)


def revisions(ctx: Ctx, indices: list[int], accept: bool) -> int:
    revs = body_revisions(ctx.body)
    try:
        targets = [revs[int(i)] for i in indices]
    except (IndexError, ValueError):
        raise EditError("修订定位失败，请刷新后重试", "stale", 409)
    n = 0
    for el in targets:
        if _attached(el, ctx.doc):
            apply_rev(el, accept)
            n += 1
    ctx.pkg.touch("word/document.xml")
    return n


def _merge_with_next(p) -> None:
    nxt = p.getnext()
    if nxt is None or nxt.tag != _w("p"):
        return
    for ch in list(nxt):
        if ch.tag != _w("pPr"):
            p.append(ch)
    nxt.getparent().remove(nxt)


def all_revisions(ctx: Ctx, accept: bool) -> int:
    n = 0
    for el in body_revisions(ctx.body):
        if _attached(el, ctx.doc):
            apply_rev(el, accept)
            n += 1
    # 段落标记修订：接受删除 / 拒绝插入 → 与下一段合并
    for p in reversed(body_paragraphs(ctx.body)):
        rpr = p.find(f"{_w('pPr')}/{_w('rPr')}")
        if rpr is None:
            continue
        for mark in list(rpr):
            name = etree.QName(mark).localname
            if name in ("ins", "del", "moveTo", "moveFrom"):
                rpr.remove(mark)
                n += 1
                if (name in ("del", "moveFrom")) == accept:
                    _merge_with_next(p)
    # 表格行的插入/删除
    for tr in list(ctx.body.iter(_w("tr"))):
        trpr = tr.find(_w("trPr"))
        if trpr is None:
            continue
        for mark in list(trpr):
            name = etree.QName(mark).localname
            if name in ("ins", "del"):
                n += 1
                if (name == "del") == accept:
                    tr.getparent().remove(tr)
                    break
                trpr.remove(mark)
    # 格式修订：接受=丢弃旧格式记录；拒绝=恢复旧格式
    for tag in ("rPrChange", "pPrChange"):
        for ch in list(ctx.body.iter(_w(tag))):
            n += 1
            holder = ch.getparent()
            if not accept:
                old = ch.find(_w(tag[:-6]))
                for c in list(holder):
                    if c is not ch and not (tag == "pPrChange" and c.tag in (_w("rPr"), _w("sectPr"))):
                        holder.remove(c)
                # 旧属性放回原位：rPr 中排在变更记录之前；pPr 中排在 rPr/sectPr 之前（schema 顺序）
                if old is not None:
                    at = holder.index(ch) if tag == "rPrChange" else 0
                    for i, c in enumerate(list(old)):
                        holder.insert(at + i, c)
            holder.remove(ch)
    for tag in ("sectPrChange", "tblPrChange", "trPrChange", "tcPrChange", "tblGridChange", "numberingChange",
                "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"):
        for ch in list(ctx.body.iter(_w(tag))):
            ch.getparent().remove(ch)
    ctx.pkg.touch("word/document.xml")
    return n


# ---------- 审阅通过后自动删批注 ----------

_SHOWN = {"tab", "br", "cr", "sym", "noBreakHyphen", "drawing", "pict", "object", "footnoteReference", "endnoteReference"}


def _visible(rev) -> bool:
    """修订里有面板看得见的内容（文字、图片、公式等）；只含域代码字符或批注锚点的不算。"""
    for el in rev.iter():
        if not isinstance(el.tag, str):
            continue
        name = etree.QName(el).localname
        if name in ("t", "delText") and el.text:
            return True
        if name in _SHOWN or el.tag == f"{{{M}}}oMath":
            return True
    return False


def _flow_para(el):
    return next((a for a in el.iterancestors(_w("p")) if review._in_flow(a)), None)


def _epoch(s: str | None) -> float | None:
    """批注 / 修订日期 → 时间戳；带 Z 按 UTC，不带时区按本地时间（与 WPS / Word 一致）。"""
    if not s:
        return None
    try:
        dt = datetime.strptime(s[:19], "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        return None
    return dt.replace(tzinfo=timezone.utc).timestamp() if s.endswith("Z") else dt.timestamp()


def _comment_spans(body, pidx: dict) -> dict[str, tuple[int, int]]:
    """批注 id → 锚定区间覆盖的段落编号范围（按文档顺序遍历起止标记）。"""
    spans: dict[str, list[int]] = {}
    cur = -1
    for el in body.iter():
        if el.tag == _w("p") and el in pidx:
            cur = pidx[el]
        elif el.tag in (_w("commentRangeStart"), _w("commentRangeEnd"), _w("commentReference")):
            cid = _attr(el, "id")
            if cid is not None and cur >= 0:
                lo, hi = spans.get(cid, [cur, cur])
                spans[cid] = [min(lo, cur), max(hi, cur)]
    return {k: (v[0], v[1]) for k, v in spans.items()}


def clean_approved(ctx: Ctx, accepted: list) -> int:
    """接受修订后：本次动过的段落若已没有看得见的待审修订，视为整段审阅通过——
    同段残留的不可见插入一并接受，并删除锚定在这些段落、区间内也已无待审修订的批注（连同回复）。
    批注里只要有一条不是修订作者写的、且晚于本次接受的修订，就当作新意见保留。"""
    touched = {p for _, p, _, _ in accepted if p is not None}
    if not touched:
        return 0
    authors = {a for _, _, a, _ in accepted if a}
    dates = [d for _, _, _, d in accepted if d is not None]
    newest = max(dates) if dates else None
    paras = body_paragraphs(ctx.body)
    pidx = {p: i for i, p in enumerate(paras)}
    pending = set()
    for e in body_revisions(ctx.body):
        p = _flow_para(e)
        if p is not None and _visible(e):
            pending.add(pidx[p])
    approved = {pidx[p] for p in touched if p in pidx and pidx[p] not in pending}
    if not approved:
        return 0
    for i in approved:
        for e in list(paras[i].iter(_w("ins"), _w("moveTo"))):
            if etree.QName(e.getparent()).localname not in ("rPr", "pPr", "trPr") and not _visible(e) \
                    and _attached(e, ctx.doc):
                apply_rev(e, True)
    root = ctx.pkg.xml("word/comments.xml")
    if root is None:
        return 0
    ext = ctx.pkg.xml("word/commentsExtended.xml")
    parent_pid = {}
    if ext is not None:
        for ce in ext.iter(f"{{{W15}}}commentEx"):
            parent_pid[ce.get(f"{{{W15}}}paraId")] = ce.get(f"{{{W15}}}paraIdParent")
    info, cid_of_pid = {}, {}
    for c in root.iter(_w("comment")):
        cid = _attr(c, "id")
        pid = _last_para_id(c, ctx, create=False)
        info[cid] = (c, pid)
        if pid:
            cid_of_pid[pid] = cid

    def thread_root(cid):
        seen = set()
        while cid not in seen:
            seen.add(cid)
            par = cid_of_pid.get(parent_pid.get(info[cid][1]))
            if par is None or par not in info:
                return cid
            cid = par
        return cid

    threads: dict[str, list[str]] = {}
    for cid in info:
        threads.setdefault(thread_root(cid), []).append(cid)
    spans = _comment_spans(ctx.body, pidx)
    removed = 0
    for top, members in threads.items():
        rng = [spans[m] for m in members if m in spans]
        if not rng:
            continue
        lo, hi = min(r[0] for r in rng), max(r[1] for r in rng)
        if not any(lo <= i <= hi for i in approved) or any(lo <= i <= hi for i in pending):
            continue
        fresh = False
        for m in members:
            c = info[m][0]
            if (_attr(c, "author") or "") in authors:
                continue
            when = _epoch(_attr(c, "date"))
            if when is None or newest is None or when > newest:
                fresh = True
                break
        if not fresh:
            removed += delete_comment(ctx, top)
    # 删除锚点后留下的空修订包裹
    for i in approved:
        for e in list(paras[i].iter(_w("ins"), _w("del"))):
            if len(e) == 0 and etree.QName(e.getparent()).localname not in ("rPr", "pPr", "trPr"):
                e.getparent().remove(e)
    ctx.pkg.touch("word/document.xml")
    return removed


# ---------- 选中改写（以修订痕迹写入） ----------

def _to_deleted(node) -> None:
    for t in list(node.iter(_w("t"))):
        t.tag = _w("delText")
        t.set(XML_SPACE, "preserve")
    for t in list(node.iter(_w("instrText"))):
        t.tag = _w("delInstrText")


def _mark_deleted(nodes: list, ctx: Ctx) -> None:
    group = None
    for node in nodes:
        name = etree.QName(node).localname
        if name == "r":
            if group is None or group.getnext() is not node:
                group = ctx.rev("del")
                node.addprevious(group)
            _to_deleted(node)
            group.append(node)
            continue
        group = None
        if name in ("ins", "moveTo"):
            # 只包住直属的 run；已有的 w:del 子元素不能再嵌进新的 w:del（Word 会报“无法读取内容”）
            inner = None
            for ch in list(node):
                if ch.tag == _w("r"):
                    if inner is None or inner.getnext() is not ch:
                        inner = ctx.rev("del")
                        ch.addprevious(inner)
                    _to_deleted(ch)
                    inner.append(ch)
                else:
                    inner = None
        elif name in ("hyperlink", "smartTag", "customXml"):
            _mark_deleted([c for c in node if not etree.QName(c).localname.endswith("Pr")], ctx)


def rewrite(ctx: Ctx, pindex: int, s: int, e: int, ptext: str, new_text: str) -> None:
    p = _paragraph(pindex, ctx.body)
    _check_text(p, ptext)
    if not 0 <= s < e <= len(ptext):
        raise EditError("选区无效")
    end_mark, start_mark = etree.Element("sentinel"), etree.Element("sentinel")
    p.insert(split_point(p, e, ctx), end_mark)
    p.insert(split_point(p, s, ctx), start_mark)
    between = []
    node = start_mark.getnext()
    while node is not None and node is not end_mark:
        between.append(node)
        node = node.getnext()
    first_run = next((r for n in between for r in ([n] if n.tag == _w("r") else n.iter(_w("r")))
                      if not any(a.tag == _w("del") for a in r.iterancestors())), None)
    _mark_deleted(between, ctx)
    if new_text:
        ins = ctx.rev("ins")
        r = etree.SubElement(ins, _w("r"))
        rpr = first_run.find(_w("rPr")) if first_run is not None else None
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        for i, line in enumerate(new_text.split("\n")):
            if i:
                etree.SubElement(r, _w("br"))
            t = etree.SubElement(r, _w("t"))
            t.text = line
            t.set(XML_SPACE, "preserve")
        end_mark.addprevious(ins)
    p.remove(start_mark)
    p.remove(end_mark)
    ctx.pkg.touch("word/document.xml")


def rewrite_range(ctx: Ctx, start: dict, end: dict, ptexts: dict, new_text: str) -> None:
    """Rewrite adjacent paragraphs without changing their paragraph boundaries."""
    first, last = int(start['p']), int(end['p'])
    if first >= last:
        raise EditError('跨段选区无效')
    paras = body_paragraphs(ctx.body)
    if first < 0 or last >= len(paras):
        raise EditError('段落定位失败，请刷新后重试', 'stale', 409)
    selected = paras[first:last + 1]
    if any(a.getnext() is not b for a, b in zip(selected, selected[1:])):
        raise EditError('不能跨表格或其他非连续段落改写')
    lines = new_text.replace('\r\n', '\n').split('\n')
    if len(lines) != len(selected):
        raise EditError(f'请保留 {len(selected)} 个段落（每段之间用一个换行分隔）')
    def py_offset(text: str, offset: int) -> int:
        units = 0
        for pos, char in enumerate(text):
            if units == offset:
                return pos
            units += 2 if ord(char) > 0xffff else 1
            if units > offset:
                break
        if units == offset:
            return len(text)
        raise EditError('选区字符偏移无效，请重新选择', 'stale', 409)

    chunks = []
    for i, p in enumerate(selected):
        expected = ptexts.get(str(first + i))
        if expected is None:
            raise EditError('段落信息不完整，请重新选择', 'stale', 409)
        _check_text(p, expected)
        s = py_offset(expected, int(start['o'])) if i == 0 else 0
        e = py_offset(expected, int(end['o'])) if i == len(selected) - 1 else len(expected)
        if not 0 <= s <= e <= len(expected) or (s == e and lines[i]):
            raise EditError('选区中有空段落或无效偏移，暂不支持改写')
        chunks.append((first + i, s, e, expected, lines[i]))
    if not any(s < e for _, s, e, _, _ in chunks):
        raise EditError('请选中要改写的文字')
    for index, s, e, expected, line in chunks:
        if s < e:
            rewrite(ctx, index, s, e, expected, line)


# ---------- 入口 ----------

def run(did: str, body: dict) -> dict:
    path = review.get_path(did)
    if path is None or not path.exists():
        raise EditError("文档不存在或已移动", "gone", 404)
    if body.get("mtime") is not None and abs(path.stat().st_mtime - float(body["mtime"])) > 1e-6:
        raise EditError("文件已在外部修改，已为你刷新，请重新操作", "stale", 409)
    author = (body.get("author") or "").strip()
    if not author:
        raise EditError("请先设置署名")
    busy = holders(path)
    if busy and not body.get("force"):
        raise EditError(f"{'、'.join(busy)} 正打开此文件", "locked", 409)
    pkg = Package(path)
    ctx = Ctx(pkg, author)
    op = body.get("op")
    result: dict = {"op": op, "holders": busy}
    if op == "add_comment":
        text = (body.get("text") or "").strip()
        if not text:
            raise EditError("批注内容不能为空")
        result["cid"] = add_comment(ctx, body["start"], body["end"], text, body.get("ptexts") or {})
    elif op == "reply":
        text = (body.get("text") or "").strip()
        if not text:
            raise EditError("回复内容不能为空")
        result["cid"] = reply(ctx, str(body["cid"]), text)
    elif op == "resolve":
        resolve(ctx, str(body["cid"]), bool(body.get("done", True)))
    elif op == "delete_comment":
        result["count"] = delete_comment(ctx, str(body["cid"]))
    elif op == "rev":
        accept = body.get("action") == "accept"
        accepted = []
        if accept and body.get("auto_clean"):
            revs = body_revisions(ctx.body)
            for i in body.get("indices") or []:
                if isinstance(i, int) or str(i).isdigit():
                    el = revs[int(i)] if int(i) < len(revs) else None
                    if el is not None:
                        accepted.append((el, _flow_para(el), _attr(el, "author"), _epoch(_attr(el, "date"))))
        result["count"] = revisions(ctx, body.get("indices") or [], accept)
        if accepted:
            result["cleaned"] = clean_approved(ctx, accepted)
    elif op == "rev_all":
        result["count"] = all_revisions(ctx, body.get("action") == "accept")
    elif op == "rewrite":
        if body.get('start') and body.get('end'):
            rewrite_range(ctx, body['start'], body['end'], body.get('ptexts') or {}, body.get('text', ''))
        else:
            rewrite(ctx, int(body["p"]), int(body["s"]), int(body["e"]), body.get("ptext", ""),
                    body.get("text", ""))
    else:
        raise EditError("未知操作")
    if not pkg.dirty:
        raise EditError("没有可写入的改动")
    result["mtime"] = pkg.save(did)
    return result


def _step(did: str, src_dir: Path, dst_dir: Path, empty_msg: str) -> dict:
    """撤销 / 重做共用：当前文件先存入对侧栈，再用本栈最新一份覆盖。"""
    path = review.get_path(did)
    if path is None or not path.exists():
        raise EditError("文档不存在或已移动", "gone", 404)
    stack = _by_created(src_dir.glob("*.docx"))
    if not stack:
        raise EditError(empty_msg)
    last = review._load_index()["docs"].get(did, {}).get("last_write_mtime")
    # 外部（WPS）改过就不动，免得把那些改动覆盖掉；不弹窗，由前端提示
    if last is None or abs(path.stat().st_mtime - last) > 1e-6:
        raise EditError("上次面板写入之后文件在 WPS 等处改过，为避免覆盖这些改动，已阻止", "external", 409)
    backup(did, path, dst_dir)
    src = stack[-1]
    shutil.copyfile(src, path)
    src.unlink()
    mtime = path.stat().st_mtime
    _set_last_write(did, mtime)
    return {"mtime": mtime, "left": len(stack) - 1}


def undo(did: str, force: bool = False) -> dict:
    return _step(did, _backup_dir(did), _redo_dir(did), "没有可撤销的面板操作")


def redo(did: str) -> dict:
    return _step(did, _redo_dir(did), _backup_dir(did), "没有可重做的操作")
