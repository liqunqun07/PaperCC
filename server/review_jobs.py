"""批注任务队列：解释 / 翻译 / 改写 / 回复草稿 / 深度委派（agent CLI）。

设计：LLM 类任务（explain/translate/rewrite/reply）2 并发消费。
结果写进内存 JOBS 并复用 review 的持久缓存（重启不丢已完成的解释/翻译/改写）。
写回仍由前端在「采用」时逐条执行（走 /api/review/edit 的备份/撤销/占用保护），队列不直接改文件。
"""
from __future__ import annotations

import asyncio
import re
import time
import uuid
from pathlib import Path

from . import llm, review, store

OPS = ("explain", "translate", "rewrite", "reply")
LLM_WORKERS = 2
MAX_JOBS_PER_DOC = 60

JOBS: dict[str, dict] = {}
_seq = 0
_llm_queue: asyncio.Queue | None = None


def summary() -> dict:
    running = sum(1 for j in JOBS.values() if j["status"] == "running")
    queued = sum(1 for j in JOBS.values() if j["status"] == "queued")
    return {"running": running, "queued": queued}


def _new_job(doc_id: str, cid: str, op: str, comment_text: str, context: str, target: str) -> dict:
    global _seq
    _seq += 1
    jid = uuid.uuid4().hex[:10]
    JOBS[jid] = {
        "id": jid, "seq": _seq, "doc_id": doc_id, "cid": cid, "op": op,
        "status": "queued", "result": "", "error": "", "hedges": {}, "cached": False,
        "comment_text": comment_text, "context": context, "target": target,
        "created": time.time(), "started": 0.0, "finished": 0.0,
    }
    return JOBS[jid]


async def _ensure_workers() -> None:
    global _llm_queue
    if _llm_queue is None:
        _llm_queue = asyncio.Queue()
        loop = asyncio.get_running_loop()
        for _ in range(LLM_WORKERS):
            loop.create_task(_llm_worker())


async def enqueue(doc_id: str, items: list[dict]) -> dict:
    """items: [{cid, op, text?, context?}]；text 是改写目标文字（前端按当前视图口径取好）。"""
    await _ensure_workers()
    doc = await asyncio.to_thread(review.render, doc_id)
    by_id = {c["id"]: c for c in doc["comments"]}
    created, skipped = [], []
    for it in items:
        cid, op = str(it.get("cid", "")), it.get("op", "")
        c = by_id.get(cid)
        if c is None or op not in OPS:
            skipped.append({"cid": cid, "op": op, "reason": "批注不存在或操作不支持"})
            continue
        if op == "rewrite" and not (it.get("text") or "").strip():
            skipped.append({"cid": cid, "op": op, "reason": "缺少改写目标文字"})
            continue
        pending = sum(1 for j in JOBS.values() if j["doc_id"] == doc_id and j["cid"] == cid
                      and j["op"] == op and j["status"] in ("queued", "running", "cancelling"))
        if pending:
            skipped.append({"cid": cid, "op": op, "reason": "同样的任务已在队列"})
            continue
        active = sum(1 for j in JOBS.values() if j["doc_id"] == doc_id
                     and j["status"] in ("queued", "running", "cancelling"))
        if active >= MAX_JOBS_PER_DOC:
            skipped.append({"cid": cid, "op": op, "reason": "该文档排队已达上限"})
            continue
        job = _new_job(doc_id, cid, op, c["text"], (it.get("context") or "").strip()[:3000],
                       (it.get("text") or "").strip())
        await _llm_queue.put(job["id"])
        created.append(job)
    return {"jobs": created, "skipped": skipped}


def cancel(ids: list[str]) -> dict:
    n = 0
    for jid in ids:
        j = JOBS.get(jid)
        if j and j["status"] == "queued":
            j["status"] = "cancelled"
            j["finished"] = time.time()
            n += 1
        elif j and j["status"] == "running":
            j["status"] = "cancelling"
            n += 1
    return {"cancelled": n}


def list_jobs(doc_id: str) -> dict:
    """返回该文档的任务；已完成超过 120 条时淘汰最旧的（cancelled 直接淘汰）。"""
    jobs = sorted((j for j in JOBS.values() if j["doc_id"] == doc_id), key=lambda j: j["seq"])
    finished = [j for j in jobs if j["status"] in ("done", "error", "cancelled")]
    for j in finished[:-120]:
        JOBS.pop(j["id"], None)
    return {"jobs": [_public(j) for j in jobs if j["status"] != "cancelled"]}


def _public(j: dict) -> dict:
    return {k: j[k] for k in ("id", "seq", "cid", "op", "status", "result", "error", "hedges",
                              "comment_text", "target", "created", "finished", "cached")}


# ---------- 提示词（复用 routes_review 的模板，运行时导入避免环） ----------

def _route_mod():
    from . import routes_review
    return routes_review


REPLY_TMPL = (
    "请为下面的审稿批注起草一条回复（response to reviewer），将直接放进逐条回复信：\n"
    "- 逐条回应批注要求，仅依据材料说明可确认的修改，无法确认时用待核对占位符；若批注问的是问题，先直接回答再说明修改位置；\n"
    "- 语气礼貌、具体、不空洞；引用修改后的结论时忠实于提供的段落，不编造数据；\n"
    "- 只输出回复正文本身，不要标题、编号或解释。\n\n"
    "【批注】\n{comment}\n\n【批注锚定的原文】\n{anchor}\n\n【所在段落（供参考）】\n{context}"
)

REPLY_SYS_ZH = "你是论文返修助手。用简体中文起草审稿回复，忠实原文、不编造；保留审稿人编号与术语英文原词。输出纯文本。\n"
REPLY_SYS_EN = (
    "You are an academic response writer. Draft a concise, specific, polite response to the reviewer "
    "comment in English. Stay faithful to the provided manuscript text; never invent data or claims. "
    "Output only the response text. Conventional opening: thank the reviewer (We thank the reviewer ...), "
    "then address each point.\n"
)


def _build_llm_task(job: dict) -> tuple[str, str, str, str]:
    """返回 (system, user, task, cache_kind)。"""
    r = _route_mod()
    c_text, context = job["comment_text"], job["context"]
    if job["op"] == "explain":
        system, user, task = r.SYS_REVIEW, r.EXPLAIN_COMMENT_TMPL.format(
            comment=c_text, anchor=job["target"][:2000] or "（无）", context=context or "（无）"), "chat"
        return system, user, task, "comment"
    if job["op"] == "translate":
        system, user, _task = r._translate_prompt(c_text)
        return system, user, "translate", "comment"
    if job["op"] == "rewrite":
        user = r.REWRITE_COMMENT_TMPL.format(comment=c_text, text=job["target"][:4000], context=context or "（无）")
        return r.SYS_REWRITE, user, "rewrite", "rewrite"
    if job["op"] == "reply":
        zh = review.mostly_cjk(c_text)
        system = REPLY_SYS_ZH if zh else REPLY_SYS_EN
        user = REPLY_TMPL.format(comment=c_text, anchor=job["target"][:2000] or "（无）", context=context or "（无）")
        return system, user, "chat", "reply"
    raise ValueError(job["op"])


# ---------- LLM worker ----------

async def _llm_call(system: str, user: str, task: str) -> str:
    """带 300s 上限的 LLM 调用。关键点：用独立 task + asyncio.wait 超时，绝不把取消
    传播进 llm 的 async generator 链——否则 finally 里的 gen.aclose()/httpx 清理会
    卡死整个事件循环（uvloop _cancel_all_tasks 实测复现）。"""
    cfg = store.load_config()
    provider, model = llm.resolve_provider(cfg, task)
    t = asyncio.ensure_future(llm.complete(provider, model, system,
                                           [{"role": "user", "content": user}],
                                           cfg.get("max_tokens", 4096), task))
    done, pending = await asyncio.wait({t}, timeout=300)
    if pending:
        t.cancel()      # 只发出取消，不等待其清理结束；僵尸任务留在后台自行了断
        raise TimeoutError("LLM 调用超过 300s，已放弃（后台调用自行收尾）")
    return t.result()


async def _llm_worker() -> None:
    while True:
        jid = await _llm_queue.get()
        job = JOBS.get(jid)
        if not job:
            _llm_queue.task_done()
            continue
        if job["status"] in ("cancelling", "cancelled"):
            job["status"], job["finished"] = "cancelled", time.time()
            _llm_queue.task_done()
            continue
        job["status"], job["started"] = "running", time.time()
        try:
            system, user, task, kind = _build_llm_task(job)
            key = review.cache_key(kind, job["op"], system, user)
            hit = review.cache_get(key)
            if hit is not None:
                job["result"], job["cached"] = hit, True
                job["status"] = "done"
            else:
                out = (await _llm_call(system, user, task)).strip()
                if not out:
                    raise ValueError("模型返回空内容")
                review.cache_put(key, out)
                job["result"] = out
                job["status"] = "done"
        except Exception as e:
            job["status"], job["error"] = "error", str(e)[:300]
        job["finished"] = time.time()
        _llm_queue.task_done()


def export_responses(doc_name: str, items: list[dict]) -> dict:
    """items: [{comment, anchor, response}]（前端按批注顺序给）。落盘 md + docx。"""
    base = store._desktop_base() / "review"
    base.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:6]
    safe = re.sub(r'[\\/:*?"<>|]', "_", Path(doc_name).stem)[:60] or "responses"
    md_path = base / f"Response_{safe}_{stamp}.md"
    lines = [f"# Response to Reviewers · {doc_name}", ""]
    for i, it in enumerate(items, 1):
        lines += [f"## {i}. Comment", "", (it.get("comment") or "").strip(), ""]
        if (it.get("anchor") or "").strip():
            lines += [f"> 原文锚点：{it['anchor'].strip()[:300]}", ""]
        lines += ["### Response", "", (it.get("response") or "").strip(), ""]
    md_path.write_text("\n".join(lines), encoding="utf-8")

    from docx import Document
    doc = Document()
    doc.add_heading(f"Response to Reviewers — {doc_name}", level=0)
    for i, it in enumerate(items, 1):
        doc.add_heading(f"{i}. Reviewer Comment", level=2)
        doc.add_paragraph((it.get("comment") or "").strip())
        if (it.get("anchor") or "").strip():
            doc.add_paragraph(it["anchor"].strip()[:300], style="Intense Quote")
        doc.add_heading("Response", level=2)
        doc.add_paragraph((it.get("response") or "").strip())
    docx_path = base / f"Response_{safe}_{stamp}.docx"
    doc.save(str(docx_path))
    return {"markdown": str(md_path), "docx": str(docx_path), "count": len(items)}
