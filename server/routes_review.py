"""审阅模式路由：浏览批注版 .docx、划词/批注/段落翻译与解释、在 Word 中打开。"""
from __future__ import annotations

import asyncio
import subprocess

from fastapi import APIRouter, HTTPException
from fastapi.responses import JSONResponse, Response

from . import llm, review, review_edit, review_jobs, store
from .llm import collect as _collect, sse

router = APIRouter()

EXPLAIN_COMMENT_TMPL = (
    "下面是论文修改稿中的一条审阅批注，以及它所锚定的原文。请用简体中文简明说明：\n"
    "1. 这条批注要求做什么（一两句话）；\n2. 具体改哪里、建议怎么改（要点列表）；\n"
    "3. 需要作者额外补充或确认的材料（没有就写“无”）。\n"
    "保留批注中的审稿人编号（如 R1-9）与术语原词。\n\n"
    "【批注】\n{comment}\n\n【锚定原文】\n{anchor}\n\n【所在段落】\n{context}"
)
EXPLAIN_SEL_TMPL = (
    "请解释论文修改稿中这段内容的含义，用简体中文，先一句话通俗解释，再补充背景/术语说明（如有）。\n"
    "【所选内容】\n{text}\n\n【所在段落（供参考）】\n{context}"
)


def _target_for(text: str) -> str:
    target = str(store.load_config().get("translate_target") or "中文").strip() or "中文"
    # 原文已是中文时反向译成英文，便于对照
    if target == "中文" and review.mostly_cjk(text):
        return "英文"
    return target


# 审阅稿里的变量名/条目编号（IN_1、PA、R1-9）必须原样保留，不套用面板的 LaTeX 规则
KEEP_RULE = "变量名、量表条目编号、审稿人编号、数字与引用一律保持原样，不要改写成 LaTeX 或加 $ 符号。\n"


def _sys_zh() -> str:
    return ("你是专业学术翻译。准确、流畅、符合中文学术表达习惯；专有名词首次出现时括注英文原词。"
            + llm.glossary_text() + "\n" + KEEP_RULE)


SYS_REVIEW = "你是论文修改助手。用简体中文说明，忠实原文、不编造；保留关键术语英文原词。输出使用 Markdown。\n" + KEEP_RULE

SYS_TRANSLATE_EN = (
    "You are a professional academic translator. Translate Chinese into accurate, concise academic English; "
    "keep reviewer codes (e.g. R1-9), variable names, numbers and citations unchanged. Output only the translation.\n"
)

SYS_REWRITE = (
    "You are an academic revising assistant. Rewrite the given passage strictly following the reviewer comment. "
    "Respond in the same language as the passage. Apply every change the comment asks for; keep the rest faithful "
    "to the original — do not add new claims or expand. Output ONLY the rewritten passage, no explanations, no quotes.\n"
    + KEEP_RULE
)
REWRITE_COMMENT_TMPL = (
    "请根据下面的审阅批注，改写【要改写的文字】：\n"
    "- 落实批注要求的全部修改；批注没涉及的内容保持原意，不新增论点、不扩写；\n"
    "- 与所在段落其余部分衔接自然，保持学术语体；\n"
    "- 只输出改写后的文字本身，不要任何解释或引号。\n\n"
    "【批注】\n{comment}\n\n【要改写的文字】\n{text}\n\n【所在段落（供参考）】\n{context}"
)


def _translate_prompt(text: str) -> tuple[str, str, str]:
    target = _target_for(text)
    system = SYS_TRANSLATE_EN if target == "英文" else _sys_zh()
    return system, llm.TRANSLATE_TEXT_TMPL.format(target=target, text=text), "translate"


def _prompt(action: str, text: str, context: str) -> tuple[str, str, str]:
    if action == "translate":
        return _translate_prompt(text)
    if action == "explain":
        return SYS_REVIEW, EXPLAIN_SEL_TMPL.format(text=text, context=context or "（无）"), "chat"
    if action == "term":
        return SYS_REVIEW, llm.TERM_TMPL.format(text=text), "chat"
    if action == "summarize":
        return SYS_REVIEW, llm.QUICK_SUM_TMPL.format(text=text), "chat"
    raise KeyError("action")


def _cached_stream(key: str | None, system: str, user: str, task: str):
    cfg = store.load_config()
    hit = review.cache_get(key) if key else None

    async def gen():
        if hit is not None:
            yield {"cached": True}
            yield {"delta": hit}
            yield {"full": hit}
            return
        async for chunk in _collect(cfg, task, system, user):
            if "full" in chunk and key:
                review.cache_put(key, chunk["full"])
            yield chunk

    return sse(gen())


@router.get("/api/review/recent")
async def review_recent():
    return review.recent()


@router.get("/api/review/folder")
async def review_folder(path: str):
    try:
        folder, files = await asyncio.to_thread(review.list_folder, path)
    except ValueError as e:
        raise HTTPException(400, str(e))
    review.remember_folder(str(folder))
    return {"folder": str(folder), "files": files}


@router.get("/api/review/dir")
async def review_dir(path: str):
    try:
        return await asyncio.to_thread(review.list_dir, path)
    except (ValueError, OSError) as e:
        raise HTTPException(400, str(e))


@router.post("/api/review/pick")
async def review_pick(body: dict):
    path = await asyncio.to_thread(review.pick, body.get("kind", "folder"))
    return {"path": path}


@router.post("/api/review/open")
async def review_open(body: dict):
    try:
        return review.register(body["path"])
    except ValueError as e:
        raise HTTPException(400, str(e))


@router.get("/api/review/doc/{did}")
async def review_doc(did: str):
    try:
        return await asyncio.to_thread(review.render, did)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))


@router.get("/api/review/doc/{did}/mtime")
async def review_mtime(did: str):
    path = review.get_path(did)
    if path is None or not path.exists():
        raise HTTPException(404, "文档不存在或已移动")
    return {"mtime": path.stat().st_mtime}


@router.get("/api/review/doc/{did}/media/{name}")
async def review_media(did: str, name: str):
    res = await asyncio.to_thread(review.media, did, name)
    if res is None:
        raise HTTPException(404, "图片不存在")
    return Response(content=res[0], media_type=res[1], headers={"Cache-Control": "max-age=3600"})


@router.post("/api/review/open-word")
async def review_open_word(body: dict):
    path = review.get_path(body.get("id", ""))
    if path is None or not path.exists():
        raise HTTPException(404, "文档不存在或已移动")
    r = subprocess.run(["open", "-b", "com.kingsoft.wpsoffice.mac", str(path)], capture_output=True, text=True)
    if r.returncode != 0:  # 没装 WPS 时交给系统默认程序
        r = subprocess.run(["open", str(path)], capture_output=True, text=True)
    if r.returncode != 0:
        raise HTTPException(500, (r.stderr or "打开失败")[:200])
    return {"ok": True}


@router.post("/api/review/quick")
async def review_quick(body: dict):
    """划词 / 段落：translate | explain | term | summarize。翻译结果持久缓存。"""
    text = (body.get("text") or "").strip()
    if not text:
        raise KeyError("text")
    action = body.get("action", "translate")
    context = (body.get("context") or "").strip()[:3000]
    system, user, task = _prompt(action, text, context)
    key = review.cache_key("quick", action, system, user) if action == "translate" else None
    return _cached_stream(key, system, user, task)


@router.post("/api/review/comment")
async def review_comment(body: dict):
    """单条批注：translate | explain，结果按内容持久缓存（文档改动后批注内容不变则复用）。"""
    did, cid = body.get("doc_id", ""), str(body.get("cid", ""))
    mode = body.get("mode", "translate")
    try:
        doc = await asyncio.to_thread(review.render, did)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    c = next((x for x in doc["comments"] if x["id"] == cid), None)
    if c is None:
        raise HTTPException(404, "批注不存在")
    if mode == "translate":
        system, user, task = _translate_prompt(c["text"])
    elif mode == "explain":
        context = (body.get("context") or "").strip()[:3000]
        system, user, task = SYS_REVIEW, EXPLAIN_COMMENT_TMPL.format(
            comment=c["text"], anchor=c["anchor"][:2000] or "（无）", context=context or "（无）"), "chat"
    else:
        raise KeyError("mode")
    return _cached_stream(review.cache_key("comment", mode, system, user), system, user, task)


@router.post("/api/review/comment/rewrite")
async def review_comment_rewrite(body: dict):
    """批注一键 AI 改写：按批注意见改写其锚定文字，流式返回预览；是否写回由前端「采用」决定。"""
    did, cid = body.get("doc_id", ""), str(body.get("cid", ""))
    text = (body.get("text") or "").strip()
    if not text:
        raise KeyError("text")
    try:
        doc = await asyncio.to_thread(review.render, did)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    c = next((x for x in doc["comments"] if x["id"] == cid), None)
    if c is None:
        raise HTTPException(404, "批注不存在")
    context = (body.get("context") or "").strip()[:3000]
    user = REWRITE_COMMENT_TMPL.format(comment=c["text"], text=text[:4000], context=context or "（无）")
    system = review_jobs._sys_with_skills(SYS_REWRITE, rewrite=True)
    return sse(_collect(store.load_config(), "rewrite", system, user))


# ---------- 原生写回（批注 / 修订 / 改写） ----------

def _edit_response(fn, *args):
    try:
        return fn(*args)
    except review_edit.EditError as e:
        return JSONResponse(status_code=e.status, content={"error": str(e), "code": e.code})


@router.post("/api/review/edit")
async def review_edit_op(body: dict):
    return await asyncio.to_thread(_edit_response, review_edit.run, body.get("id", ""), body)


@router.post("/api/review/undo")
async def review_undo(body: dict):
    return await asyncio.to_thread(_edit_response, review_edit.undo, body.get("id", ""), bool(body.get("force")))


@router.post("/api/review/redo")
async def review_redo(body: dict):
    return await asyncio.to_thread(_edit_response, review_edit.redo, body.get("id", ""))


@router.get("/api/review/doc/{did}/holders")
async def review_holders(did: str):
    path = review.get_path(did)
    if path is None or not path.exists():
        raise HTTPException(404, "文档不存在或已移动")
    return {"holders": await asyncio.to_thread(review_edit.holders, path)}


# ---------- 批注任务队列（解释 / 翻译 / 改写 / 回复草稿 / 深度委派） ----------

@router.post("/api/review/jobs")
async def jobs_enqueue(body: dict):
    doc_id = str(body.get("doc_id", ""))
    items = body.get("items") or []
    if not doc_id or not items:
        raise HTTPException(400, "doc_id 与 items 不能为空")
    try:
        return await review_jobs.enqueue(doc_id, items)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))


@router.get("/api/review/jobs")
async def jobs_list(doc_id: str):
    return review_jobs.list_jobs(doc_id)


@router.get("/api/review/jobs/summary")
async def jobs_summary():
    return review_jobs.summary()


@router.post("/api/review/jobs/cancel")
async def jobs_cancel(body: dict):
    return review_jobs.cancel([str(x) for x in (body.get("ids") or [])])


@router.post("/api/review/responses/export")
async def responses_export(body: dict):
    items = body.get("items") or []
    if not items:
        raise HTTPException(400, "没有可导出的回复")
    try:
        return await asyncio.to_thread(
            review_jobs.export_responses, str(body.get("doc_name") or "manuscript"), items)
    except Exception as e:
        raise HTTPException(500, str(e)[:200])
