TRANSLATE_TEXT_TMPL = "请将以下内容翻译为{target}，只输出译文：\n\n{text}"
TERM_TMPL = "请解释这段文字中的关键学术术语（中英对照、简明定义、在本文语境中的含义），Markdown 列表输出：\n\n{text}"
QUICK_SUM_TMPL = "请用简体中文概括这段论文内容（3-5 个要点）：\n\n{text}"

import json
import os

import httpx
from fastapi.responses import StreamingResponse


def glossary_text():
    return ''


def resolve_provider(cfg, task):
    provider = cfg.get('ai') or {}
    return provider, provider.get('model')


async def complete(provider, model, system, messages, max_tokens=4096, task='chat'):
    base = provider.get('base_url', '').rstrip('/')
    key = os.environ.get('PAPERCC_API_KEY')
    if not base or not model or not key:
        raise ValueError('请配置 AI 地址、模型和 PAPERCC_API_KEY，再重启 PaperCC。')
    try:
        async with httpx.AsyncClient(timeout=90, follow_redirects=False) as client:
            response = await client.post(
                base + '/chat/completions', headers={'Authorization': f'Bearer {key}'},
                json={'model': model, 'messages': [{'role': 'system', 'content': system}, *messages]})
            response.raise_for_status()
    except httpx.HTTPStatusError as exc:
        raise ValueError(f'模型接口返回 HTTP {exc.response.status_code}；请检查配置。') from None
    except httpx.RequestError:
        raise ValueError('模型连接失败或超时；请检查配置。') from None
    try:
        text = response.json()['choices'][0]['message']['content']
        if not isinstance(text, str) or not text.strip():
            raise ValueError
    except (KeyError, IndexError, TypeError, ValueError):
        raise ValueError('模型未返回可用文本。') from None
    return text


async def collect(cfg, task, system, user):
    provider, model = resolve_provider(cfg, task)
    try:
        text = await complete(provider, model, system, [{'role': 'user', 'content': user}])
    except ValueError as exc:
        yield {'error': str(exc)}
        return
    yield {'delta': text}
    yield {'full': text}


def sse(gen):
    async def frames():
        async for chunk in gen:
            yield f'data: {json.dumps(chunk, ensure_ascii=False)}\n\n'
        yield 'data: {"done": true}\n\n'
    return StreamingResponse(frames(), media_type='text/event-stream', headers={'Cache-Control': 'no-store'})
