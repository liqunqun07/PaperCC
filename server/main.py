"""Loopback-only PaperCC service."""
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from . import store
from .routes_review import router

app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)


@app.middleware('http')
async def local_browser_only(request: Request, call_next):
    host = request.headers.get('host', '')
    if host.split(':', 1)[0] != '127.0.0.1':
        return JSONResponse({'detail': '请使用 127.0.0.1 访问。'}, status_code=403)
    origin = request.headers.get('origin')
    if (request.headers.get('sec-fetch-site') in ('cross-site', 'same-site')
            or (origin and origin != 'http://' + host)
            or (request.method == 'POST' and origin != 'http://' + host)):
        return JSONResponse({'detail': '仅允许本页面发起请求。'}, status_code=403)
    response = await call_next(request)
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['Content-Security-Policy'] = (
        "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; "
        "img-src 'self' data:; object-src 'none'; frame-ancestors 'none'; base-uri 'none'")
    return response


@app.get('/')
async def home():
    return RedirectResponse('/review.html')


app.include_router(router)
app.mount('/', StaticFiles(directory=Path(__file__).resolve().parents[1] / 'web'))

if __name__ == '__main__':
    import os
    import uvicorn

    os.umask(0o077)
    store.DATA.mkdir(parents=True, exist_ok=True, mode=0o700)
    uvicorn.run(app, host='127.0.0.1', port=int(os.environ.get('PAPERCC_PORT', '8765')),
                proxy_headers=False, access_log=False)
