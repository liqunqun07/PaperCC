import json
import os
from pathlib import Path
import socket
import subprocess
import sys
import tempfile
import threading
import time
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from docx import Document
import httpx


class Model(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def do_POST(self):
        request = json.loads(self.rfile.read(int(self.headers['Content-Length'])))
        assert self.path == '/v1/chat/completions' and request['model'] == 'fixture'
        assert self.headers['Authorization'] == 'Bearer test-placeholder'
        assert any('Add scope' in msg['content'] for msg in request['messages'])
        data = json.dumps({'choices': [{'message': {'content': 'Scope response fixture.'}}]}).encode()
        self.send_response(200)
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)


class ReviewTest(unittest.TestCase):
    def test_end_to_end(self):
        with tempfile.TemporaryDirectory() as tmp, ThreadingHTTPServer(('127.0.0.1', 0), Model) as model:
            thread = threading.Thread(target=model.serve_forever, daemon=True)
            thread.start()
            self.addCleanup(model.shutdown)
            data = Path(tmp)
            (data / 'config.json').write_text(json.dumps({'ai': {'base_url': f'http://127.0.0.1:{model.server_port}/v1', 'model': 'fixture'}}))
            with socket.socket() as sock:
                sock.bind(('127.0.0.1', 0))
                port = sock.getsockname()[1]
            origin = f'http://127.0.0.1:{port}'
            env = dict(os.environ, PAPERCC_PORT=str(port), PAPERCC_DATA_DIR=tmp,
                       PAPERCC_API_KEY='test-placeholder', NO_PROXY='127.0.0.1', no_proxy='127.0.0.1')
            with subprocess.Popen([sys.executable, '-m', 'server.main'], env=env) as server:
                try:
                    with httpx.Client(base_url=origin, headers={'Origin': origin}, trust_env=False) as client:
                        self.exercise(client, data)
                finally:
                    server.terminate()
                    server.wait(timeout=10)

    def exercise(self, client, data):
        for _ in range(50):
            try:
                page = client.get('/review.html')
                break
            except httpx.ConnectError:
                time.sleep(.1)
        else:
            self.fail('service did not start')
        self.assertIn('PaperCC', page.text)
        self.assertEqual(client.get('/js/review.js').status_code, 200)
        self.assertEqual(client.get('/api/review/recent', headers={'Host': 'other.example'}).status_code, 403)
        self.assertEqual(client.post('/api/review/open', headers={'Origin': 'https://other.example'}, json={}).status_code, 403)

        path = data / 'fixture.docx'
        doc = Document()
        run = doc.add_paragraph('Original academic statement.').runs[0]
        doc.add_comment(run, 'Add scope', 'Reviewer', 'R')
        doc.save(path)
        def post(url, **payload):
            response = client.post('/api/review/' + url, json=payload)
            self.assertTrue(response.is_success, response.text)
            return response.json()
        did = post('open', path=str(path))['id']
        def current():
            return client.get(f'/api/review/doc/{did}').raise_for_status().json()
        before = current()
        comment = before['comments'][0]
        self.assertEqual(comment['anchor'], 'Original academic statement.')
        post('edit', id=did, mtime=before['mtime'], author='Test', op='rewrite', p=0,
             s=0, e=8, ptext='Original academic statement.', text='Revised')
        self.assertIn('<ins ', current()['html'])
        post('edit', id=did, author='Test', op='rev_all', action='accept')
        self.assertIn('Revised academic statement.', Document(path).paragraphs[0].text)
        post('undo', id=did)
        self.assertIn('<ins ', current()['html'])
        post('redo', id=did)
        self.assertNotIn('<ins ', current()['html'])
        post('jobs', doc_id=did, items=[{'cid': comment['id'], 'op': 'reply'}])
        for _ in range(50):
            job = client.get('/api/review/jobs', params={'doc_id': did}).json()['jobs'][0]
            if job['status'] in ('done', 'error'):
                break
            time.sleep(.1)
        self.assertEqual(job['result'], 'Scope response fixture.', job)
        exported = post('responses/export', doc_name='fixture.docx', items=[{'comment': 'Add scope', 'response': job['result']}])
        self.assertIn(job['result'], Path(exported['markdown']).read_text())
        self.assertIn(job['result'], [p.text for p in Document(exported['docx']).paragraphs])
