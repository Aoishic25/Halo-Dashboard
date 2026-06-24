"""
HALO Dashboard Launcher
Double-click to start the dashboard. Browser opens automatically.
Close the console window to stop the server.
"""
import http.server
import json
import os
import re
import socketserver
import subprocess
import sys
import threading
import time
import urllib.request
import urllib.error
import webbrowser

PORT = 8765
ROOT = os.path.dirname(os.path.abspath(sys.argv[0])) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(ROOT, 'Ops files', 'Monthly Review Meeting.xlsx')
EXTRACT_SCRIPT = os.path.join(ROOT, '_extract_data.ps1')
DATA_BLOCK = os.path.join(ROOT, '_data_block.js')
CLOUDFLARED = os.path.join(ROOT, 'cloudflared.exe')

extract_status = {'running': False, 'last_run': 0, 'last_error': None, 'last_mtime': 0}
OLLAMA_URL = 'http://localhost:11434/api/chat'
OLLAMA_MODEL = 'llama3.2:3b'
GROQ_URL = 'https://api.groq.com/openai/v1/chat/completions'
GROQ_MODEL = 'llama-3.3-70b-versatile'
GROQ_KEY_FILE = os.path.join(ROOT, '.groq_key')
_data_context = ''
_full_data = {}

tunnel_url = ''
watcher_active = False
watcher_thread = None


def _num(v):
    try:
        return float(str(v).replace(',', ''))
    except (ValueError, TypeError):
        return 0


def build_data_context():
    global _data_context, _full_data
    if not os.path.exists(DATA_BLOCK):
        _data_context = 'No data loaded yet.'
        return
    try:
        with open(DATA_BLOCK, 'r', encoding='utf-8-sig') as f:
            raw = f.read()
        raw = re.sub(r'^var\s+EMBEDDED_DATA\s*=\s*', '', raw).rstrip().rstrip(';')
        _full_data = json.loads(raw)
        lines = []
        for key, rows in _full_data.items():
            if not isinstance(rows, list) or not rows:
                continue
            cols = list(rows[0].keys())
            lines.append(f'{key}: {len(rows)} rows, cols={",".join(cols[:5])}')
            num_totals = {}
            for c in cols[:6]:
                vals_list = [_num(r.get(c)) for r in rows]
                total = sum(vals_list)
                if total != 0 and any(v != 0 for v in vals_list[:3]):
                    num_totals[c] = f'{total:,.0f}'
            if num_totals:
                lines.append(f'  Totals: {num_totals}')
        _data_context = '\n'.join(lines)
        print(f'Data context built: {len(_data_context):,} chars')
    except Exception as e:
        _data_context = f'Error loading data: {e}'
        print(f'Error building data context: {e}')


def get_relevant_data(question):
    q = question.lower()
    keyword_map = {
        'stock': ['stockFlow'],
        'inbound': ['stockFlow', 'inboundFlow', 'inboundSummary', 'hubInbound'],
        'outbound': ['stockFlow', 'outboundCases', 'hubOutbound'],
        'cost': ['costingDet', 'costingCases'],
        'freight': ['freights'],
        'quality': ['qualityIssues'],
        'damage': ['damages', 'whDamages'],
        'complaint': ['customerComplaints'],
        'copacking': ['coPacking', 'copackingOrders', 'copackingWeekly'],
        'co-packing': ['coPacking', 'copackingOrders', 'copackingWeekly'],
        'order': ['weeklyOrders', 'manualOrders', 'copackingOrders'],
        'pallet': ['palletAging'],
        'waiting': ['waitingCharges'],
        'invoice': ['invoiceSummary'],
        'temp': ['tempReport'],
        'storage': ['storage', 'storageHub'],
        'expired': ['expiredStock'],
        'truck': ['truckTurnover'],
        'subject': ['subjects'],
        'action': ['subjects'],
        'vendor': ['hubInbound', 'qualityIssues', 'inboundIssues'],
        'ksa': ['freights', 'outboundCases', 'hubOutbound', 'waitingCharges'],
        'turkey': ['freights', 'outboundCases', 'hubOutbound'],
        'issue': ['qualityIssues', 'inboundIssues'],
    }
    matched = set()
    for kw, datasets in keyword_map.items():
        if kw in q:
            matched.update(datasets)
    if not matched:
        matched = set(list(_full_data.keys())[:5])

    lines = []
    for key in matched:
        rows = _full_data.get(key, [])
        if not rows:
            continue
        cols = list(rows[0].keys())
        lines.append(f'\n## {key} ({len(rows)} rows)')
        show = rows[-10:] if len(rows) > 10 else rows
        for r in show:
            vals = [f'{c}={r[c]}' for c in cols if r.get(c)]
            lines.append('  ' + ' | '.join(vals))
    return '\n'.join(lines)


def run_extraction():
    if extract_status['running']:
        return False, 'Extraction already running'
    extract_status['running'] = True
    extract_status['last_error'] = None
    try:
        result = subprocess.run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', EXTRACT_SCRIPT],
            capture_output=True, text=True, timeout=600, cwd=ROOT
        )
        if result.returncode == 0:
            extract_status['last_run'] = time.time()
            extract_status['last_mtime'] = os.path.getmtime(DATA_BLOCK) if os.path.exists(DATA_BLOCK) else 0
            return True, result.stdout
        else:
            extract_status['last_error'] = result.stderr or result.stdout
            return False, result.stderr or result.stdout
    except Exception as e:
        extract_status['last_error'] = str(e)
        return False, str(e)
    finally:
        extract_status['running'] = False


def run_extraction_bg():
    if extract_status['running']:
        return
    threading.Thread(target=run_extraction, daemon=True).start()


def watcher_loop():
    global watcher_active
    last_mtime = 0
    try:
        last_mtime = os.path.getmtime(EXCEL_PATH)
    except OSError:
        pass
    while watcher_active:
        time.sleep(3)
        try:
            mtime = os.path.getmtime(EXCEL_PATH)
            if mtime > last_mtime and not extract_status['running']:
                last_mtime = mtime
                time.sleep(2)
                run_extraction()
        except OSError:
            pass


def start_tunnel():
    global tunnel_url
    if not os.path.exists(CLOUDFLARED):
        return
    def _run():
        global tunnel_url
        proc = subprocess.Popen(
            [CLOUDFLARED, 'tunnel', '--url', f'http://localhost:{PORT}'],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True
        )
        for line in proc.stdout:
            m = re.search(r'(https://[a-z0-9-]+\.trycloudflare\.com)', line)
            if m:
                tunnel_url = m.group(1)
                print(f'\n  Public URL: {tunnel_url}\n')
    threading.Thread(target=_run, daemon=True).start()


class ThreadingHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True


class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def do_GET(self):
        if self.path == '/' or self.path == '':
            if os.path.exists(os.path.join(ROOT, 'index.html')):
                self.send_response(302)
                self.send_header('Location', '/index.html')
            else:
                self.send_response(302)
                self.send_header('Location', '/HALO_Dashboard.html')
            self.end_headers()
            return
        elif self.path == '/api/status':
            self.send_json({
                'running': extract_status['running'],
                'lastRun': extract_status['last_run'],
                'lastError': extract_status['last_error'],
                'dataMtime': extract_status['last_mtime'],
                'excelExists': os.path.exists(EXCEL_PATH),
                'excelName': os.path.basename(EXCEL_PATH),
                'watching': watcher_active,
                'tunnelUrl': tunnel_url
            })
        elif self.path == '/api/open-excel':
            if not os.path.exists(EXCEL_PATH):
                self.send_json({'ok': False, 'error': 'File not found: ' + EXCEL_PATH}, 404)
                return
            try:
                os.startfile(EXCEL_PATH)
                self.send_json({'ok': True, 'file': os.path.basename(EXCEL_PATH)})
            except Exception as e:
                self.send_json({'ok': False, 'error': str(e)}, 500)
        elif self.path == '/api/extract':
            if extract_status['running']:
                self.send_json({'ok': False, 'message': 'Extraction already running'})
            else:
                run_extraction_bg()
                self.send_json({'ok': True, 'message': 'Extraction started in background'})
        elif self.path == '/api/watch-start':
            self.start_watcher()
            self.send_json({'ok': True, 'watching': True})
        elif self.path == '/api/watch-stop':
            self.stop_watcher()
            self.send_json({'ok': True, 'watching': False})
        elif self.path == '/api/data-mtime':
            mtime = os.path.getmtime(DATA_BLOCK) if os.path.exists(DATA_BLOCK) else 0
            self.send_json({'mtime': mtime, 'extracting': extract_status['running']})
        elif self.path == '/api/tunnel-url':
            self.send_json({'url': tunnel_url})
        elif self.path == '/api/agent-status':
            has_key = os.path.exists(GROQ_KEY_FILE) and os.path.getsize(GROQ_KEY_FILE) > 0
            ollama_ok = False
            try:
                urllib.request.urlopen('http://localhost:11434/api/tags', timeout=2)
                ollama_ok = True
            except Exception:
                pass
            self.send_json({'hasKey': has_key, 'ollamaRunning': ollama_ok, 'dataLoaded': bool(_data_context)})
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == '/api/chat':
            self.handle_chat()
        elif self.path == '/api/set-key':
            self.handle_set_key()
        else:
            self.send_error(404)

    def handle_set_key(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length))
            key = body.get('key', '').strip()
            if not key:
                self.send_json({'error': 'No key provided'}, 400)
                return
            with open(GROQ_KEY_FILE, 'w') as f:
                f.write(key)
            self.send_json({'ok': True})
        except Exception as e:
            self.send_json({'error': str(e)}, 500)

    def handle_chat(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length))
            question = body.get('question', '')
            history = body.get('history', [])

            if not question:
                self.send_json({'error': 'No question provided'}, 400)
                return

            if not _data_context or _data_context == 'No data loaded yet.':
                build_data_context()

            relevant = get_relevant_data(question)
            system_msg = (
                'You are Halo, data analyst for HALO Supply Chain Dashboard '
                '(Himalaya Wellness, METAP Region). Answer ONLY from the data below. '
                'Be concise. Format numbers with commas.\n\n'
                'AVAILABLE DATASETS:\n' + _data_context + '\n\n'
                'RELEVANT DATA:\n' + relevant
            )

            messages = [{'role': 'system', 'content': system_msg}]
            for h in history[-10:]:
                messages.append({'role': h.get('role', 'user'), 'content': h.get('content', '')})
            messages.append({'role': 'user', 'content': question})

            groq_key = self._get_groq_key()
            if groq_key:
                answer = self._call_groq(messages, groq_key)
            else:
                answer = self._call_ollama(messages)
            self.send_json({'answer': answer})

        except Exception as e:
            self.send_json({'error': str(e)}, 500)

    def _get_groq_key(self):
        if os.path.exists(GROQ_KEY_FILE):
            with open(GROQ_KEY_FILE, 'r') as f:
                return f.read().strip()
        return ''

    def _call_groq(self, messages, api_key):
        payload = json.dumps({
            'model': GROQ_MODEL,
            'messages': messages,
            'temperature': 0.1,
            'max_tokens': 1024
        }).encode()
        req = urllib.request.Request(
            GROQ_URL, data=payload,
            headers={'Content-Type': 'application/json', 'Authorization': f'Bearer {api_key}', 'User-Agent': 'HALO-Dashboard/1.0'},
            method='POST'
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as res:
                result = json.loads(res.read())
                return result['choices'][0]['message']['content']
        except urllib.error.HTTPError as e:
            body = e.read().decode() if e.fp else ''
            if e.code in (401, 403):
                os.remove(GROQ_KEY_FILE) if os.path.exists(GROQ_KEY_FILE) else None
                raise Exception('INVALID_KEY: Your Groq API key is invalid. Please enter a new one.')
            raise Exception(f'Groq API error {e.code}: {body[:200]}')

    def _call_ollama(self, messages):
        payload = json.dumps({
            'model': OLLAMA_MODEL,
            'messages': messages,
            'stream': False,
            'options': {'temperature': 0.1, 'num_ctx': 6144}
        }).encode()
        req = urllib.request.Request(
            OLLAMA_URL, data=payload,
            headers={'Content-Type': 'application/json'}, method='POST'
        )
        with urllib.request.urlopen(req, timeout=300) as res:
            result = json.loads(res.read())
            return result.get('message', {}).get('content', 'No response.')

    def end_headers(self):
        if self.path and '_data_block' in self.path:
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        super().end_headers()

    def send_json(self, data, code=200):
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', len(body))
        self.send_header('Cache-Control', 'no-cache')
        self.end_headers()
        self.wfile.write(body)

    def start_watcher(self):
        global watcher_active, watcher_thread
        if watcher_active:
            return
        watcher_active = True
        watcher_thread = threading.Thread(target=watcher_loop, daemon=True)
        watcher_thread.start()

    def stop_watcher(self):
        global watcher_active
        watcher_active = False

    def log_message(self, fmt, *args):
        try:
            msg = str(args[0]) if args else ''
            if '/api/' in msg or 'error' in msg.lower():
                super().log_message(fmt, *args)
        except Exception:
            super().log_message(fmt, *args)


if __name__ == '__main__':
    os.system('title HALO Dashboard')

    if os.path.exists(DATA_BLOCK):
        extract_status['last_mtime'] = os.path.getmtime(DATA_BLOCK)

    build_data_context()

    server = ThreadingHTTPServer(('', PORT), DashboardHandler)
    print(f'  HALO Dashboard running at http://localhost:{PORT}')
    print(f'  Close this window to stop the server.')
    print()

    threading.Timer(1.5, lambda: webbrowser.open(f'http://localhost:{PORT}')).start()

    start_tunnel()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nShutting down.')
        server.shutdown()
