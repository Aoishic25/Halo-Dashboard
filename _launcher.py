"""
HALO Dashboard Launcher
Double-click to start the dashboard. Browser opens automatically.
Close the console window to stop the server.
"""
import os
import sys
import threading
import webbrowser

# When running as a PyInstaller EXE, the bundled files are in a temp dir.
# We need to ensure the real project root (where Excel/data live) is used,
# and that the bundled _server.py and _extract_openpyxl.py are importable.
if getattr(sys, 'frozen', False):
    BUNDLE_DIR = sys._MEIPASS
    PROJECT_ROOT = os.path.dirname(os.path.abspath(sys.argv[0]))
    sys.path.insert(0, BUNDLE_DIR)
else:
    PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

os.chdir(PROJECT_ROOT)

# Override ROOT in _server before it initializes, so it points to the
# project directory (where index.html, Ops files/, etc. live).
import _server
_server.ROOT = PROJECT_ROOT
_server.EXCEL_PATH = os.path.join(PROJECT_ROOT, 'Ops files', 'Monthly Review Meeting.xlsx')
_server.DATA_BLOCK = os.path.join(PROJECT_ROOT, '_data_block.js')
_server.GROQ_KEY_FILE = os.path.join(PROJECT_ROOT, '.groq_key')
_server.CLOUDFLARED = os.path.join(PROJECT_ROOT, 'cloudflared.exe')
EXCEL_DIR = os.path.join(PROJECT_ROOT, 'Ops files')

if __name__ == '__main__':
    try:
        os.system('title HALO Dashboard')
    except Exception:
        pass

    if os.path.exists(_server.DATA_BLOCK):
        _server.extract_status['last_mtime'] = os.path.getmtime(_server.DATA_BLOCK)

    _server.build_data_context()

    server = _server.ThreadingHTTPServer(('', _server.PORT), _server.DashboardHandler)
    print(f'  HALO Dashboard running at http://localhost:{_server.PORT}')
    print(f'  Close this window to stop the server.')
    print()

    threading.Timer(1.5, lambda: webbrowser.open(f'http://localhost:{_server.PORT}')).start()

    _server.start_tunnel()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nShutting down.')
        server.shutdown()
