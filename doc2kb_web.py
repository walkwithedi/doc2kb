#!/usr/bin/env python3
"""doc2kb_web.py - the browser interface of doc2kb

Python 3 standard library only. The desktop launcher starts this program: it listens on a random
port of 127.0.0.1 only, opens doc2kb.html from the same folder in the browser with a one-time
token in the address, and exits a few seconds after the last page is closed. While a run is in
progress, closing the pages or pressing "Quit" waits for the run to finish.

The conversion itself is doc2kb_core.py, run as a child process with the same options as on the
command line; this program only collects the choices, reads the core's progress lines and shows
the result.

Usage:
    python3 doc2kb_web.py [--no-browser] [--lang en|zh] [--log-level LEVEL]
"""

import argparse
import hmac
import json
import logging
import os
import re
import secrets
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlsplit

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import doc2kb_core as core  # noqa: E402  constants and labels only; conversions run in a child

LOG_FORMAT = core.LOG_FORMAT
log = logging.getLogger('doc2kb_web')

PAGE_FILE = os.path.join(HERE, 'doc2kb.html')
CORE = os.path.join(HERE, 'doc2kb_core.py')
VERSION_FILE = os.path.join(HERE, 'VERSION')
LOG_NAME = 'doc2kb_web.log'
PING_TIMEOUT = 600      # seconds without a heartbeat before a page counts as closed (fallback only)
OPEN_TIMEOUT = 120      # seconds to wait for the first page before exiting
BYE_GRACE = 5           # seconds after the last page closed; a reloaded page registers again meanwhile
CANCEL_WAIT = 10        # seconds the core gets to clean up after an interrupt
MAX_BODY = 4 * 1024 * 1024
MAX_LIST = 5000         # entries shown for one folder
OUTPUT_TAIL = 200       # lines of core output kept for the page
PROGRESS_RE = re.compile(r'^\[(\d+)/(\d+)\]\s*(.*)$')
DATE_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')

# messages from this program, (English, Chinese)
MESSAGES = {
    'denied': ('Access denied', '拒绝访问'),
    'not_found': ('No such address', '没有这个地址'),
    'page_missing': ('Page file doc2kb.html not found', '找不到页面文件 doc2kb.html'),
    'path_abs': ('%s must be an absolute path or start with ~/', '%s须为绝对路径或以 ~/ 开头'),
    'path_missing': ('%s not found: %s', '%s不存在：%s'),
    'not_dir': ('%s is not a folder: %s', '%s不是目录：%s'),
    'cannot_list': ('cannot open the folder: %s', '无法打开该目录：%s'),
    'src': ('The source folder', '源目录'),
    'folder': ('The folder', '目录'),
    'out': ('The output folder', '输出目录'),
    'out_not_writable': ('cannot write to the output folder: %s', '输出目录不可写：%s'),
    'out_is_src': ('the output folder cannot be the source folder itself', '输出目录不能就是源目录本身'),
    'bad_date': ('the date must be YYYY-MM-DD', '日期格式应为 YYYY-MM-DD'),
    'busy': ('another task is running; wait for it to finish or cancel it',
             '已有任务正在运行，请等它结束或先取消'),
    'no_files': ('no file ticked', '没有勾选任何文件'),
    'no_job': ('nothing is running', '没有正在运行的任务'),
    'core_missing': ('doc2kb_core.py not found next to doc2kb_web.py',
                     '找不到同目录的 doc2kb_core.py'),
    'no_xdg_open': ('xdg-open not found; open the folder yourself: %s',
                    '未找到 xdg-open，请手动打开：%s'),
    'cannot_run': ('cannot start doc2kb_core.py: %s', '无法启动 doc2kb_core.py：%s'),
    'tell_page_missing': ('Page file %s not found. Unpack the whole doc2kb folder again.',
                          '找不到页面文件 %s，请重新解压完整的 doc2kb 目录。'),
    'tell_no_browser': ('Could not open a browser. Open this address in one: %s',
                        '未能自动打开浏览器，请在浏览器中打开：%s'),
    'serving': ('Browser interface: %s', '网页界面：%s'),
}

LANG = 'en'


def M(key, *args):
    text = MESSAGES[key][1 if LANG == 'zh' else 0]
    return text % args if args else text


def detect_lang():
    """The same rule as doc2kb.sh: DOC2KB_LANG, else a locale starting with zh, else English."""
    lang = os.environ.get('DOC2KB_LANG', '').lower()
    if lang in core.LANGS:
        return lang
    loc = os.environ.get('LC_ALL') or os.environ.get('LC_MESSAGES') or os.environ.get('LANG') or ''
    return 'zh' if loc.startswith('zh') else 'en'


def read_version():
    try:
        with open(VERSION_FILE, encoding='utf-8') as f:
            return f.read().strip()
    except OSError:
        return ''


# ---------------------------------------------------------------- paths

def norm_path(value, label):
    """An absolute path or one starting with ~; ValueError otherwise."""
    value = (value or '').strip()
    if value.startswith('~'):
        value = os.path.expanduser(value)
    if not value or not os.path.isabs(value):
        raise ValueError(M('path_abs', label))
    return os.path.normpath(value)


def existing_dir(value, label):
    path = norm_path(value, label)
    if not os.path.exists(path):
        raise ValueError(M('path_missing', label, path))
    if not os.path.isdir(path):
        raise ValueError(M('not_dir', label, path))
    return path


def file_kind(name):
    """'doc', 'archive', 'legacy' or None: what the converter makes of a file name."""
    kind = core.archive_kind(name)
    if kind in ('zip', 'tar'):
        return 'archive'
    ext = os.path.splitext(name)[1].lower()
    if ext in core.SUPPORTED_EXT:
        return 'doc'
    if ext in core.LEGACY_EXT:
        return 'legacy'
    return None


def roots():
    """Starting points for the folder browser: home, file system root and mount folders."""
    home = os.path.expanduser('~')
    out = [home, '/']
    user = os.path.basename(home)
    for d in ('/media/' + user, '/run/media/' + user, '/media', '/mnt'):
        if os.path.isdir(d) and d not in out:
            out.append(d)
    return out


def list_dir(path):
    """Sub-folders and the files the converter reads; hidden entries are left out."""
    try:
        names = sorted(os.listdir(path), key=lambda n: n.lower())
    except OSError as e:
        raise ValueError(M('cannot_list', e.strerror or e))
    dirs, files = [], []
    for name in names:
        if name.startswith('.'):
            continue
        full = os.path.join(path, name)
        if os.path.isdir(full):
            dirs.append(core.fs_text(name))
        else:
            kind = file_kind(name)
            if kind and not name.startswith('~$'):
                try:
                    size = os.path.getsize(full)
                except OSError:
                    continue
                files.append({'name': core.fs_text(name), 'kind': kind,
                              'size': core.human_size(size)})
    parent = os.path.dirname(path)
    return {'path': path, 'parent': parent if parent != path else None,
            'dirs': dirs[:MAX_LIST], 'files': files[:MAX_LIST],
            'truncated': len(dirs) > MAX_LIST or len(files) > MAX_LIST}


def default_out(src):
    return os.path.join(src, core.T('out_dir_suffix'))


def writable_target(path):
    """The folder can be written, or created inside the nearest folder that exists."""
    p = path
    while not os.path.exists(p):
        parent = os.path.dirname(p)
        if parent == p:
            return False
        p = parent
    return os.path.isdir(p) and os.access(p, os.W_OK | os.X_OK)


def has_tool(name):
    return shutil.which(name) is not None


# ---------------------------------------------------------------- tasks

class Job:
    """The current or latest scan or conversion. One at a time."""

    def __init__(self):
        self.lock = threading.Lock()
        self.reset(None, None)

    def reset(self, kind, params):
        self.kind = kind
        self.params = params
        self.running = False
        self.cancelled = False
        self.proc = None
        self.rc = None
        self.started = None
        self.started_wall = None
        self.finished = None
        self.lines = []
        self.cur = 0
        self.total = 0
        self.current = ''
        self.result = None
        self.tmpdir = None

    def start(self, kind, params, argv, tmpdir=None):
        """tmpdir: a private folder of this task, removed when it ends."""
        with self.lock:
            if self.running:
                return False
            self.reset(kind, params)
            self.tmpdir = tmpdir
            self.running = True
            self.started = time.strftime('%Y-%m-%d %H:%M:%S')
            self.started_wall = time.time()
        threading.Thread(target=self._run, args=(argv,), daemon=True).start()
        return True

    def _run(self, argv):
        log.info('%s start: %s', self.kind, argv)
        env = dict(os.environ, PYTHONIOENCODING='utf-8', DOC2KB_LANG=LANG)
        rc = None
        try:
            proc = subprocess.Popen(argv, env=env, stdin=subprocess.DEVNULL,
                                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                    universal_newlines=True, encoding='utf-8', errors='replace')
            with self.lock:
                self.proc = proc
                cancel_now = self.cancelled
            if cancel_now:
                self._interrupt(proc)
            with proc.stdout:
                for line in proc.stdout:
                    self._line(line.rstrip('\n'))
            rc = proc.wait()
        except OSError as e:
            log.error('cannot run the core: %s', e)
            self._line(M('cannot_run', e))
            rc = -1
        with self.lock:
            self.rc = rc
            self.proc = None
            self.finished = time.strftime('%Y-%m-%d %H:%M:%S')
            try:
                self.result = self._result()
            except Exception:   # a broken result must not leave the job running for ever
                log.exception('cannot build the result')
                self.result = {}
            self.running = False
            if self.tmpdir:
                shutil.rmtree(self.tmpdir, ignore_errors=True)
                self.tmpdir = None
        log.info('%s finished rc=%s cancelled=%s', self.kind, rc, self.cancelled)

    def _line(self, line):
        with self.lock:
            self.lines.append(line)
            m = PROGRESS_RE.match(line)
            if m:
                self.cur, self.total, self.current = int(m.group(1)), int(m.group(2)), m.group(3)

    def cancel(self):
        with self.lock:
            if not self.running:
                return False
            self.cancelled = True
            proc = self.proc
        if proc:
            threading.Thread(target=self._interrupt, args=(proc,), daemon=True).start()
        return True

    @staticmethod
    def _interrupt(proc):
        """SIGINT first, so the core removes its private temporary folder; then stronger."""
        for sig, wait in ((signal.SIGINT, CANCEL_WAIT), (signal.SIGTERM, 5), (signal.SIGKILL, 5)):
            try:
                proc.send_signal(sig)
            except OSError:
                return
            try:
                proc.wait(wait)
                return
            except subprocess.TimeoutExpired:
                continue

    # -- results, built once the child has ended (lock held)

    def _result(self):
        if self.kind == 'scan':
            return self._scan_result()
        return self._run_result()

    def _scan_result(self):
        ok, legacy, filtered = [], [], 0
        if self.rc != 0 or self.cancelled:
            return {}
        for line in self.lines:
            parts = line.split('\t')
            if parts[0] == 'OK' and len(parts) >= 3:
                ok.append({'rel': parts[1], 'size': parts[2]})
            elif parts[0] == 'LEGACY' and len(parts) >= 2:
                legacy.append(parts[1])
            elif parts[0] == 'FILTERED':
                filtered += 1
        src = self.params['src']
        return {'files': ok, 'legacy': legacy, 'filtered': filtered,
                'out_default': default_out(src)}

    def _run_result(self):
        p = self.params
        out = p['out']
        stats, failed, summary = {}, [], []
        for line in self.lines:
            parts = line.split('\t')
            if parts[0] == 'STAT' and len(parts) == 3:
                stats[parts[1]] = parts[2]
            elif parts[0] == 'FAILED' and len(parts) >= 2:
                failed.append(parts[1])
        # the human summary: the lines after the last blank line, without STAT and FAILED
        blank = max((i for i, l in enumerate(self.lines) if not l.strip()), default=-1)
        if self.rc == 0:
            summary = [l for l in self.lines[blank + 1:]
                       if not l.startswith(('STAT\t', 'FAILED\t'))]
        report_path = os.path.join(out, core.T('report_file'))
        report = fresh_report(report_path, self.started_wall)
        written = written_md(out, self.started_wall) if not p['merge'] else 0
        return {'stats': stats, 'failed': failed, 'summary': summary,
                'report': report, 'report_path': report_path if report is not None else None,
                'written_md': written, 'skipped_pdf': p.get('skipped_pdf', [])}

    def snapshot(self, tail=OUTPUT_TAIL):
        with self.lock:
            if self.kind is None:
                return {'state': 'idle'}
            state = 'running' if self.running else ('cancelled' if self.cancelled else 'done')
            shown = [l for l in self.lines if not l.startswith(('OK\t', 'LEGACY\t', 'FILTERED\t'))]
            return {'kind': self.kind, 'state': state, 'params': self.params,
                    'rc': self.rc, 'started': self.started, 'finished': self.finished,
                    'cur': self.cur, 'total': self.total, 'current': self.current,
                    'output': '\n'.join(shown[-tail:]),
                    'result': self.result if not self.running else None}


def fresh_report(path, since_wall):
    """The report text up to the per-document logs, only when this run wrote it."""
    try:
        if os.path.getmtime(path) < since_wall - 1:
            return None
        with open(path, encoding='utf-8', errors='replace') as f:
            text = f.read()
    except OSError:
        return None
    head = core.T('rep_logs')
    lines = text.splitlines()
    if head in lines:
        lines = lines[:lines.index(head)]
    return '\n'.join(lines).rstrip() + '\n'


def written_md(out, since_wall):
    """Per-document md files written in the output folder since the run started."""
    n = 0
    try:
        names = os.listdir(out)
    except OSError:
        return 0
    for name in names:
        if name.endswith('.md') and not name.startswith('_'):
            try:
                if os.path.getmtime(os.path.join(out, name)) >= since_wall - 1:
                    n += 1
            except OSError:
                pass
    return n


def scan_argv(src, since):
    argv = [sys.executable, '-u', CORE, '--batch', src, '--scan-only', '--lang', LANG]
    if since:
        argv += ['--since', since]
    return argv


def run_plan(data):
    """Check a run request; returns (params, argv, only-list lines) or raises ValueError."""
    src = existing_dir(data.get('src'), M('src'))
    out = norm_path(data.get('out') or default_out(src), M('out'))
    if out == src:
        raise ValueError(M('out_is_src'))
    if not writable_target(out):
        raise ValueError(M('out_not_writable', out))
    since = (data.get('since') or '').strip()
    if since and not DATE_RE.match(since):
        raise ValueError(M('bad_date'))
    only = [r for r in (data.get('only') or []) if isinstance(r, str) and r.strip()]
    skipped_pdf = []
    if not has_tool('pdftotext'):
        skipped_pdf = [r for r in only if r.lower().endswith('.pdf')]
        only = [r for r in only if not r.lower().endswith('.pdf')]
    if not only:
        raise ValueError(M('no_files'))
    merge = bool(data.get('merge'))
    images = bool(data.get('images'))
    sheets = bool(data.get('sheets')) and images and has_tool('montage')
    params = {'src': src, 'out': out, 'since': since, 'merge': merge,
              'inline': bool(data.get('inline')), 'images': images,
              'keep_dups': bool(data.get('keep_dups')), 'sheets': sheets,
              'force': bool(data.get('force')) and not merge,
              'count': len(only), 'skipped_pdf': skipped_pdf}
    argv = [sys.executable, '-u', CORE, '--batch', src, '-o', out, '--lang', LANG]
    if since:
        argv += ['--since', since]
    if merge:
        argv.append('--merge')
    if not params['inline']:
        argv.append('--no-inline')
    if not images:
        argv.append('--no-images')
    if params['keep_dups']:
        argv.append('--no-dedup')
    if sheets:
        argv.append('--sheets')
    if params['force']:
        argv.append('--force')
    return params, argv, only


def run_check(params):
    """What a run would replace, for the confirmation on the page."""
    out = params['out']
    replace, existing = [], 0
    if params['merge']:
        name = core.safe_name(os.path.basename(params['src'])) + '.md'
        if os.path.isfile(os.path.join(out, name)):
            replace.append(name)
    elif params['force']:
        try:
            existing = sum(1 for n in os.listdir(out) if n.endswith('.md') and not n.startswith('_'))
        except OSError:
            existing = 0
    return {'replace': replace, 'existing_md': existing, 'out': out,
            'count': params['count'], 'skipped_pdf': params['skipped_pdf']}


# ---------------------------------------------------------------- web service

def log_file_path():
    """The log goes next to the program, or to ~/.cache/doc2kb/ when that folder is read-only."""
    if os.access(HERE, os.W_OK):
        return os.path.join(HERE, LOG_NAME)
    d = os.path.join(os.path.expanduser('~'), '.cache', 'doc2kb')
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, LOG_NAME)


class WebState:
    def __init__(self, port, token):
        self.port = port
        self.token = token
        self.origin = 'http://127.0.0.1:%d' % port
        self.started = time.monotonic()
        self.lock = threading.Lock()
        self.pages = {}         # page id -> time of its last heartbeat
        self.seen = False
        self.empty_since = None
        self.quit = threading.Event()
        self.job = Job()

    def ping(self, page):
        with self.lock:
            self.pages[page] = time.monotonic()
            self.seen = True
            self.empty_since = None

    def bye(self, page):
        with self.lock:
            self.pages.pop(page, None)
            if self.seen and not self.pages and self.empty_since is None:
                self.empty_since = time.monotonic()

    def should_exit(self):
        """The reason to exit, or None; never while a task runs."""
        now = time.monotonic()
        with self.lock:
            for page, t in list(self.pages.items()):
                if now - t > PING_TIMEOUT:
                    log.info('page %s heartbeat lost', page)
                    del self.pages[page]
                    if not self.pages:
                        self.empty_since = now - BYE_GRACE
            if self.job.running:
                return None
            if self.quit.is_set():
                return 'quit requested'
            if not self.seen:
                return 'page never opened' if now - self.started > OPEN_TIMEOUT else None
            if not self.pages and self.empty_since is not None and now - self.empty_since >= BYE_GRACE:
                return 'all pages closed'
        return None


def tell_user(msg):
    """Started from the launcher there is no terminal: show the message with zenity or notify-send."""
    log.warning('%s', msg)
    print(msg, flush=True)
    if sys.stdout.isatty():
        return
    for cmd in (['zenity', '--info', '--no-wrap', '--no-markup', '--text=' + msg], ['notify-send', msg]):
        if not has_tool(cmd[0]):
            continue
        try:
            subprocess.Popen(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL)
            return
        except OSError as e:
            log.debug('%s failed: %s', cmd[0], e)


class Handler(BaseHTTPRequestHandler):
    state = None

    def log_message(self, fmt, *args):
        log.debug('%s - %s', self.address_string(), fmt % args)

    def send_json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8', 'replace')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def allowed(self, api, token=None):
        """Host and Origin must be this server; API calls also need the token (X-Token header).
        Answers 403 itself when the check fails."""
        st = self.state
        ok = self.headers.get('Host') == '127.0.0.1:%d' % st.port
        origin = self.headers.get('Origin')
        if ok and origin is not None:
            ok = origin == st.origin
        if ok and api:
            if token is None:
                token = self.headers.get('X-Token', '')
            ok = isinstance(token, str) and hmac.compare_digest(token, st.token)
        if not ok:
            log.warning('rejected %s %s', self.command, urlsplit(self.path).path)
            self.send_json(403, {'error': M('denied')})
        return ok

    def read_json(self):
        try:
            n = int(self.headers.get('Content-Length') or 0)
        except ValueError:
            n = 0
        if n <= 0 or n > MAX_BODY:
            return {}
        try:
            data = json.loads(self.rfile.read(n).decode('utf-8'))
        except (ValueError, UnicodeDecodeError):
            return {}
        return data if isinstance(data, dict) else {}

    def do_GET(self):
        url = urlsplit(self.path)
        api = url.path.startswith('/api/')
        if not self.allowed(api):
            return
        q = parse_qs(url.query)
        if url.path in ('/', '/index.html'):
            try:
                with open(PAGE_FILE, 'rb') as f:
                    body = f.read()
            except OSError as e:
                log.error('page missing: %s', e)
                self.send_json(500, {'error': M('page_missing')})
                return
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('Content-Security-Policy',
                             "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; "
                             "connect-src 'self'; img-src 'self' data:")
            self.end_headers()
            self.wfile.write(body)
        elif url.path == '/api/status':
            self.send_json(200, {
                'lang': LANG, 'version': read_version(), 'home': os.path.expanduser('~'),
                'roots': roots(), 'core': os.path.isfile(CORE),
                'tools': {'pdftotext': has_tool('pdftotext'), 'montage': has_tool('montage'),
                          'xdg-open': has_tool('xdg-open')},
                'job': self.state.job.snapshot(),
            })
        elif url.path == '/api/ls':
            try:
                # a folder still to be created (a new output folder): show the nearest one that exists
                path = norm_path((q.get('path') or [''])[0] or '~', M('folder'))
                while not os.path.isdir(path) and os.path.dirname(path) != path:
                    path = os.path.dirname(path)
                self.send_json(200, list_dir(path))
            except ValueError as e:
                self.send_json(400, {'error': str(e)})
        elif url.path == '/api/job':
            self.send_json(200, self.state.job.snapshot())
        else:
            self.send_json(404, {'error': M('not_found')})

    def do_POST(self):
        url = urlsplit(self.path)
        if url.path == '/api/bye':
            # sent by sendBeacon, which cannot set headers: the token is in the body
            data = self.read_json()
            if not self.allowed(True, data.get('t')):
                return
            log.info('page %s closed', data.get('page'))
            self.state.bye(str(data.get('page')))
            self.send_json(200, {'ok': True})
            return
        if not self.allowed(True):
            return
        data = self.read_json()
        job = self.state.job
        if url.path == '/api/ping':
            self.state.ping(str(data.get('page')))
            self.send_json(200, {'ok': True})
        elif url.path == '/api/quit':
            running = job.running
            log.info('quit requested by a page%s', ', waiting for the task' if running else '')
            self.send_json(200, {'ok': True, 'waiting': running})
            self.state.quit.set()
        elif url.path == '/api/scan':
            try:
                src = existing_dir(data.get('src'), M('src'))
                since = (data.get('since') or '').strip()
                if since and not DATE_RE.match(since):
                    raise ValueError(M('bad_date'))
            except ValueError as e:
                self.send_json(400, {'error': str(e)})
                return
            if not os.path.isfile(CORE):
                self.send_json(500, {'error': M('core_missing')})
                return
            if not job.start('scan', {'src': src, 'since': since}, scan_argv(src, since)):
                self.send_json(409, {'error': M('busy')})
                return
            self.send_json(200, {'ok': True})
        elif url.path in ('/api/check', '/api/run'):
            try:
                params, argv, only = run_plan(data)
            except ValueError as e:
                self.send_json(400, {'error': str(e)})
                return
            if url.path == '/api/check':
                self.send_json(200, run_check(params))
                return
            if not os.path.isfile(CORE):
                self.send_json(500, {'error': M('core_missing')})
                return
            if job.running:
                self.send_json(409, {'error': M('busy')})
                return
            tmp = tempfile.mkdtemp(prefix='doc2kb_web_')
            only_file = os.path.join(tmp, 'only.txt')
            with open(only_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(only) + '\n')
            argv += ['--only', only_file]
            if not job.start('run', params, argv, tmpdir=tmp):
                shutil.rmtree(tmp, ignore_errors=True)
                self.send_json(409, {'error': M('busy')})
                return
            self.send_json(200, {'ok': True})
        elif url.path == '/api/cancel':
            if not job.cancel():
                self.send_json(409, {'error': M('no_job')})
                return
            self.send_json(200, {'ok': True})
        elif url.path == '/api/open':
            # only the output folder of the latest run
            with job.lock:
                out = job.params.get('out') if job.kind == 'run' and job.params else None
            if not out or not os.path.isdir(out):
                self.send_json(400, {'error': M('path_missing', M('out'), out or '')})
                return
            if not has_tool('xdg-open'):
                self.send_json(500, {'error': M('no_xdg_open', out)})
                return
            subprocess.Popen(['xdg-open', out], stdin=subprocess.DEVNULL,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            self.send_json(200, {'ok': True})
        else:
            self.send_json(404, {'error': M('not_found')})


def watchdog(server, st, interval=1):
    """Stops the server on "Quit", when every page is closed, or when no page ever opened;
    a running task is waited for."""
    while True:
        st.quit.wait(interval)
        reason = st.should_exit()
        if reason:
            log.info('%s, exiting', reason)
            break
        if st.quit.is_set():
            time.sleep(interval)    # quit asked while a task runs: quit.wait no longer blocks
    server.shutdown()


def serve(open_browser=True):
    if not os.path.isfile(PAGE_FILE):
        tell_user(M('tell_page_missing', PAGE_FILE))
        return 1
    try:
        fh = logging.FileHandler(log_file_path(), encoding='utf-8')
        fh.setFormatter(logging.Formatter(LOG_FORMAT))
        logging.getLogger().addHandler(fh)
    except OSError as e:
        log.warning('no log file: %s', e)
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    st = WebState(server.server_address[1], secrets.token_urlsafe(24))
    Handler.state = st
    url = '%s/?t=%s' % (st.origin, st.token)
    log.info('serving on %s', st.origin)
    print(M('serving', url), flush=True)
    threading.Thread(target=watchdog, args=(server, st), daemon=True).start()
    if open_browser and not webbrowser.open(url):
        tell_user(M('tell_no_browser', url))
    try:
        server.serve_forever(poll_interval=0.5)
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        st.job.cancel()
    log.info('server stopped')
    return 0


def main(argv=None):
    global LANG
    ap = argparse.ArgumentParser(description='Browser interface of doc2kb (what the desktop launcher starts).')
    ap.add_argument('--no-browser', action='store_true', help='do not open a browser, only print the address')
    ap.add_argument('--lang', choices=core.LANGS, default=None,
                    help='interface language (default: DOC2KB_LANG, else zh for a zh locale, else en)')
    ap.add_argument('--log-level', default=os.environ.get('LOG_LEVEL', 'WARNING').upper(),
                    type=str.upper, choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                    help='default: environment variable LOG_LEVEL, else WARNING')
    args = ap.parse_args(argv)
    logging.basicConfig(level=args.log_level, format=LOG_FORMAT, stream=sys.stderr)
    LANG = args.lang or detect_lang()
    core.set_lang(LANG)
    return serve(open_browser=not args.no_browser)


if __name__ == '__main__':
    sys.exit(main())
