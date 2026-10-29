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
import doc2kb_dedup as dedup  # noqa: E402  labels only; the clean-up runs in a child

LOG_FORMAT = core.LOG_FORMAT
log = logging.getLogger('doc2kb_web')

PAGE_FILE = os.path.join(HERE, 'doc2kb.html')
CORE = os.path.join(HERE, 'doc2kb_core.py')
SHELL = os.path.join(HERE, 'doc2kb.sh')
DEDUP = os.path.join(HERE, 'doc2kb_dedup.py')
VERSION_FILE = os.path.join(HERE, 'VERSION')
LOG_NAME = 'doc2kb_web.log'
PING_TIMEOUT = 600      # seconds without a heartbeat before a page counts as closed (fallback only)
OPEN_TIMEOUT = 120      # seconds to wait for the first page before exiting
BYE_GRACE = 5           # seconds after the last page closed; a reloaded page registers again meanwhile
CANCEL_WAIT = 10        # seconds the core gets to clean up after an interrupt
MAX_BODY = 32 * 1024 * 1024  # a ticked list of tens of thousands of long paths fits
MAX_LIST = 5000         # entries shown for one folder
OUTPUT_TAIL = 200       # lines of core output kept for the page
PROGRESS_RE = re.compile(r'^\[(\d+)/(\d+)\]\s*(.*)$')
BATCH_RE = re.compile(r'^BATCH\t(\d+)\t(\d+)$')
MACHINE_LINES = ('OK\t', 'LEGACY\t', 'FILTERED\t', 'BATCH\t')   # not shown on the page
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
    'file': ('A file', '文件'),
    'task_name': ('the task needs a name', '任务需要一个名字'),
    'task_exists': ('a task named %s exists already', '已有名为「%s」的任务'),
    'task_missing': ('no task named %s', '没有名为「%s」的任务'),
    'no_task': ('no task chosen', '没有选中任务'),
    'not_md': ('choose a merged or incremental md file: %s', '请选择合并或增量 md 文件：%s'),
    'bad_days': ('the number of days must be a whole number from 1 to 36500',
                 '天数须为 1 到 36500 的整数'),
    'file_missing': ('file not found: %s', '文件不存在：%s'),
    'file_unsupported': ('not a document doc2kb reads: %s', '不是 doc2kb 能处理的文档：%s'),
    'out_not_writable': ('cannot write to the output folder: %s', '输出目录不可写：%s'),
    'out_is_src': ('the output folder cannot be the source folder itself', '输出目录不能就是源目录本身'),
    'bad_date': ('the date must be YYYY-MM-DD', '日期格式应为 YYYY-MM-DD'),
    'busy': ('another task is running; wait for it to finish or cancel it',
             '已有任务正在运行，请等它结束或先取消'),
    'no_files': ('no file ticked', '没有勾选任何文件'),
    'bad_batch_size': ('files per batch must be a whole number from 0 to 1000000',
                       '每批文档数须为 0 到 1000000 的整数'),
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


def log_safe(value):
    """A request value for a log line: line breaks escaped, so it cannot start a line of its own."""
    return str(value).replace('\r', '\\r').replace('\n', '\\n')


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
    """'doc', 'archive', 'legacy', 'md' (output, for the clean-up) or None."""
    if name.lower().endswith('.md'):
        return 'md'
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
    """The current or latest task: a scan, a conversion, a series of incremental tasks or a
    clean-up. One at a time. A task is one or more steps, each a child process run in turn."""

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
        self.batch = 0
        self.batches = 0
        self.task = 0           # tasks run: the task running now, of how many, its name
        self.tasks = 0
        self.task_name = ''
        self.result = None
        self.tmpdir = None
        self.steps = []

    def start(self, kind, params, argv, tmpdir=None, steps=None):
        """argv: the command of a one-step task. steps: instead of argv, a list of dicts with
        'name' and 'argv' (None when the step cannot run; 'error' says why); each gets 'rc',
        'first' (its first output line) and 'last'. tmpdir: a private folder of this task,
        removed when it ends."""
        with self.lock:
            if self.running:
                return False
            self.reset(kind, params)
            self.tmpdir = tmpdir
            self.running = True
            self.started = time.strftime('%Y-%m-%d %H:%M:%S')
            self.started_wall = time.time()
            self.steps = steps if steps is not None else [{'name': '', 'argv': argv}]
        threading.Thread(target=self._run, daemon=True).start()
        return True

    def _run(self):
        rc = None
        for i, step in enumerate(self.steps, 1):
            with self.lock:
                if self.cancelled:
                    break
                step['first'] = len(self.lines)
                # document progress ([i/n], BATCH lines) is counted per step
                self.cur = self.total = self.batch = self.batches = 0
                self.current = ''
                self.task, self.tasks, self.task_name = i, len(self.steps), step['name']
            if step.get('argv'):
                rc = self._run_step(step['argv'])
            else:
                self._line(step.get('error', ''))
                rc = None
            with self.lock:
                step['rc'] = rc
                step['last'] = len(self.lines)
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

    def _run_step(self, argv):
        log.info('%s step: %s', self.kind, log_safe(argv))
        # DOC2KB_PROGRESS: doc2kb.sh passes the converter's progress lines on as they come
        env = dict(os.environ, PYTHONIOENCODING='utf-8', DOC2KB_LANG=LANG, DOC2KB_PROGRESS='1')
        try:
            # a session of its own, so a cancel reaches doc2kb.sh and the converter it starts
            proc = subprocess.Popen(argv, env=env, stdin=subprocess.DEVNULL,
                                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                    universal_newlines=True, encoding='utf-8', errors='replace',
                                    start_new_session=True)
        except OSError as e:
            log.error('cannot run %s: %s', argv[0], e)
            self._line(M('cannot_run', e))
            return -1
        with self.lock:
            self.proc = proc
            cancel_now = self.cancelled
        if cancel_now:
            self._interrupt(proc)
        with proc.stdout:
            for line in proc.stdout:
                self._line(line.rstrip('\n'))
        rc = proc.wait()
        with self.lock:
            self.proc = None
        return rc

    def _line(self, line):
        with self.lock:
            self.lines.append(line)
            m = PROGRESS_RE.match(line)
            if m:
                self.cur, self.total, self.current = int(m.group(1)), int(m.group(2)), m.group(3)
            m = BATCH_RE.match(line)
            if m:
                self.batch, self.batches = int(m.group(1)), int(m.group(2))

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
        """SIGINT first, so the core removes its private temporary folder; then stronger.
        The signal goes to the step's whole process group."""
        for sig, wait in ((signal.SIGINT, CANCEL_WAIT), (signal.SIGTERM, 5), (signal.SIGKILL, 5)):
            try:
                os.killpg(proc.pid, sig)
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
        if self.kind == 'tasks':
            return self._tasks_result()
        if self.kind == 'dedup':
            return self._dedup_result()
        return self._run_result()

    def _dedup_result(self):
        found = {}
        for line in self.lines:
            for key, rx in DEDUP_LINES.items():
                m = rx.match(line)
                if m and key not in found:
                    # one language's groups match, the other's are None
                    found[key] = tuple(g for g in m.groups() if g is not None)
        sections, duplicate, kept = (int(x) for x in found.get('summary', (0, 0, 0)))
        report_path = found.get('report', (None,))[0]
        report = None
        if report_path and self.rc == 0:
            try:
                if os.path.getmtime(report_path) >= self.started_wall - 1:
                    with open(report_path, encoding='utf-8', errors='replace') as f:
                        report = f.read()
            except OSError:
                report = None
        return {'sections': sections, 'duplicate': duplicate, 'kept': kept,
                'report': report, 'report_path': report_path,
                'backup': found.get('backup', (None,))[0], 'dry_run': self.params['dry_run']}

    def _tasks_result(self):
        out = []
        for step in self.steps:
            if 'rc' not in step:
                continue    # not reached: cancelled before
            lines = [l for l in self.lines[step['first']:step['last']] if l.strip()]
            out.append({'name': step['name'], 'rc': step['rc'], 'first_run': step['first_run'],
                        'dry_run': step['dry_run'], 'log': step['log'],
                        'last_line': lines[-1] if lines else '',
                        'output': '\n'.join(lines[-OUTPUT_TAIL:])})
        return {'tasks': out}

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
        try:
            failed_n = int(stats.get('failed', len(failed)))
        except ValueError:
            failed_n = len(failed)
        return {'stats': stats, 'failed': failed, 'failed_n': failed_n, 'summary': summary,
                'report': report, 'report_path': report_path if report is not None else None,
                'written_md': written, 'skipped_pdf': p.get('skipped_pdf', []),
                'skipped_legacy': p.get('skipped_legacy', [])}

    def snapshot(self, tail=OUTPUT_TAIL):
        with self.lock:
            if self.kind is None:
                return {'state': 'idle'}
            state = 'running' if self.running else ('cancelled' if self.cancelled else 'done')
            shown = [l for l in self.lines if not l.startswith(MACHINE_LINES)]
            return {'kind': self.kind, 'state': state, 'params': self.params,
                    'rc': self.rc, 'started': self.started, 'finished': self.finished,
                    'cur': self.cur, 'total': self.total, 'current': self.current,
                    'batch': self.batch, 'batches': self.batches,
                    'task': self.task, 'tasks': self.tasks, 'task_name': self.task_name,
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
                pass  # removed meanwhile: not counted
    return n


def scan_argv(src, since):
    argv = [sys.executable, '-u', CORE, '--batch', src, '--scan-only', '--lang', LANG]
    if since:
        argv += ['--since', since]
    return argv


def common_options(data, out, merge):
    """The options both run kinds share: (params, core arguments)."""
    images = bool(data.get('images'))
    sheets = bool(data.get('sheets')) and images and has_tool('montage')
    params = {'out': out, 'merge': merge, 'inline': bool(data.get('inline')), 'images': images,
              'keep_dups': bool(data.get('keep_dups')), 'sheets': sheets}
    argv = ['-o', out, '--lang', LANG]
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
    return params, argv


def checked_out(data, default):
    out = norm_path(data.get('out') or default, M('out'))
    if not writable_target(out):
        raise ValueError(M('out_not_writable', out))
    return out


def split_pdf(paths):
    """(kept, PDF files left out because pdftotext is missing)."""
    if has_tool('pdftotext'):
        return paths, []
    return ([p for p in paths if not p.lower().endswith('.pdf')],
            [p for p in paths if p.lower().endswith('.pdf')])


def run_plan(data):
    """Check a run request. Returns (params, argv, list lines, list option) or raises ValueError;
    the list lines go to a file passed with the list option."""
    if data.get('mode') == 'files':
        return files_plan(data)
    src = existing_dir(data.get('src'), M('src'))
    out = checked_out(data, default_out(src))
    if out == src:
        raise ValueError(M('out_is_src'))
    since = (data.get('since') or '').strip()
    if since and not DATE_RE.match(since):
        raise ValueError(M('bad_date'))
    size = data.get('batch_size', core.DEFAULT_BATCH_SIZE)
    if isinstance(size, bool) or not isinstance(size, int) or not 0 <= size <= 1000000:
        raise ValueError(M('bad_batch_size'))
    # every scanned file ticked: no list, which can be large; the core finds the same files
    if data.get('all'):
        only, skipped_pdf, count = None, [], data.get('count') if isinstance(data.get('count'), int) else 0
    else:
        only = [r for r in (data.get('only') or []) if isinstance(r, str) and r.strip()]
        only, skipped_pdf = split_pdf(only)
        if not only:
            raise ValueError(M('no_files'))
        count = len(only)
    merge = bool(data.get('merge'))
    params, opts = common_options(data, out, merge)
    # force: one file per document, convert again; merged, drop an unfinished run's batches
    params.update({'kind': 'run', 'src': src, 'since': since, 'force': bool(data.get('force')),
                   'batch_size': size, 'all': only is None,
                   'count': count, 'skipped_pdf': skipped_pdf, 'skipped_legacy': []})
    argv = [sys.executable, '-u', CORE, '--batch', src, '--batch-size', str(size)] + opts
    if since:
        argv += ['--since', since]
    if params['force']:
        argv.append('--force')
    return params, argv, only, '--only'


def files_plan(data):
    """Chosen files: absolute paths of documents or archives; legacy formats are left out."""
    files, legacy = [], []
    for f in data.get('files') or []:
        path = norm_path(f if isinstance(f, str) else '', M('file'))
        if not os.path.isfile(path):
            raise ValueError(M('file_missing', path))
        kind = file_kind(os.path.basename(path))
        if kind == 'legacy':
            legacy.append(path)
        elif kind in ('doc', 'archive'):
            files.append(path)
        else:
            raise ValueError(M('file_unsupported', path))
    files, skipped_pdf = split_pdf(files)
    if not files:
        raise ValueError(M('no_files'))
    src = os.path.dirname(files[0])
    merge = bool(data.get('merge'))
    out = checked_out(data, default_out(src))
    params, opts = common_options(data, out, merge)
    params.update({'kind': 'files', 'src': src, 'since': '', 'force': False, 'files': files,
                   'count': len(files), 'skipped_pdf': skipped_pdf, 'skipped_legacy': legacy})
    return params, [sys.executable, '-u', CORE] + opts, files, '--files'


def run_check(params):
    """What a run would replace, for the confirmation on the page."""
    out = params['out']
    replace, existing = [], 0
    unfinished = False
    if params['merge']:
        stem = core.safe_name(os.path.basename(params['src']))
        size = params.get('batch_size') or 0
        n = -(-params['count'] // size) if size and params['count'] > size else 1
        names = core.volume_names(stem, n)
        if n > 1:
            names.append(stem + core.T('vol_index_suffix'))
        replace = [x + '.md' for x in names if os.path.isfile(os.path.join(out, x + '.md'))]
        unfinished = params['kind'] == 'run' and os.path.isfile(
            os.path.join(out, core.WORK_DIR, core.safe_component(stem), core.STATE_FILE))
    elif params['kind'] == 'files':
        # chosen files always write their md again; archives are named after their members
        plain = [{'rel': os.path.basename(p), 'path': p} for p in params['files']
                 if file_kind(os.path.basename(p)) == 'doc']
        replace = sorted(b + '.md' for b in core.base_names(plain)
                         if os.path.isfile(os.path.join(out, b + '.md')))
    elif params['force'] and not params['merge']:
        try:
            existing = sum(1 for n in os.listdir(out) if n.endswith('.md') and not n.startswith('_'))
        except OSError:
            existing = 0
    return {'replace': replace, 'existing_md': existing, 'out': out, 'count': params['count'],
            'unfinished': unfinished,
            'skipped_pdf': params['skipped_pdf'], 'skipped_legacy': params['skipped_legacy']}


# ---------------------------------------------------------------- clean-up of duplicate text

def label_re(text):
    """A doc2kb_dedup.py output label as a pattern: %d and %s become groups."""
    parts = re.split(r'(%d|%s)', text)
    return ''.join('(\\d+)' if p == '%d' else '(.+)' if p == '%s' else re.escape(p) for p in parts)


# the dedup tool's summary lines, in either language
DEDUP_LINES = {key: re.compile('^(?:%s)$' % '|'.join(label_re(t) for t in dedup.LABELS[key]))
               for key in ('summary', 'report', 'backup')}


def dedup_plan(data):
    """Check a clean-up request; returns (params, argv) or raises ValueError."""
    target = norm_path(data.get('target'), M('file'))
    if not os.path.isfile(target):
        raise ValueError(M('file_missing', target))
    if not target.lower().endswith('.md'):
        raise ValueError(M('not_md', target))
    bases = []
    for b in data.get('base') or []:
        path = norm_path(b, M('file'))
        if not os.path.exists(path):
            raise ValueError(M('file_missing', path))
        bases.append(path)
    params = {'target': target, 'base': bases, 'dry_run': bool(data.get('dry_run')),
              'no_db': bool(data.get('no_db'))}
    argv = [sys.executable, '-u', DEDUP, target, '--log-level', 'WARNING']
    for b in bases:
        argv += ['-b', b]
    if params['dry_run']:
        argv.append('-n')
    if params['no_db']:
        argv.append('-N')
    return params, argv


# ---------------------------------------------------------------- incremental tasks

def tasks_conf():
    """The tasks file, shared with earlier versions: one task per line, name<TAB>source<TAB>output."""
    return os.path.join(os.path.expanduser('~'), '.config', 'doc2kb', 'watch_tasks.conf')


def read_tasks():
    tasks = []
    try:
        with open(tasks_conf(), encoding='utf-8', errors='replace') as f:
            for line in f:
                parts = line.rstrip('\n').split('\t')
                if len(parts) >= 3 and parts[0]:
                    tasks.append({'name': parts[0], 'src': parts[1], 'out': parts[2]})
    except OSError:
        pass    # no file yet: no tasks
    return tasks


def write_tasks(tasks):
    path = tasks_conf()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        for t in tasks:
            f.write('%s\t%s\t%s\n' % (t['name'], t['src'], t['out']))
    os.replace(tmp, path)


def last_run(out):
    """When doc2kb.sh last finished a run into this output folder, or None (no watermark yet)."""
    try:
        with open(os.path.join(out, '.doc2kb_watch_state'), encoding='utf-8') as f:
            for line in f:
                if line.startswith('last_run='):
                    return line.strip().split('=', 1)[1] or None
    except OSError:
        pass    # no state file: never run
    return None


def task_fields(data):
    name = re.sub(r'[\t\r\n]', '', str(data.get('name') or '')).strip()
    if not name:
        raise ValueError(M('task_name'))
    src = existing_dir(data.get('src'), M('src'))
    out = norm_path(data.get('out'), M('out'))
    if out == src:
        raise ValueError(M('out_is_src'))
    if not writable_target(out):
        raise ValueError(M('out_not_writable', out))
    return {'name': name, 'src': src, 'out': out}


def change_tasks(data):
    """Add, edit or delete a task; raises ValueError with the reason."""
    tasks = read_tasks()
    names = [t['name'] for t in tasks]
    action = data.get('action')
    if action == 'add':
        t = task_fields(data)
        if t['name'] in names:
            raise ValueError(M('task_exists', t['name']))
        tasks.append(t)
    elif action == 'edit':
        old = data.get('old')
        if old not in names:
            raise ValueError(M('task_missing', old))
        t = task_fields(data)
        if t['name'] != old and t['name'] in names:
            raise ValueError(M('task_exists', t['name']))
        tasks[names.index(old)] = t
    elif action == 'delete':
        if data.get('name') not in names:
            raise ValueError(M('task_missing', data.get('name')))
        del tasks[names.index(data.get('name'))]
    else:
        raise ValueError(M('not_found'))
    write_tasks(tasks)


def task_steps(data):
    """The steps of a tasks run: doc2kb.sh -s SOURCE -o OUTPUT, with -d DAYS before the first run
    and -n for a preview."""
    try:
        days = int(data.get('days') or 30)
    except (TypeError, ValueError):
        days = 0
    if not 1 <= days <= 36500:
        raise ValueError(M('bad_days'))
    dry = bool(data.get('dry_run'))
    by_name = {t['name']: t for t in read_tasks()}
    wanted = [n for n in (data.get('names') or []) if n in by_name]
    if not wanted:
        raise ValueError(M('no_task'))
    steps = []
    for name in wanted:
        t = by_name[name]
        first = last_run(t['out']) is None
        step = {'name': name, 'first_run': first, 'dry_run': dry,
                'log': os.path.join(t['out'], '_watch.log')}
        if not os.path.isdir(t['src']):
            step['argv'] = None
            step['error'] = M('path_missing', M('src'), t['src'])
        else:
            argv = ['bash', SHELL, '-s', t['src'], '-o', t['out']]
            if first:
                argv += ['-d', str(days)]
            if dry:
                argv.append('-n')
            step['argv'] = argv
        steps.append(step)
    return steps


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
                    log.info('page %s heartbeat lost', log_safe(page))
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
        log.debug('%s - %s', self.address_string(), log_safe(fmt % args))

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
            log.warning('rejected %s %s', self.command, log_safe(urlsplit(self.path).path))
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
                'out_suffix': core.T('out_dir_suffix'),
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
        elif url.path == '/api/tasks':
            self.send_json(200, {'conf': tasks_conf(), 'tasks': [
                dict(t, last_run=last_run(t['out']), src_ok=os.path.isdir(t['src']))
                for t in read_tasks()]})
        else:
            self.send_json(404, {'error': M('not_found')})

    def do_POST(self):
        url = urlsplit(self.path)
        if url.path == '/api/bye':
            # sent by sendBeacon, which cannot set headers: the token is in the body
            data = self.read_json()
            if not self.allowed(True, data.get('t')):
                return
            log.info('page %s closed', log_safe(data.get('page')))
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
                params, argv, lines, list_opt = run_plan(data)
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
            if lines is not None:
                list_file = os.path.join(tmp, 'list.txt')
                with open(list_file, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(lines) + '\n')
                argv += [list_opt, list_file]
            if not job.start(params['kind'], params, argv, tmpdir=tmp):
                shutil.rmtree(tmp, ignore_errors=True)
                self.send_json(409, {'error': M('busy')})
                return
            self.send_json(200, {'ok': True})
        elif url.path == '/api/tasks':
            try:
                change_tasks(data)
            except ValueError as e:
                self.send_json(400, {'error': str(e)})
                return
            self.send_json(200, {'ok': True})
        elif url.path == '/api/tasks/run':
            try:
                steps = task_steps(data)
            except ValueError as e:
                self.send_json(400, {'error': str(e)})
                return
            params = {'names': [st['name'] for st in steps], 'dry_run': bool(data.get('dry_run'))}
            if not job.start('tasks', params, None, steps=steps):
                self.send_json(409, {'error': M('busy')})
                return
            self.send_json(200, {'ok': True})
        elif url.path == '/api/dedup':
            try:
                params, argv = dedup_plan(data)
            except ValueError as e:
                self.send_json(400, {'error': str(e)})
                return
            if not job.start('dedup', params, argv):
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
                out = job.params.get('out') if job.kind in ('run', 'files') and job.params else None
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
        pass  # Ctrl+C in a terminal: stop like "Quit"
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
