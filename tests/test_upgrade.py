"""Tests of the upgrade-module copy check (scripts/check_tool_upgrade.py) and of --upgrade/--rollback."""
import hashlib
import io
import os
import shutil
import stat
import subprocess
import sys
import tempfile
import unittest
import zipfile
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))
sys.path.insert(0, ROOT)
import check_tool_upgrade as ctu  # noqa: E402

RELEASE = ['doc2kb_core.py', 'doc2kb.sh', 'doc2kb_dedup.py', 'doc2kb_web.py', 'doc2kb.html', '文档转知识库.desktop',
           'tool_upgrade.py', 'upgrade-keep.txt', 'TOOL_UPGRADE_VERSION', 'SPEC.md', 'README.md', 'README.zh-CN.md',
           'LICENSE', 'CHANGELOG.md', 'VERSION']


class CopyCheck(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix='d2k-ctu-')
        self.addCleanup(shutil.rmtree, self.tmp, True)
        for n in ('tool_upgrade.py', 'TOOL_UPGRADE_VERSION'):
            shutil.copy(os.path.join(ROOT, n), self.tmp)

    def test_repository_copy_is_pinned(self):
        self.assertEqual(ctu.offline_problems(ROOT), [])

    def test_edited_copy_and_wrong_version(self):
        with open(os.path.join(self.tmp, 'tool_upgrade.py'), 'a') as f:
            f.write('# edited\n')
        self.assertIn('differs', ctu.offline_problems(self.tmp)[0])
        with open(os.path.join(self.tmp, 'TOOL_UPGRADE_VERSION'), 'w') as f:
            f.write('9.9\nsha256 %s\n' % ('0' * 64))
        self.assertTrue(any('MODULE_VERSION' in p for p in ctu.offline_problems(self.tmp)))
        with open(os.path.join(self.tmp, 'TOOL_UPGRADE_VERSION'), 'w') as f:
            f.write('1.0\n')
        self.assertIn('must hold', ctu.offline_problems(self.tmp)[0])

    def test_against_master_history(self):
        master = os.path.join(self.tmp, 'master')
        path = os.path.join(master, '.github', 'upgrade')
        os.makedirs(path)
        git = lambda *a: subprocess.run(['git', '-C', master] + list(a), check=True, capture_output=True)
        git('init', '-q')
        git('config', 'user.email', 'test@example.com')
        git('config', 'user.name', 'test')
        with open(os.path.join(ROOT, 'tool_upgrade.py'), 'rb') as f:
            data = f.read()
        with open(os.path.join(path, 'tool_upgrade.py'), 'wb') as f:
            f.write(data)
        git('add', '-A')
        git('commit', '-qm', 'one')
        newer = data.replace(b"MODULE_VERSION = '", b"MODULE_VERSION = '9")
        with open(os.path.join(path, 'tool_upgrade.py'), 'wb') as f:
            f.write(newer)
        git('commit', '-qam', 'two')
        self.assertEqual(ctu.master_problems(master, self.tmp), [])       # found in history
        with open(os.path.join(self.tmp, 'tool_upgrade.py'), 'ab') as f:
            f.write(b'# x\n')
        self.assertIn('differs from every master revision', ctu.master_problems(master, self.tmp)[0])
        with open(os.path.join(self.tmp, 'TOOL_UPGRADE_VERSION'), 'w') as f:
            f.write('7.7\nsha256 %s\n' % hashlib.sha256(b'').hexdigest())
        self.assertIn('no revision', ctu.master_problems(master, self.tmp)[0])


class CommandLineUpgrade(unittest.TestCase):
    """doc2kb_web.py --upgrade (also through doc2kb.sh) from a copy with an older VERSION, then --rollback."""

    def test_upgrade_and_rollback(self):
        tmp = tempfile.mkdtemp(prefix='d2k-cli-up-')
        self.addCleanup(shutil.rmtree, tmp, True)
        tool = os.path.join(tmp, 'apps', 'doc2kb')
        os.makedirs(tool)
        for n in RELEASE:
            shutil.copy2(os.path.join(ROOT, n), tool)
        with open(os.path.join(tool, 'VERSION'), 'w') as f:
            f.write('0.0.1\n')
        z = os.path.join(tmp, 'new.zip')
        with zipfile.ZipFile(z, 'w') as zf:
            for n in RELEASE:
                zi = zipfile.ZipInfo('doc2kb/' + n)
                zi.create_system = 3
                zi.external_attr = (stat.S_IFREG | 0o644) << 16
                with open(os.path.join(ROOT, n), 'rb') as f:
                    zf.writestr(zi, f.read())
        env = dict(os.environ, DOC2KB_LANG='en')
        run = lambda *a, **kw: subprocess.run([sys.executable, os.path.join(tool, 'doc2kb_web.py')] + list(a),
                                              capture_output=True, text=True, env=env, **kw)
        r = run('--upgrade', z, input='n\n')
        self.assertEqual(r.returncode, 1)
        self.assertIn('Current version v0.0.1 → new version v', r.stdout)
        self.assertIn('Not confirmed', r.stderr)
        r = run('--upgrade', z, '--yes')
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn('Upgraded to v', r.stdout)
        self.assertTrue(os.path.isdir(os.path.join(tmp, 'apps', '.doc2kb.上一版')))
        # doc2kb.sh passes --rollback through
        r = subprocess.run(['bash', os.path.join(tool, 'doc2kb.sh'), '--rollback', '--yes'], capture_output=True,
                           text=True, env=env)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn('Rolled back to v0.0.1', r.stdout)
        with open(os.path.join(tool, 'VERSION')) as f:
            self.assertEqual(f.read().strip(), '0.0.1')

    def test_missing_module(self):
        sys.path.insert(0, ROOT)
        import doc2kb_web
        with mock.patch.object(doc2kb_web, 'tool_upgrade', None), mock.patch('sys.stderr', io.StringIO()) as err:
            self.assertEqual(doc2kb_web.main(['--rollback']), 1)
        self.assertIn('tool_upgrade.py', err.getvalue())


if __name__ == '__main__':
    unittest.main()
