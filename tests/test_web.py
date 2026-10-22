"""Tests for the browser interface (doc2kb_web.py): access checks, the folder list, a scan and a
run through the service, cancelling, stale reports and the exit rules.
Standard library only: python3 -m unittest discover tests"""
import http.client
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TESTS = os.path.dirname(os.path.abspath(__file__))
WEB = os.path.join(ROOT, "doc2kb_web.py")
sys.path.insert(0, TESTS)
sys.path.insert(0, ROOT)
import doc2kb_fixtures as fx  # noqa: E402
import doc2kb_web as web  # noqa: E402


class Server:
    """doc2kb_web.py in a child process, without a browser."""

    def __init__(self, lang="en"):
        env = dict(os.environ, PYTHONIOENCODING="utf-8", LOG_LEVEL="WARNING")
        self.proc = subprocess.Popen([sys.executable, WEB, "--no-browser", "--lang", lang],
                                     stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                                     universal_newlines=True, encoding="utf-8", env=env)
        line = self.proc.stdout.readline()
        url = line[line.index("http://"):].strip()
        self.base, self.token = url.split("/?t=")
        self.port = int(self.base.rsplit(":", 1)[1])

    def request(self, method, path, body=None, token=True, headers=None):
        h = {"Host": "127.0.0.1:%d" % self.port}
        if token:
            h["X-Token"] = self.token
        h.update(headers or {})
        data = None
        if body is not None:
            data = json.dumps(body).encode("utf-8")
            h["Content-Type"] = "application/json"
        c = http.client.HTTPConnection("127.0.0.1", self.port, timeout=30)
        c.request(method, path, body=data, headers=h)
        r = c.getresponse()
        raw = r.read()
        c.close()
        try:
            return r.status, json.loads(raw.decode("utf-8"))
        except ValueError:
            return r.status, raw.decode("utf-8", "replace")

    def wait_job(self, timeout=60):
        end = time.time() + timeout
        while time.time() < end:
            _, job = self.request("GET", "/api/job")
            if job.get("state") not in ("running", "idle"):
                return job
            time.sleep(0.2)
        raise AssertionError("job did not finish")

    def stop(self):
        if self.proc.poll() is None:
            self.proc.kill()
        self.proc.wait(10)
        self.proc.stdout.close()


class WebTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="doc2kb_web_test_")
        self.src = os.path.join(self.tmp, "src")
        os.makedirs(self.src)
        fx.build_all(self.src, english=True)
        self.srv = Server()
        self.addCleanup(self.srv.stop)
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_access_checks(self):
        st, _ = self.srv.request("GET", "/api/status", token=False)
        self.assertEqual(st, 403)
        st, _ = self.srv.request("GET", "/api/status", headers={"X-Token": "wrong"})
        self.assertEqual(st, 403)
        st, _ = self.srv.request("GET", "/api/status", headers={"Host": "localhost:%d" % self.srv.port})
        self.assertEqual(st, 403)
        st, _ = self.srv.request("GET", "/api/status", headers={"Origin": "http://example.com"})
        self.assertEqual(st, 403)
        st, _ = self.srv.request("POST", "/api/bye", body={"t": "wrong", "page": "x"}, token=False)
        self.assertEqual(st, 403)
        st, page = self.srv.request("GET", "/", token=False)
        self.assertEqual(st, 200)
        self.assertIn("<html", page)
        st, s = self.srv.request("GET", "/api/status")
        self.assertEqual(st, 200)
        self.assertEqual(s["lang"], "en")

    def test_folder_list(self):
        os.makedirs(os.path.join(self.src, "sub"))
        os.makedirs(os.path.join(self.src, ".hidden"))
        st, d = self.srv.request("GET", "/api/ls?path=" + self.src)
        self.assertEqual(st, 200)
        self.assertEqual(d["dirs"], ["sub"])
        kinds = {f["name"]: f["kind"] for f in d["files"]}
        self.assertEqual(kinds["report-en.docx"], "doc")
        self.assertEqual(kinds["legacy-en.doc"], "legacy")
        st, d = self.srv.request("GET", "/api/ls?path=relative/path")
        self.assertEqual(st, 400)
        # a folder still to be created: the nearest existing one is listed
        st, d = self.srv.request("GET", "/api/ls?path=" + os.path.join(self.src, "new", "deeper"))
        self.assertEqual((st, d["path"]), (200, self.src))

    def test_scan_and_run(self):
        st, _ = self.srv.request("POST", "/api/scan", body={"src": self.src})
        self.assertEqual(st, 200)
        job = self.srv.wait_job()
        rels = [f["rel"] for f in job["result"]["files"]]
        self.assertIn("report-en.docx", rels)
        self.assertEqual(job["result"]["legacy"], ["legacy-en.doc"])
        out = job["result"]["out_default"]
        self.assertEqual(out, os.path.join(self.src, "_doc2kb"))
        body = {"src": self.src, "out": out, "only": ["report-en.docx", "mail-en.eml"],
                "merge": True, "inline": True}
        st, check = self.srv.request("POST", "/api/check", body=body)
        self.assertEqual((st, check["replace"]), (200, []))
        st, _ = self.srv.request("POST", "/api/run", body=body)
        self.assertEqual(st, 200)
        job = self.srv.wait_job()
        self.assertEqual((job["state"], job["rc"]), ("done", 0), job["output"])
        r = job["result"]
        self.assertEqual(r["stats"]["new"], "2")
        self.assertEqual(r["failed"], [])
        self.assertIn("Batch report", r["report"])
        self.assertTrue(os.path.isfile(os.path.join(out, "src.md")))
        self.assertFalse(os.path.isdir(os.path.join(out, "images")), "images are off by default")
        # a second run would replace the merged file: the check says so
        st, check = self.srv.request("POST", "/api/check", body=body)
        self.assertEqual(check["replace"], ["src.md"])

    def test_run_refusals(self):
        st, d = self.srv.request("POST", "/api/run", body={"src": self.src, "only": []})
        self.assertEqual(st, 400)
        st, d = self.srv.request("POST", "/api/run", body={"src": self.src, "out": self.src, "only": ["a.docx"]})
        self.assertEqual(st, 400)
        st, d = self.srv.request("POST", "/api/scan", body={"src": self.src, "since": "yesterday"})
        self.assertEqual(st, 400)

    def test_page_closed_exits(self):
        self.srv.request("POST", "/api/ping", body={"page": "p1"})
        self.srv.request("POST", "/api/bye", body={"t": self.srv.token, "page": "p1"}, token=False)
        self.srv.proc.wait(web.BYE_GRACE + 10)
        self.assertEqual(self.srv.proc.returncode, 0)


class JobTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="doc2kb_web_job_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        web.core.set_lang("en")

    def wait(self, job, timeout=30):
        end = time.time() + timeout
        while job.running and time.time() < end:
            time.sleep(0.1)
        self.assertFalse(job.running)

    def test_cancel_counts_written_files(self):
        script = ("import time,os\nprint('[1/3] a.docx', flush=True)\n"
                  "open(os.path.join(%r, 'a.md'), 'w').close()\n"
                  "print('[2/3] b.docx', flush=True)\ntime.sleep(60)\n") % self.tmp
        job = web.Job()
        job.start("run", {"src": self.tmp, "out": self.tmp, "merge": False}, [sys.executable, "-c", script])
        end = time.time() + 10
        while job.snapshot()["cur"] < 2 and time.time() < end:
            time.sleep(0.1)
        self.assertTrue(job.cancel())
        self.wait(job)
        snap = job.snapshot()
        self.assertEqual(snap["state"], "cancelled")
        self.assertEqual(snap["result"]["written_md"], 1)
        self.assertIsNone(snap["result"]["report"])

    def test_stale_report_not_shown(self):
        rp = os.path.join(self.tmp, web.core.T("report_file"))
        with open(rp, "w", encoding="utf-8") as f:
            f.write("an earlier run\n")
        os.utime(rp, (time.time() - 3600, time.time() - 3600))
        job = web.Job()
        job.start("run", {"src": self.tmp, "out": self.tmp, "merge": True},
                  [sys.executable, "-c", "import sys; print('Traceback'); sys.exit(1)"])
        self.wait(job)
        snap = job.snapshot()
        self.assertEqual(snap["rc"], 1)
        self.assertIsNone(snap["result"]["report"])
        self.assertEqual(snap["result"]["summary"], [])

    def test_no_exit_while_running(self):
        st = web.WebState(1, "t")
        st.quit.set()
        st.job.running = True
        self.assertIsNone(st.should_exit())
        st.job.running = False
        self.assertEqual(st.should_exit(), "quit requested")

    def test_pdf_skipped_without_pdftotext(self):
        os.makedirs(os.path.join(self.tmp, "s"))
        body = {"src": os.path.join(self.tmp, "s"), "only": ["a.pdf", "b.docx"]}
        with mock.patch.object(web, "has_tool", lambda name: False):
            params, argv, only = web.run_plan(body)
        self.assertEqual(only, ["b.docx"])
        self.assertEqual(params["skipped_pdf"], ["a.pdf"])
        self.assertIn("--no-images", argv)
        self.assertEqual(params["out"], os.path.join(self.tmp, "s", "_doc2kb"))


if __name__ == "__main__":
    unittest.main()
