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
import unittest.mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TESTS = os.path.dirname(os.path.abspath(__file__))
WEB = os.path.join(ROOT, "doc2kb_web.py")
sys.path.insert(0, TESTS)
sys.path.insert(0, ROOT)
import doc2kb_fixtures as fx  # noqa: E402
import doc2kb_web as web  # noqa: E402


class Server:
    """doc2kb_web.py in a child process, without a browser."""

    def __init__(self, lang="en", home=None, log_level="WARNING"):
        env = dict(os.environ, PYTHONIOENCODING="utf-8", LOG_LEVEL=log_level)
        if home:
            env["HOME"] = home
        self.err = tempfile.TemporaryFile(mode="w+", encoding="utf-8")
        self.proc = subprocess.Popen([sys.executable, WEB, "--no-browser", "--lang", lang],
                                     stdout=subprocess.PIPE, stderr=self.err,
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
        self.err.close()

    def stderr(self):
        self.err.seek(0)
        return self.err.read()


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
        self.assertEqual(s["out_suffix"], "_doc2kb")

    def test_folder_list(self):
        os.makedirs(os.path.join(self.src, "sub"))
        os.makedirs(os.path.join(self.src, ".hidden"))
        st, d = self.srv.request("GET", "/api/ls?path=" + self.src)
        self.assertEqual(st, 200)
        self.assertEqual(d["dirs"], ["sub"])
        kinds = {f["name"]: f["kind"] for f in d["files"]}
        self.assertEqual(kinds["report-en.docx"], "doc")
        self.assertNotIn("notes.txt", kinds)
        self.assertEqual(kinds["legacy-en.msg"], "legacy")
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
        self.assertEqual(job["result"]["legacy"], ["legacy-en.msg"])
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

    def test_run_all_in_batches(self):
        self.srv.request("POST", "/api/scan", body={"src": self.src})
        files = self.srv.wait_job()["result"]["files"]
        out = os.path.join(self.tmp, "out")
        # every file ticked: no list is sent; two files per batch, so the merged text comes in volumes
        body = {"src": self.src, "out": out, "all": True, "count": len(files), "merge": True, "batch_size": 2}
        st, check = self.srv.request("POST", "/api/check", body=body)
        self.assertEqual((st, check["replace"], check["unfinished"]), (200, [], False))
        st, _ = self.srv.request("POST", "/api/run", body=body)
        self.assertEqual(st, 200)
        job = self.srv.wait_job()
        self.assertEqual((job["state"], job["rc"]), ("done", 0), job["output"])
        n = -(-len(files) // 2)
        self.assertEqual((job["batch"], job["batches"]), (n, n))
        self.assertNotIn("BATCH\t", job["output"])                 # machine lines stay off the page
        self.assertTrue(os.path.isfile(os.path.join(out, "src-01.md")))
        self.assertTrue(os.path.isfile(os.path.join(out, "src-index.md")))
        self.assertIn("Volumes:", job["result"]["report"])
        # a reloaded page gets the same task back, with its settings and result
        st, status = self.srv.request("GET", "/api/status")
        self.assertEqual(status["job"]["params"]["out"], out)
        self.assertEqual(status["job"]["params"]["batch_size"], 2)
        self.assertEqual(status["job"]["result"]["stats"], job["result"]["stats"])
        st, check = self.srv.request("POST", "/api/check", body=body)
        self.assertIn("src-01.md", check["replace"])
        self.assertIn("src-index.md", check["replace"])

    def test_unfinished_merged_run_is_reported(self):
        out = os.path.join(self.tmp, "out")
        state_dir = os.path.join(out, web.core.WORK_DIR, "src")
        os.makedirs(state_dir)
        with open(os.path.join(state_dir, web.core.STATE_FILE), "w", encoding="utf-8") as f:
            f.write("{}")
        body = {"src": self.src, "out": out, "only": ["report-en.docx"], "merge": True}
        st, check = self.srv.request("POST", "/api/check", body=body)
        self.assertEqual((st, check["unfinished"]), (200, True))
        st, check = self.srv.request("POST", "/api/check", body=dict(body, merge=False))
        self.assertFalse(check["unfinished"])

    def test_files_run(self):
        files = [os.path.join(self.src, n) for n in ("report-en.docx", "mail-en.eml", "legacy-en.msg")]
        body = {"mode": "files", "files": files, "merge": True, "inline": True}
        st, check = self.srv.request("POST", "/api/check", body=body)
        self.assertEqual(st, 200, check)
        self.assertEqual(check["skipped_legacy"], [files[2]])
        self.assertEqual(check["out"], os.path.join(self.src, "_doc2kb"))
        self.assertEqual(check["count"], 2)
        st, _ = self.srv.request("POST", "/api/run", body=body)
        self.assertEqual(st, 200)
        job = self.srv.wait_job()
        self.assertEqual((job["kind"], job["state"], job["rc"]), ("files", "done", 0), job["output"])
        self.assertEqual(job["result"]["stats"]["new"], "2")
        self.assertEqual(job["result"]["failed_n"], 0)
        self.assertIn("Batch report", job["result"]["report"])
        self.assertTrue(os.path.isfile(os.path.join(self.src, "_doc2kb", "src.md")))

    def test_files_check_names_replaced_md(self):
        out = os.path.join(self.tmp, "out")
        os.makedirs(out)
        open(os.path.join(out, "report-en.md"), "w").close()
        body = {"mode": "files", "out": out, "merge": False,
                "files": [os.path.join(self.src, "report-en.docx"), os.path.join(self.src, "mail-en.eml")]}
        st, check = self.srv.request("POST", "/api/check", body=body)
        self.assertEqual((st, check["replace"]), (200, ["report-en.md"]))

    def test_files_refusals(self):
        for files in ([], ["relative.docx"], [os.path.join(self.src, "missing.docx")],
                      [os.path.join(self.src, "legacy-en.msg")]):
            st, d = self.srv.request("POST", "/api/run", body={"mode": "files", "files": files})
            self.assertEqual(st, 400, (files, d))
            self.assertNotIn("source folder", d["error"])

    def test_batch_size_refusals(self):
        body = {"src": self.src, "out": os.path.join(self.tmp, "out"), "only": ["report-en.docx"]}
        for size in (-1, "10", 1.5, True, 1000001):
            st, r = self.srv.request("POST", "/api/check", body=dict(body, batch_size=size))
            self.assertEqual(st, 400, size)
            self.assertIn("per batch", r["error"])

    def test_run_refusals(self):
        st, d = self.srv.request("POST", "/api/run", body={"src": self.src, "only": []})
        self.assertEqual(st, 400)
        st, d = self.srv.request("POST", "/api/run", body={"src": self.src, "out": self.src, "only": ["a.docx"]})
        self.assertEqual(st, 400)
        st, d = self.srv.request("POST", "/api/scan", body={"src": self.src, "since": "yesterday"})
        self.assertEqual(st, 400)

    def test_folder_list_shows_md(self):
        with open(os.path.join(self.src, "kb.md"), "w", encoding="utf-8") as f:
            f.write("# kb\n")
        st, d = self.srv.request("GET", "/api/ls?path=" + self.src)
        kinds = {f["name"]: f["kind"] for f in d["files"]}
        self.assertEqual(kinds["kb.md"], "md")

    def test_page_closed_exits(self):
        self.srv.request("POST", "/api/ping", body={"page": "p1"})
        self.srv.request("POST", "/api/bye", body={"t": self.srv.token, "page": "p1"}, token=False)
        self.srv.proc.wait(web.BYE_GRACE + 10)
        self.assertEqual(self.srv.proc.returncode, 0)


class TaskTest(unittest.TestCase):
    """Incremental tasks: kept in ~/.config/doc2kb/watch_tasks.conf, run through doc2kb.sh."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="doc2kb_web_task_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.home = os.path.join(self.tmp, "home")
        os.makedirs(self.home)
        self.src = os.path.join(self.tmp, "src")
        os.makedirs(self.src)
        fx.build_all(self.src, english=True)
        self.out = os.path.join(self.tmp, "out")
        self.srv = Server(home=self.home)
        self.addCleanup(self.srv.stop)

    def tasks(self):
        st, d = self.srv.request("GET", "/api/tasks")
        self.assertEqual(st, 200, d)
        return d["tasks"]

    def edit(self, **body):
        return self.srv.request("POST", "/api/tasks", body=body)

    def test_add_edit_delete(self):
        st, d = self.edit(action="add", name="Project\tAtlas", src=self.src, out=self.out)
        self.assertEqual(st, 200, d)
        conf = os.path.join(self.home, ".config", "doc2kb", "watch_tasks.conf")
        with open(conf, encoding="utf-8") as f:
            self.assertEqual(f.read(), "ProjectAtlas\t%s\t%s\n" % (self.src, self.out))
        self.assertEqual([t["name"] for t in self.tasks()], ["ProjectAtlas"])
        self.assertEqual(self.edit(action="add", name="ProjectAtlas", src=self.src, out=self.out)[0], 400)
        self.assertEqual(self.edit(action="add", name="x", src=os.path.join(self.tmp, "nope"), out=self.out)[0], 400)
        self.assertEqual(self.edit(action="add", name="", src=self.src, out=self.out)[0], 400)
        self.assertEqual(self.edit(action="add", name="Other", src=self.src, out=self.out + "2")[0], 200)
        st, d = self.edit(action="edit", old="Other", name="ProjectAtlas", src=self.src, out=self.out)
        self.assertEqual(st, 400, "renaming onto another task's name is refused")
        st, d = self.edit(action="edit", old="Other", name="Riverside", src=self.src, out=self.out + "3")
        self.assertEqual(st, 200, d)
        self.assertEqual([t["name"] for t in self.tasks()], ["ProjectAtlas", "Riverside"])
        self.assertEqual(self.edit(action="delete", name="Riverside")[0], 200)
        self.assertEqual([t["name"] for t in self.tasks()], ["ProjectAtlas"])
        self.assertTrue(os.path.isdir(self.src), "deleting a task touches no files")

    def run_tasks(self, **body):
        st, d = self.srv.request("POST", "/api/tasks/run", body=body)
        self.assertEqual(st, 200, d)
        job = self.srv.wait_job(120)
        self.assertEqual((job["kind"], job["state"]), ("tasks", "done"), job["output"])
        return job["result"]["tasks"]

    def test_first_run_takes_days_then_watermark(self):
        self.edit(action="add", name="Atlas", src=self.src, out=self.out)
        self.assertIsNone(self.tasks()[0]["last_run"])
        r = self.run_tasks(names=["Atlas"], days=30)
        self.assertEqual((r[0]["rc"], r[0]["first_run"]), (0, True), r)
        self.assertIn("src_", r[0]["last_line"])
        self.assertEqual(r[0]["log"], os.path.join(self.out, "_watch.log"))
        self.assertTrue(self.tasks()[0]["last_run"])
        r = self.run_tasks(names=["Atlas"], days=30)
        self.assertEqual((r[0]["rc"], r[0]["first_run"]), (0, False), r)

    def test_preview_writes_nothing(self):
        self.edit(action="add", name="Atlas", src=self.src, out=self.out)
        r = self.run_tasks(names=["Atlas"], days=30, dry_run=True)
        self.assertEqual(r[0]["rc"], 0, r)
        self.assertIn("report-en.docx", r[0]["output"])
        self.assertFalse([n for n in os.listdir(self.out) if n.endswith(".md")])

    def test_missing_source_is_reported_and_others_run(self):
        self.edit(action="add", name="Gone", src=self.src, out=self.out + "1")
        self.edit(action="add", name="Atlas", src=self.src, out=self.out)
        conf = os.path.join(self.home, ".config", "doc2kb", "watch_tasks.conf")
        with open(conf, encoding="utf-8") as f:
            text = f.read()
        with open(conf, "w", encoding="utf-8") as f:
            f.write(text.replace("Gone\t" + self.src, "Gone\t" + self.src + "-moved"))
        self.assertFalse(self.tasks()[0]["src_ok"])
        r = self.run_tasks(names=["Gone", "Atlas"], days=30)
        self.assertEqual([x["name"] for x in r], ["Gone", "Atlas"])
        self.assertIsNone(r[0]["rc"])
        self.assertEqual(r[1]["rc"], 0, r)


class DedupTest(unittest.TestCase):
    """The clean-up tab runs doc2kb_dedup.py: a dry run first, then the clean-up."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="doc2kb_web_dedup_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        src = os.path.join(self.tmp, "src")
        fx.build_all(src, english=True)
        self.base = os.path.join(self.tmp, "base")
        subprocess.run([sys.executable, os.path.join(ROOT, "doc2kb_core.py"), "--batch", src, "-o", self.base,
                        "--merge", "--merge-name", "full", "--no-images", "--lang", "en"],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        os.makedirs(os.path.join(self.tmp, "inc"))
        self.target = os.path.join(self.tmp, "inc", "big.md")
        shutil.copy(os.path.join(self.base, "full.md"), self.target)
        self.srv = Server()
        self.addCleanup(self.srv.stop)

    def dedup(self, **body):
        st, d = self.srv.request("POST", "/api/dedup", body=dict(target=self.target, base=[self.base], **body))
        self.assertEqual(st, 200, d)
        job = self.srv.wait_job()
        self.assertEqual((job["kind"], job["state"], job["rc"]), ("dedup", "done", 0), job["output"])
        return job["result"]

    def test_dry_run_then_clean(self):
        with open(self.target, encoding="utf-8") as f:
            before = f.read()
        r = self.dedup(dry_run=True)
        self.assertGreater(r["sections"], 0)
        self.assertEqual(r["duplicate"], r["sections"])
        self.assertIn("Duplicate text clean-up report", r["report"])
        self.assertIsNone(r["backup"])
        with open(self.target, encoding="utf-8") as f:
            self.assertEqual(f.read(), before)
        r = self.dedup(dry_run=False)
        self.assertTrue(r["backup"] and os.path.isfile(r["backup"]), r)
        with open(self.target, encoding="utf-8") as f:
            self.assertLess(len(f.read()), len(before))
        self.assertTrue(os.path.isfile(os.path.join(self.tmp, "inc", ".doc2kb_watch_fp")))

    def test_no_fingerprint_file_when_asked(self):
        self.dedup(dry_run=False, no_db=True)
        self.assertFalse(os.path.exists(os.path.join(self.tmp, "inc", ".doc2kb_watch_fp")))

    def test_refusals(self):
        for body in ({"target": os.path.join(self.tmp, "missing.md")},
                     {"target": os.path.join(self.base, "_batch-report.txt")},
                     {"target": self.target, "base": [os.path.join(self.tmp, "nope")]}):
            st, d = self.srv.request("POST", "/api/dedup", body=body)
            self.assertEqual(st, 400, (body, d))


class LogTest(unittest.TestCase):
    def test_request_values_cannot_forge_log_lines(self):
        srv = Server(log_level="INFO")
        self.addCleanup(srv.stop)
        srv.request("POST", "/api/bye", body={"t": srv.token, "page": "p1\nFORGED line"}, token=False)
        srv.request("GET", "/api/status%0AFORGED", token=False)
        time.sleep(0.3)
        lines = srv.stderr().splitlines()
        self.assertFalse([l for l in lines if l.startswith("FORGED")], lines)
        self.assertTrue([l for l in lines if "p1" in l], lines)


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

    def test_task_progress_with_one_task(self):
        script = ("import time\nprint('[1/4] a.docx', flush=True)\nprint('BATCH', 2, 3, sep=chr(9), flush=True)\n"
                  "print('[3/4] c.docx', flush=True)\ntime.sleep(60)\n")
        job = web.Job()
        job.start("tasks", {"names": ["one"], "dry_run": False}, None,
                  steps=[{"name": "one", "argv": [sys.executable, "-c", script], "first_run": False,
                          "dry_run": False, "log": ""}])
        end = time.time() + 10
        while job.snapshot()["cur"] < 3 and time.time() < end:
            time.sleep(0.1)
        snap = job.snapshot()
        self.assertEqual((snap["task"], snap["tasks"], snap["task_name"]), (1, 1, "one"))
        self.assertEqual((snap["cur"], snap["total"], snap["batch"], snap["batches"]), (3, 4, 2, 3))
        self.assertTrue(job.cancel())
        self.wait(job)

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
        with unittest.mock.patch.object(web, "has_tool", lambda name: False):
            params, argv, only, opt = web.run_plan(body)
        self.assertEqual(only, ["b.docx"])
        self.assertEqual(params["skipped_pdf"], ["a.pdf"])
        self.assertIn("--no-images", argv)
        self.assertEqual(params["out"], os.path.join(self.tmp, "s", "_doc2kb"))


if __name__ == "__main__":
    unittest.main()
