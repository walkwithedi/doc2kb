// Browser tests of Upgrade and Roll back on the doc2kb page:  node tests/upgrade.test.js
// A copy of the tool with an older VERSION serves the page; the working tree, zipped like a release,
// is the new version. Checks the messages, stage text, disabled buttons, the 360 px layout, the token,
// Host and Origin checks of the new endpoints, and that the page reconnects after the restart.
'use strict';
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const {spawn, execFileSync} = require('child_process');
const {chromium} = require('playwright');

const ROOT = path.resolve(__dirname, '..');
const FILES = ['doc2kb_core.py', 'doc2kb.sh', 'doc2kb_dedup.py', 'doc2kb_web.py', 'doc2kb.html', '文档转知识库.desktop',
  'tool_upgrade.py', 'upgrade-keep.txt', 'TOOL_UPGRADE_VERSION', 'SPEC.md', 'README.md', 'README.zh-CN.md', 'LICENSE',
  'CHANGELOG.md', 'VERSION'];
const ZIPPER = `
import os, stat, sys, zipfile
src, top, out = sys.argv[1], sys.argv[2], sys.argv[3]
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in sys.argv[4:]:
        p = os.path.join(src, n)
        zi = zipfile.ZipInfo(top + '/' + n)
        zi.create_system = 3
        zi.external_attr = (stat.S_IFREG | stat.S_IMODE(os.stat(p).st_mode)) << 16
        with open(p, 'rb') as f:
            z.writestr(zi, f.read())
`;

function prepare() {
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'd2k-up-'));
  const dir = path.join(work, 'apps', 'doc2kb');
  fs.mkdirSync(dir, {recursive: true});
  for (const n of FILES) {
    fs.copyFileSync(path.join(ROOT, n), path.join(dir, n));
    fs.chmodSync(path.join(dir, n), fs.statSync(path.join(ROOT, n)).mode & 0o777);
  }
  const newVer = fs.readFileSync(path.join(ROOT, 'VERSION'), 'utf8').trim();
  // the older version: the second CHANGELOG section
  const sections = fs.readFileSync(path.join(ROOT, 'CHANGELOG.md'), 'utf8').match(/^## (\d+\.\d+\.\d+)$/mg);
  const oldVer = sections[1].slice(3);
  fs.writeFileSync(path.join(dir, 'VERSION'), oldVer + '\n');
  fs.writeFileSync(path.join(dir, 'my-notes.txt'), 'notes\n');
  const zip = path.join(work, 'doc2kb-v' + newVer + '.zip');
  execFileSync('python3', ['-c', ZIPPER, ROOT, 'doc2kb', zip, ...FILES]);
  const bad = path.join(work, 'bad.zip');
  execFileSync('python3', ['-c', ZIPPER, ROOT, 'other', bad, 'VERSION']);
  return {work, dir, zip, bad, newVer: 'v' + newVer, oldVer: 'v' + oldVer};
}

function startServer(ctx, lang) {
  const env = Object.assign({}, process.env, {DOC2KB_LANG: lang, HOME: ctx.work});
  const proc = spawn('python3', [path.join(ctx.dir, 'doc2kb_web.py'), '--no-browser'], {env});
  let log = '';
  proc.stderr.on('data', d => { log += d; });
  return new Promise((resolve, reject) => {
    let out = '';
    proc.stdout.on('data', d => {
      out += d;
      const m = out.match(/(http:\/\/127\.0\.0\.1:\d+\/\?t=\S+)/);
      if (m) resolve({proc, url: m[1], log: () => log});
    });
    proc.on('exit', code => reject(new Error('server exited ' + code + ' ' + log)));
    setTimeout(() => reject(new Error('no address printed')), 20000);
  });
}

function raw(port, opts, body) {
  return new Promise((resolve, reject) => {
    const req = http.request({host: '127.0.0.1', port, agent: false, ...opts}, res => { res.resume(); res.on('end', () => resolve(res.statusCode)); });
    req.on('error', reject);
    req.end(body);
  });
}

async function noScroll(page, what) {
  assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= 360), 'horizontal scroll at 360px: ' + what);
}

(async () => {
  const browser = await chromium.launch(process.env.CHROMIUM_PATH ? {executablePath: process.env.CHROMIUM_PATH} : {});
  const ctx = prepare();
  const tmpBefore = new Set(fs.readdirSync(os.tmpdir()).filter(n => n.startsWith('tool-upgrade-')));
  const leftUploads = () => fs.readdirSync(os.tmpdir()).filter(n => n.startsWith('tool-upgrade-') && !tmpBefore.has(n));
  let srv;
  try {
    srv = await startServer(ctx, 'zh');
    const u = new URL(srv.url), token = u.searchParams.get('t'), host = '127.0.0.1:' + u.port;
    const bctx = await browser.newContext({viewport: {width: 360, height: 780}, bypassCSP: true});
    const page = await bctx.newPage();
    const errors = [], dialogs = [], stages = [];
    page.on('pageerror', e => errors.push(e.message));
    page.on('dialog', d => { dialogs.push(d.message()); d.accept(); });
    await page.exposeFunction('recordStage', t => { if (t && stages[stages.length - 1] !== t) stages.push(t); });
    await page.addInitScript(() => document.addEventListener('DOMContentLoaded', () => {
      const s = document.getElementById('up-status');
      new MutationObserver(() => window.recordStage(s.textContent)).observe(s, {childList: true, characterData: true, subtree: true});
    }));
    await page.goto(srv.url);
    await page.waitForFunction(() => document.getElementById('up-open').textContent === '升级');
    await page.click('#up-open');
    assert.ok(await page.isVisible('#upgrade'));
    await page.waitForFunction(() => document.getElementById('up-rollback').title !== '');
    assert.ok(await page.isDisabled('#up-rollback'), 'no previous version yet');
    assert.strictEqual(await page.textContent('#up-rollback'), '退回上一版');
    await noScroll(page, 'panel');

    // the new endpoints check the token, Host and Origin
    const body = fs.readFileSync(ctx.bad);
    assert.strictEqual(await raw(u.port, {method: 'POST', path: '/api/upgrade/check', headers: {Host: host}}, body), 403);
    assert.strictEqual(await raw(u.port, {method: 'POST', path: '/api/upgrade/check', headers: {Host: host, 'X-Token': token, Origin: 'http://evil.example'}}, body), 403);
    assert.strictEqual(await raw(u.port, {method: 'POST', path: '/api/upgrade/check', headers: {Host: 'evil.example', 'X-Token': token}}, body), 403);
    assert.strictEqual(await raw(u.port, {method: 'POST', path: '/api/upgrade/rollback', headers: {Host: host, 'X-Token': 'x'}}), 403);
    assert.strictEqual(await raw(u.port, {method: 'GET', path: '/api/upgrade', headers: {Host: host}}), 403);

    await page.setInputFiles('#up-file', ctx.bad);
    await page.waitForFunction(() => document.getElementById('up-status').className.includes('bad'));
    assert.match(await page.textContent('#up-status'), /不能升级：这不是本工具的发行包/);

    await page.setInputFiles('#up-file', ctx.zip);
    await page.waitForFunction(() => document.getElementById('up-status').className.includes('ok'));
    const report = await page.textContent('#up-report');
    assert.ok(report.includes('当前版本 ' + ctx.oldVer + ' → 新版本 ' + ctx.newVer + '（升级）'), report);
    assert.ok(report.includes(ctx.newVer + '：'), report);
    assert.match(report, /my-notes\.txt/);
    assert.strictEqual((await page.textContent('#up-go')).trim(), '升级到 ' + ctx.newVer);
    await noScroll(page, 'report');
    if (process.env.SHOT_DIR) await page.screenshot({path: path.join(process.env.SHOT_DIR, 'doc2kb-upgrade-360.png'), fullPage: true});

    const pid = srv.proc.pid;
    const loaded = page.waitForEvent('load', {timeout: 60000});
    await page.click('#up-go');
    assert.match(dialogs[dialogs.length - 1], /多出来的文件/);
    assert.ok(await page.isDisabled('#up-go'), 'the button is disabled while upgrading');
    assert.ok(await page.isDisabled('#up-rollback'));
    await loaded;
    assert.ok(stages.some(s => /正在/.test(s)), stages.join(' | '));
    assert.ok(stages.some(s => s.includes('已升级到 ' + ctx.newVer)), stages.join(' | '));
    assert.strictEqual(srv.proc.exitCode, null, 'the service keeps running');
    assert.strictEqual(srv.proc.pid, pid);
    const st = await page.evaluate(t => fetch('/api/upgrade', {headers: {'X-Token': t}}).then(r => r.json()), token);
    assert.strictEqual(st.current, ctx.newVer);
    assert.strictEqual(st.rollback.previous, ctx.oldVer);
    await page.waitForFunction(v => document.getElementById('version').textContent.includes(v.slice(1)), ctx.newVer);
    assert.ok(fs.existsSync(path.join(ctx.work, 'apps', '.doc2kb.上一版', 'my-notes.txt')));

    await page.click('#up-open');
    await page.waitForFunction(() => !document.getElementById('up-rollback').disabled);
    const back = page.waitForEvent('load', {timeout: 60000});
    await page.click('#up-rollback');
    assert.ok(dialogs[dialogs.length - 1].includes('会把当前 ' + ctx.newVer + ' 换成上一版 ' + ctx.oldVer), dialogs.join('|'));
    await back;
    const st2 = await page.evaluate(t => fetch('/api/upgrade', {headers: {'X-Token': t}}).then(r => r.json()), token);
    assert.strictEqual(st2.current, ctx.oldVer);
    assert.deepStrictEqual(leftUploads(), []);
    assert.deepStrictEqual(errors, []);
    srv.proc.kill();

    // English labels
    srv = await startServer(ctx, 'en');
    const p2 = await (await browser.newContext({bypassCSP: true})).newPage();
    await p2.goto(srv.url);
    await p2.waitForFunction(() => document.getElementById('up-open').textContent === 'Upgrade');
    await p2.click('#up-open');
    assert.strictEqual(await p2.textContent('#up-rollback'), 'Roll back to previous version');
    await p2.setInputFiles('#up-file', ctx.zip);
    await p2.waitForFunction(() => /new version/.test(document.getElementById('up-report').textContent));
    assert.match(await p2.textContent('#up-go'), /^Upgrade to v/);
    // quitting removes a checked zip that was never applied
    await p2.click('#quit');
    await new Promise(resolve => srv.proc.on('exit', resolve));
    assert.deepStrictEqual(leftUploads(), []);
    console.log('upgrade page tests passed');
  } catch (e) {
    if (srv) e.message += '\nservice log:\n' + srv.log().slice(-3000);
    throw e;
  } finally {
    await browser.close();
    if (srv && srv.proc.exitCode === null) srv.proc.kill();
    fs.rmSync(ctx.work, {recursive: true, force: true});
  }
})().catch(e => { console.error(e); process.exit(1); });
