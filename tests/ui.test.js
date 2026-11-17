// Browser tests of the doc2kb page layout and flows:  node tests/ui.test.js
// Steps of the folder batch, the run summary and start button, a conversion with its result figures, tabs
// that keep their state, the incremental task cards, the dry-run gate of the duplicate clean-up, both
// languages without missing texts, light and dark themes, 360 px and 1366×768. Needs Playwright (development only).
'use strict';
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {spawn, execFileSync} = require('node:child_process');
const {chromium} = require('playwright');

const ROOT = path.resolve(__dirname, '..');

function startServer(home, lang) {
  const env = Object.assign({}, process.env, {DOC2KB_LANG: lang, HOME: home});
  const proc = spawn('python3', [path.join(ROOT, 'doc2kb_web.py'), '--no-browser'], {env});
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

async function noScroll(page, width, what) {
  const w = await page.evaluate(() => document.documentElement.scrollWidth);
  assert.ok(w <= width, 'horizontal scroll at ' + width + 'px (' + w + '): ' + what);
}

(async () => {
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'doc2kb-ui-'));
  const home = path.join(work, 'home'), src = path.join(home, 'docs');
  fs.mkdirSync(src, {recursive: true});
  execFileSync('python3', ['-c', 'import sys; sys.path.insert(0, sys.argv[1]); import doc2kb_fixtures as f; ' +
    'f.docx(sys.argv[2] + "/a.docx", english=True); f.xlsx(sys.argv[2] + "/b.xlsx", english=True); f.pptx(sys.argv[2] + "/c.pptx", english=True)',
    path.join(ROOT, 'tests'), src]);
  const browser = await chromium.launch(fs.existsSync('/usr/bin/chromium') ? {executablePath: '/usr/bin/chromium'} : {});
  let srv;
  try {
    srv = await startServer(home, 'zh');
    const errors = [];
    const page = await browser.newPage({viewport: {width: 1366, height: 768}});
    page.on('pageerror', e => errors.push(e.message));
    page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
    page.on('dialog', d => { errors.push('native dialog: ' + d.message()); d.dismiss(); });
    await page.goto(srv.url);
    await page.waitForFunction(() => document.getElementById('quit').textContent === '退出');
    // steps 2 and 3 wait for the scan; the summary already shows the start button (disabled)
    assert.ok(await page.evaluate(() => document.getElementById('bstep2').classList.contains('is-locked')));
    assert.ok(await page.isVisible('#run'));
    assert.ok(await page.isDisabled('#run'));
    assert.doesNotMatch(await page.textContent('#bstep1'), /\d\. /, 'no "1." numbering inside the tab');
    await page.fill('#src', src);
    await page.click('#scan');
    await page.waitForFunction(() => !document.getElementById('listCard').hidden, null, {timeout: 60000});
    assert.ok(await page.evaluate(() => document.getElementById('bstep1').classList.contains('is-done')));
    assert.strictEqual(await page.locator('#fileRows tr').count(), 3);
    assert.match(await page.textContent('#run'), /开始转换 3 份/);
    const box = await page.locator('#run').boundingBox();
    assert.ok(box && box.y + box.height <= 768, 'the start button is on the first screen at 1366×768');
    await noScroll(page, 1366, 'scanned');
    // merge mode as radio cards
    await page.click('.opts[data-prefix="b"] .o-mode-single');
    assert.ok(await page.evaluate(() => document.querySelector('.opts[data-prefix="b"] .o-mode-single').classList.contains('is-checked')));
    assert.match(await page.textContent('#kMode'), /每份文档一个/);
    await page.click('.opts[data-prefix="b"] .o-mode-merge');
    // tabs keep their state
    await page.click('#tab-dedup');
    assert.ok(await page.isDisabled('#dedupClean'), 'clean-up needs a dry run first');
    assert.match(await page.textContent('#dedupBackupNote'), /\.bak-/);
    await page.click('#tab-batch');
    assert.strictEqual(await page.locator('#fileRows tr').count(), 3, 'scan result kept after switching tabs');
    // keyboard: arrow keys move between tabs one at a time
    await page.focus('#tab-batch');
    await page.keyboard.press('ArrowRight');
    assert.strictEqual(await page.getAttribute('#tab-files', 'aria-selected'), 'true');
    await page.click('#tab-batch');
    // run: the task bar shows while it runs, then the result card with figures
    await page.click('#run');
    await page.waitForFunction(() => !document.getElementById('batchResult').hidden, null, {timeout: 120000});
    assert.ok(await page.isHidden('#jobbar'));
    assert.match(await page.textContent('#batchResult .r-badge'), /已完成|需注意/);
    assert.ok(await page.isVisible('#batchResult .r-stats'));
    // a second run asks before replacing, in the page's own dialog
    await page.click('#run');
    await page.waitForFunction(() => document.getElementById('ask').open, null, {timeout: 20000});
    await page.click('#ask-cancel');
    assert.ok(!(await page.isDisabled('#run')));
    // incremental tasks: a card with a status dot
    await page.click('#tab-tasks');
    await page.click('#taskNew');
    await page.fill('#taskName', 'example');
    await page.fill('#taskSrc', src);
    await page.fill('#taskOut', path.join(home, 'inc'));
    await page.click('#taskSave');
    await page.waitForSelector('#taskRows .taskcard');
    assert.ok(await page.isVisible('#taskRows .taskcard .sdot'));
    assert.match(await page.textContent('#taskRows .taskcard'), /尚未运行|not run yet/);
    // delete asks first; Back keeps the task
    await page.click('#taskRows .t-del');
    await page.waitForFunction(() => document.getElementById('ask').open);
    await page.click('#ask-cancel');
    assert.strictEqual(await page.locator('#taskRows .taskcard').count(), 1);
    // English: every text exists
    await page.click('#lang button[data-lang="en"]');
    assert.strictEqual(await page.textContent('#quit'), 'Quit');
    assert.strictEqual(await page.textContent('#tab-tasks'), 'Incremental tasks');
    // dark theme and the narrow layout
    await page.click('[data-theme-toggle]');
    await page.setViewportSize({width: 360, height: 780});
    for (const tab of ['#tab-batch', '#tab-files', '#tab-tasks', '#tab-dedup']) { await page.click(tab); await noScroll(page, 360, tab); }
    await page.click('#lang button[data-lang="zh"]');
    assert.deepStrictEqual(errors, []);
    await page.goto('about:blank');
    console.log('ui page tests passed');
  } finally {
    if (srv) srv.proc.kill();
    await browser.close();
    fs.rmSync(work, {recursive: true, force: true});
  }
})().catch(e => { console.error(e); process.exitCode = 1; });
