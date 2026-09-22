// Run after make; requires Puppeteer and a Chromium browser. No native bridge.
const assert = require('node:assert/strict');
const path = require('node:path');
const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({
    executablePath: process.env.CHROMIUM_PATH,
    headless: true,
    args: ['--no-sandbox'],
  });
  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 480, height: 540 });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.evaluateOnNewDocument(() => {
      window.messages = [];
      window.chrome = { webview: { postMessage: text => window.messages.push(JSON.parse(text)) } };
    });
    await page.goto('file://' + path.resolve(__dirname, '../assets/dist/index.html'));
    await page.waitForFunction(() => window.messages.some(m => m.action === 'getInit'));
    await page.evaluate(() => window.onInit({ view: 'config', config: {
      url: '', healthyInterval: 60, downInterval: 10, loggingEnabled: false,
      historyLimit: 100, autoCheckForUpdates: false, updateCheckPending: false,
      updatePromptPending: false,
    } }));
    const click = async (text, modal = false) => {
      await page.waitForFunction((text, modal) => [...document.querySelectorAll(
        modal ? '[role="alertdialog"] button' : 'button')].some(b => b.textContent === text), {}, text, modal);
      await page.evaluate((text, modal) => [...document.querySelectorAll(
        modal ? '[role="alertdialog"] button' : 'button')].find(b => b.textContent === text).click(), text, modal);
    };
    const result = async (status, automatic = false) => {
      await page.evaluate((status, automatic) => window.onUpdateResult({
        status, title: 'Update available', message: 'A newer version is ready to install.',
        currentVersion: '1.0.7', remoteVersion: '1.0.8', automatic,
      }), status, automatic);
      await page.waitForFunction(() => !!document.querySelector('[role="alertdialog"]'));
    };
    const checkbox = '[role="alertdialog"] input[type="checkbox"]';
    const unchecked = async () => {
      await page.waitForSelector(checkbox);
      assert.equal(await page.$eval(checkbox, e => e.checked), false);
      assert.equal(await page.$eval(checkbox, e => e.parentElement.textContent.trim()), 'Reopen settings after update');
    };
    await click('Update');
    await page.waitForFunction(() => document.body.textContent.includes('Checking...'));
    // Each expected label differs from the previous one, so every step waits for a re-render.
    for (const [percent, expected] of [[0, 0], [7, 7], [42.9, 42], [150, 100], [-5, 0], [100, 100], [42, 42]]) {
      await page.evaluate(percent => window.onUpdateProgress({percent}), percent);
      const label = `Checking (${expected}%)...`;
      await page.waitForFunction(label => [...document.querySelectorAll('button')].some(b =>
        b.textContent === label && getComputedStyle(b).backgroundColor === 'rgb(220, 38, 38)'), {}, label);
      const busy = await page.evaluate(label => {
        const button = [...document.querySelectorAll('button')].find(b => b.textContent === label);
        return { disabled: button.disabled, digits: getComputedStyle(button).fontVariantNumeric };
      }, label);
      assert.equal(busy.disabled, false); // Clicking again stops the download.
      assert.equal(busy.digits, 'tabular-nums');
    }
    await click('Checking (42%)...');
    await page.waitForFunction(() => [...document.querySelectorAll('button')].some(b => b.textContent === 'Stopping...' && b.disabled));
    assert.equal(await page.evaluate(() => window.messages.filter(m => m.action === 'cancelUpdateCheck').length), 1);
    await page.evaluate(() => window.onUpdateResult({status: 'cancelled'}));
    await click('Update');
    await page.waitForFunction(() => document.body.textContent.includes('Checking...')); // Percentage cleared.
    await result('newer');
    await unchecked();
    await page.click(checkbox);
    await page.screenshot({path: '/tmp/apimonitor-update-modal.png'});
    await click('Cancel', true);
    await result('newer');
    await unchecked(); // Cancellation cannot retain the choice.
    await click('Update', true);
    await page.waitForFunction(() => window.messages.some(m => m.action === 'installUpdate'));
    assert.equal(await page.evaluate(() => window.messages.filter(m => m.action === 'installUpdate').at(-1).reopenSettingsAfterUpdate), false);
    await result('error');
    assert.equal(await page.$(checkbox), null);
    await click('OK', true);
    await result('same');
    await unchecked();
    await page.click(checkbox);
    await click('Force update', true);
    await page.waitForFunction(() => window.messages.filter(m => m.action === 'installUpdate').length === 2);
    assert.equal(await page.evaluate(() => window.messages.filter(m => m.action === 'installUpdate').at(-1).reopenSettingsAfterUpdate), true);
    assert.equal(await page.$eval(checkbox, e => e.disabled), true);
    await result('error');
    await click('OK', true);
    await result('newer', true);
    await unchecked(); // Failure cannot retain the choice.
    await page.click(checkbox);
    await click('Ignore this version', true);
    await result('newer');
    await unchecked();
    await click('Cancel', true);
    await result('older');
    assert.equal(await page.$(checkbox), null);
    assert.equal(await page.$$eval('[role="alertdialog"] button', buttons => buttons.map(b => b.textContent).join()), 'OK');
    assert.deepEqual(errors, []);
    console.log('Browser update UI, percentage label, styling, cancellation, choice reset, and bridge tests passed');
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
