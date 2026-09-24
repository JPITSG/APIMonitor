// Run after make; uses the built UI and a mocked native WebView bridge.
const assert = require('node:assert/strict');
const path = require('node:path');
const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({
    executablePath: process.env.CHROMIUM_PATH, headless: true, args: ['--no-sandbox'],
  });
  try {
    const page = await browser.newPage();
    await page.setViewport({width: 480, height: 410});
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    const config = {
      url: 'https://example.com/status', healthyInterval: 60, downInterval: 10,
      loggingEnabled: false, historyLimit: 100, startWithWindows: false,
      autoCheckForUpdates: false, updateCheckPending: false, updatePromptPending: false,
    };
    await page.evaluateOnNewDocument(() => {
      window.messages = [];
      window.chrome = {webview: {postMessage: text => window.messages.push(JSON.parse(text))}};
    });
    const reset = async () => {
      await page.goto('file://' + path.resolve(__dirname, '../assets/dist/index.html'));
      await page.waitForFunction(() => window.messages.some(m => m.action === 'getInit'));
      await page.evaluate(config => window.onInit({view: 'config', config}), config);
      await page.waitForFunction(() => window.messages.some(m => m.action === 'configReady'));
    };
    const click = async (text, modal = false) => {
      const scope = modal ? '[role="alertdialog"] button' : 'button';
      const buttons = await page.$$(scope);
      const button = (await Promise.all(buttons.map(async button => ({
        button, text: await button.evaluate(el => el.textContent),
      })))).find(button => button.text === text);
      assert.ok(button, text);
      await button.button.click();
    };
    const changeText = async (selector, value) => {
      await page.click(selector, {clickCount: 3});
      await page.keyboard.press('Backspace');
      if (value) await page.type(selector, value);
    };
    const actions = () => page.evaluate(() => window.messages.filter(m =>
      ['close', 'saveSettings'].includes(m.action)));
    const prompt = async () => {
      await page.waitForSelector('#save-alert-title');
      assert.equal(await page.$eval('#save-alert-message', el => el.textContent), 'Save changes before closing?');
      assert.deepEqual(await actions(), []);
      assert.equal(await page.$$eval('[role="alertdialog"]', el => el.length), 1);
      assert.equal(await page.$eval('#api-url', el => !!el.closest('[inert]')), true);
    };
    const noPrompt = () => page.waitForFunction(() => !document.querySelector('#save-alert-title'));
    const nativeClose = () => page.evaluate(() => window.onCloseRequested());

    // No changes: both footer and native close finish without a question.
    for (const trigger of [() => click('Cancel'), nativeClose]) {
      await reset();
      await trigger();
      await page.waitForFunction(() => window.messages.some(m => m.action === 'close'));
      assert.deepEqual(await actions(), [{action: 'close'}]);
      await noPrompt();
    }

    // Every editable setting triggers a question; reverting it restores clean closing.
    const settings = [
      [() => changeText('#api-url', 'https://example.com/changed'), () => changeText('#api-url', config.url)],
      [() => page.select('#healthy-interval', '120'), () => page.select('#healthy-interval', '60')],
      [() => page.select('#down-interval', '30'), () => page.select('#down-interval', '10')],
      [() => page.click('#logging'), () => page.click('#logging')],
      [() => changeText('#history-limit', '200'), () => changeText('#history-limit', '100')],
      [() => page.click('#start-with-windows'), () => page.click('#start-with-windows')],
      [() => page.click('#auto-update'), () => page.click('#auto-update')],
    ];
    for (const [edit, revert] of settings) {
      await reset();
      await edit();
      await nativeClose();
      await prompt();
      await nativeClose();
      await prompt();
      await click('Keep editing', true);
      await noPrompt();
      assert.deepEqual(await actions(), []);
      await revert();
      await click('Cancel');
      await page.waitForFunction(() => window.messages.some(m => m.action === 'close'));
      await noPrompt();
    }

    await reset();
    await page.click('#start-with-windows');
    await click('Cancel');
    await prompt();
    assert.equal(await page.evaluate(() => document.activeElement.textContent), 'Keep editing');
    await page.keyboard.press('Tab');
    await page.keyboard.press('Tab');
    assert.equal(await page.evaluate(() => document.activeElement.textContent), 'Save');
    await page.keyboard.press('Tab');
    assert.equal(await page.evaluate(() => document.activeElement.textContent), 'Keep editing');
    await page.keyboard.down('Shift');
    await page.keyboard.press('Tab');
    await page.keyboard.up('Shift');
    assert.equal(await page.evaluate(() => document.activeElement.textContent), 'Save');
    // Outside clicks cannot dismiss or alter the underlying form.
    await page.mouse.click(8, 8);
    await prompt();
    await page.keyboard.press('Escape');
    await noPrompt();
    assert.equal(await page.$eval('#start-with-windows', el => el.getAttribute('aria-checked')), 'true');
    await page.keyboard.press('Escape');
    await prompt();
    await page.screenshot({path: '/tmp/apimonitor-unsaved-changes.png'});
    await click('Discard', true);
    await page.waitForFunction(() => window.messages.some(m => m.action === 'close'));
    assert.deepEqual(await actions(), [{action: 'close'}]);

    // Save uses the same payload as the normal Save button and never discards.
    for (const throughPrompt of [false, true]) {
      await reset();
      await page.click('#start-with-windows');
      await changeText('#history-limit', '250');
      if (throughPrompt) { await nativeClose(); await prompt(); }
      await click('Save', throughPrompt);
      await page.waitForFunction(() => window.messages.some(m => m.action === 'saveSettings'));
      const {updateCheckPending, updatePromptPending, ...expected} = config;
      assert.deepEqual(await actions(), [{action: 'saveSettings', ...expected,
        startWithWindows: true, historyLimit: 250}]);
    }

    // Empty URL cannot silently lose edits on Save.
    await reset();
    await changeText('#api-url', '');
    await nativeClose();
    await prompt();
    await click('Save', true);
    await noPrompt();
    assert.deepEqual(await actions(), []);
    assert.equal(await page.$eval('#save-error', el => el.textContent), 'Enter an API URL.');
    await page.waitForFunction(() => document.activeElement.id === 'api-url');

    // An update arriving while deciding cannot replace the unsaved-changes question.
    await reset();
    await page.click('#logging');
    await nativeClose();
    await prompt();
    const saveOverlay = await page.$eval('[role="alertdialog"]', el => getComputedStyle(el.parentElement).backgroundColor);
    await page.evaluate(() => window.onUpdateResult({status:'newer', title:'Update available',
      message:'A newer version is ready to install.', currentVersion:'1.0.13', remoteVersion:'1.0.14', automatic:true}));
    await prompt();
    await click('Keep editing', true);
    await page.waitForSelector('#update-alert-title');
    assert.equal(await page.$eval('[role="alertdialog"]', el => getComputedStyle(el.parentElement).backgroundColor), saveOverlay);
    assert.equal(saveOverlay, 'rgba(0, 0, 0, 0.35)');
    await nativeClose();
    await prompt();
    await page.keyboard.press('Escape');
    await page.waitForSelector('#update-alert-title');
    await click('Cancel', true);
    await page.waitForFunction(() => !document.querySelector('[role="alertdialog"]'));
    assert.equal(await page.$eval('#logging', el => el.getAttribute('aria-checked')), 'true');
    assert.deepEqual(await actions(), []);
    assert.deepEqual(errors, []);
    console.log('Unsaved settings, close routes, save/discard, keyboard, overlay, and update overlap checks passed');
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
