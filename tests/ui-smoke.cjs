const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright');
const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs');
(async () => {
  const browser = await chromium.launch({headless:true, executablePath:process.env.CHROME_PATH});
  try {
    const root = path.join(__dirname, '..');
    const localized = fs.existsSync(path.join(root,'i18n.js'));
    for (const language of localized ? ['zh-TW', 'en'] : ['zh-TW']) {
      const page = await browser.newPage({viewport:{width:1280,height:900}});
      const errors=[];
      page.on('pageerror', e => errors.push(e.message));
      await page.route('http://e3-test.local/**', route => route.fulfill({contentType:'text/html',body:'<!doctype html><html><head></head><body></body></html>'}));
      await page.addInitScript(language => {
        const listeners=[];
        const data={interfaceLanguage:language,themePreference:'light',assignments:[{eventId:'manual-1',name:'作業中文原文',course:'課程中文原文',deadline:Date.now()+86400000,manualStatus:'pending',isManual:true}],courses:[],announcements:[],lastSyncTime:Date.now(),lastSeenVersion:'2.2.0'};
        window.chrome={i18n:{getUILanguage:()=>language},runtime:{id:'test',getManifest:()=>({version:'2.2.0'}),getURL:p=>p,onMessage:{addListener(){}},sendMessage:(m,cb)=>{cb?.({success:true});return Promise.resolve({success:true});}},storage:{onChanged:{addListener:fn=>listeners.push(fn)},local:{get:async(keys,cb)=>{cb?.(data);return data;},set:async update=>{Object.assign(data,update);listeners.forEach(fn=>fn(Object.fromEntries(Object.entries(update).map(([k,v])=>[k,{newValue:v}])),'local'));}}}};
      },language);
      await page.goto('http://e3-test.local/');
      if(localized) await page.addScriptTag({path:path.join(root,'i18n.js')});
      await page.addScriptTag({path:path.join(root,'content.js')});
      await page.locator('.e3-helper-sidebar-toggle').click();
      for(const width of [280,350,480,800]) {
        const result=await page.evaluate(width=>{
          const sidebar=document.querySelector('.e3-helper-sidebar');sidebar.style.width=width+'px';
          const tabs=sidebar.querySelector('.e3-helper-tabs'), bounds=tabs.getBoundingClientRect();
          return tabs.scrollWidth<=tabs.clientWidth && [...tabs.children].every(tab=>{const r=tab.getBoundingClientRect();return r.left>=bounds.left && r.right<=bounds.right && r.top>=bounds.top && r.bottom<=bounds.bottom;});
        },width);
        assert.ok(result,`Tabs fit ${width}px`);
      }
      for(const tab of ['assignments','grades','downloads','announcements','notifications','help']) {
        await page.locator(`[data-tab="${tab}"]`).click();
        assert.ok((await page.locator(`[data-content="${tab}"]`).innerText()).trim(),`${tab} renders`);
      }
      assert.equal(await page.locator('.e3-helper-help a').filter({hasText:'GitHub'}).getAttribute('href'),'https://github.com/NYCU-Chung/portal_e3_helper');
      await page.locator('[data-tab="assignments"]').click();
      assert.ok((await page.locator('[data-content="assignments"]').innerText()).includes('作業中文原文'));
      await page.locator('#e3-helper-more-btn').click();
      await page.locator('#e3-helper-settings-btn').click();
      assert.equal(await page.locator('#e3-helper-theme').inputValue(),'light');
      await page.locator('#e3-helper-theme').selectOption('dark');
      await page.locator('#e3-helper-save-settings').click();
      assert.equal(await page.locator('html').getAttribute('data-e3-helper-theme'),'dark');
      assert.equal(await page.evaluate(async () => (await chrome.storage.local.get(['themePreference'])).themePreference),'dark');
      assert.deepEqual(errors,[]);
      await page.close();
    }
    console.log('Passed: six tabs, responsive layout, original course text, settings and saved theme.');
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
