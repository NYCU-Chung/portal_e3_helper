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
        const data=JSON.parse(localStorage.getItem('fixtureStorage') || 'null') || {interfaceLanguage:language,themePreference:'light',aiSettings:{enabled:false,geminiApiKey:'fixture-gemini',geminiModel:'saved-model',openaiSummaryApiKey:'fixture-openai',openaiSummaryModel:'gpt-5-nano'},assignments:[{eventId:'manual-1',name:'作業中文原文',course:'課程中文原文',deadline:Date.now()+86400000,manualStatus:'pending',isManual:true}],courses:[],announcements:[],lastSyncTime:Date.now(),lastSeenVersion:'2.2.0'};
        window.chrome={i18n:{getUILanguage:()=>language},runtime:{id:'test',getManifest:()=>({version:'2.2.0'}),getURL:p=>p,onMessage:{addListener(){}},sendMessage:(m,cb)=>{window.fixtureRequests ||= [];window.fixtureRequests.push(m);const response=m.action==='listGeminiModels'?(window.fixtureModelFailure?{success:false,error:'HTTP 502'}:{success:true,data:[{id:'future-flash',name:'Future Flash'},{id:'future-pro',name:'Future Pro'}]}):{success:true,data:'fixture summary'};cb?.(response);return Promise.resolve(response);}},storage:{onChanged:{addListener:fn=>listeners.push(fn)},local:{get:async(keys,cb)=>{cb?.(data);return data;},set:async update=>{Object.assign(data,update);localStorage.setItem('fixtureStorage',JSON.stringify(data));listeners.forEach(fn=>fn(Object.fromEntries(Object.entries(update).map(([k,v])=>[k,{newValue:v}])),'local'));}}}};
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
      await page.locator('#e3-helper-enable-ai').check();
      assert.equal(await page.locator('#e3-helper-ai-provider').inputValue(),'gemini');
      assert.equal(await page.locator('#e3-helper-gemini-model-id').inputValue(),'saved-model');
      assert.ok(await page.locator('#e3-helper-gemini-model option[value="future-flash"]').count());
      assert.ok(await page.locator('#e3-helper-openai-settings').isHidden());
      await page.locator('#e3-helper-test-ai-btn').click();
      await page.waitForFunction(() => document.querySelector('#e3-helper-ai-status').textContent.includes('連接成功'));
      assert.equal(await page.evaluate(() => window.fixtureRequests.at(-1).action),'callGeminiApi');
      await page.locator('#e3-helper-gemini-model').selectOption('future-pro');
      assert.equal(await page.locator('#e3-helper-gemini-model-id').inputValue(),'future-pro');
      await page.evaluate(() => {window.fixtureModelFailure=true;});
      await page.locator('#e3-helper-refresh-gemini-models').click();
      await page.waitForFunction(() => document.querySelector('#e3-helper-gemini-model-status').textContent.includes('HTTP 502'));
      assert.ok(await page.locator('#e3-helper-gemini-model option[value="future-flash"]').count());
      await page.locator('#e3-helper-gemini-model-id').fill('manual-future-model');
      await page.locator('#e3-helper-save-settings').click();
      assert.equal(await page.evaluate(async () => (await chrome.storage.local.get(['aiSettings'])).aiSettings.geminiModel),'manual-future-model');
      await page.locator('#e3-helper-more-btn').click();
      await page.locator('#e3-helper-settings-btn').click();
      assert.equal(await page.locator('#e3-helper-gemini-model-id').inputValue(),'manual-future-model');
      assert.ok(await page.locator('#e3-helper-gemini-model option[value="future-flash"]').count());
      await page.locator('#e3-helper-ai-provider').selectOption('openai');
      assert.ok(await page.locator('#e3-helper-gemini-settings').isHidden());
      assert.ok(await page.locator('#e3-helper-openai-settings').isVisible());
      assert.equal(await page.locator('#e3-helper-openai-summary-key').inputValue(),'fixture-openai');
      await page.locator('#e3-helper-test-ai-btn').click();
      await page.waitForFunction(() => document.querySelector('#e3-helper-ai-status').textContent.includes('連接成功'));
      assert.equal(await page.evaluate(() => window.fixtureRequests.at(-1).action),'callOpenAIResponsesApi');
      await page.locator('#e3-helper-theme').selectOption('dark');
      await page.locator('#e3-helper-save-settings').click();
      assert.equal(await page.locator('html').getAttribute('data-e3-helper-theme'),'dark');
      const savedAI = await page.evaluate(async () => (await chrome.storage.local.get(['aiSettings'])).aiSettings);
      assert.equal(savedAI.summaryProvider,'openai');
      assert.equal(savedAI.geminiApiKey,'fixture-gemini');
      assert.equal(savedAI.geminiModel,'manual-future-model');
      assert.equal(savedAI.openaiSummaryApiKey,'fixture-openai');

      const discoveryCount = await page.evaluate(() => window.fixtureRequests.filter(request => request.action === 'listGeminiModels').length);
      await page.locator('#e3-helper-more-btn').click();
      await page.locator('#e3-helper-settings-btn').click();
      assert.equal(await page.locator('#e3-helper-ai-provider').inputValue(),'openai');
      assert.equal(await page.evaluate(() => window.fixtureRequests.filter(request => request.action === 'listGeminiModels').length),discoveryCount);
      await page.locator('#e3-helper-save-settings').click();
      await page.evaluate(async () => {
        const {aiSettings} = await chrome.storage.local.get(['aiSettings']);
        await chrome.storage.local.set({aiSettings:{...aiSettings,geminiModel:''}});
        window.fixtureModelFailure=false;
      });
      await page.locator('#e3-helper-more-btn').click();
      await page.locator('#e3-helper-settings-btn').click();
      assert.equal(await page.evaluate(() => window.fixtureRequests.filter(request => request.action === 'listGeminiModels').length),discoveryCount);
      assert.equal(await page.locator('#e3-helper-gemini-model').inputValue(),'');
      assert.equal(await page.locator('#e3-helper-gemini-model-id').inputValue(),'');
      await page.locator('#e3-helper-ai-provider').selectOption('gemini');
      await page.waitForFunction(() => document.querySelector('#e3-helper-gemini-model-status').textContent === '模型清單已更新。');
      assert.equal(await page.evaluate(() => window.fixtureRequests.filter(request => request.action === 'listGeminiModels').length),discoveryCount+1);
      assert.equal(await page.locator('#e3-helper-gemini-model').inputValue(),'');
      assert.equal(await page.locator('#e3-helper-gemini-model-id').inputValue(),'');
      const requestCount = await page.evaluate(() => window.fixtureRequests.length);
      await page.locator('#e3-helper-test-ai-btn').click();
      assert.equal(await page.evaluate(() => window.fixtureRequests.length),requestCount);
      assert.ok((await page.locator('#e3-helper-ai-status').innerText()).includes('請輸入 API Key 並選擇模型'));
      await page.locator('#e3-helper-gemini-model').selectOption('future-flash');
      assert.equal(await page.locator('#e3-helper-gemini-model-id').inputValue(),'future-flash');
      await page.locator('#e3-helper-save-settings').click();

      assert.equal(await page.evaluate(async () => (await chrome.storage.local.get(['themePreference'])).themePreference),'dark');
      await page.reload();
      if(localized) await page.addScriptTag({path:path.join(root,'i18n.js')});
      await page.addScriptTag({path:path.join(root,'content.js')});
      assert.equal(await page.locator('html').getAttribute('data-e3-helper-theme'),'dark');
      assert.equal(await page.locator('.e3-helper-sidebar').evaluate(el => getComputedStyle(el).backgroundColor),'rgb(18, 17, 16)');
      const script = fs.readFileSync(path.join(root,'content.js'),'utf8');
      const functions = ['helperIcon', 'escapeHtml', 'showTemporaryMessage'].map(name => {
        const start = script.indexOf(`function ${name}(`);
        return script.slice(start,script.indexOf('\n}',start)+2);
      }).join('\n');
      await page.addScriptTag({content:functions});
      for (const theme of ['light', 'dark']) {
        await page.evaluate(theme => chrome.storage.local.set({themePreference:theme}), theme);
        const treatments = [];
        for (const type of ['success', 'error', 'warning', 'info']) {
          await page.evaluate(type => showTemporaryMessage('Status fixture', type, 60000), type);
          const toast = page.locator('.e3-helper-toast').last();
          assert.equal(await toast.getAttribute('role'), type === 'error' ? 'alert' : 'status');
          const colors = await toast.evaluate(el => {
            const style = getComputedStyle(el);
            return {border:style.borderLeftColor,width:style.borderLeftWidth,bg:style.backgroundColor,text:style.color};
          });
          assert.equal(colors.width,'4px');
          assert.notEqual(colors.border,colors.bg);
          assert.notEqual(colors.text,colors.bg);
          treatments.push(colors.border);
        }
        assert.equal(new Set(treatments).size,4,`${theme} has distinct toast status colors`);
        await page.locator('.e3-helper-toast').evaluateAll(elements => elements.forEach(el => el.remove()));
      }
      assert.deepEqual(errors,[]);
      await page.close();
    }
    console.log('Passed: six tabs, responsive layout, original course text, settings, restored dark theme, toast status colors, provider switching, OpenAI settings without Gemini requests, and explicit Gemini model selection/discovery/fallback.');
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
