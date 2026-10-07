const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../content.js'), 'utf8');
const functions = ['translateText', 'translateHTMLContent', 'translateWithGoogleFree'].map(name => {
  const start = source.indexOf(`async function ${name}(`);
  assert.ok(start >= 0);
  return source.slice(start, source.indexOf('\n}', start) + 2);
}).join('\n');
(async () => {
  const browser = await chromium.launch({headless:true, executablePath:process.env.CHROME_PATH});
  try {
    const page = await browser.newPage();
    await page.addScriptTag({content:functions});
    for (const target of ['en', 'zh-TW']) {
      const result = await page.evaluate(async target => {
        const container = document.createElement('div');
        container.innerHTML = '<p>第一段</p><p>第二段 <a href="https://e3.nycu.edu.tw/file">附件</a> 結尾</p><p>第一段</p><code>保留程式</code><style>.unchanged{color:red}</style>';
        const original = container.innerHTML;
        const requests = [];
        const translated = {'第一段':'Paragraph one','第二段':'Paragraph two','附件':'Attachment','結尾':'Ending'};
        window.fetch = async url => {
          const q = new URL(url).searchParams.get('q');
          requests.push(q);
          const text = q.replace(/第一段|第二段|附件|結尾/g, match => translated[match])
            .replace(/\n<<<SEPARATOR>>>\n/g, target === 'en' ? ' <<<separator>>> ' : ' <<<分隔符號>>> ');
          return {ok:true,json:async()=>[[[text,q]]]};
        };
        const html = await translateHTMLContent(container, 'auto', target);
        const output = document.createElement('div');
        output.innerHTML = html;
        return {html,text:output.textContent,paragraphs:[...output.querySelectorAll('p')].map(p=>p.textContent),href:output.querySelector('a').getAttribute('href'),code:output.querySelector('code').textContent,style:output.querySelector('style').textContent,originalUnchanged:container.innerHTML===original,requests};
      }, target);
      assert.ok(!/separator|分隔符號/i.test(result.text), `${target}: translation must not expose internal separators`);
      assert.deepEqual(result.paragraphs,['Paragraph one','Paragraph two Attachment Ending','Paragraph one']);
      assert.equal(result.href,'https://e3.nycu.edu.tw/file');
      assert.equal(result.code,'保留程式');
      assert.equal(result.style,'.unchanged{color:red}');
      assert.equal(result.originalUnchanged,true);
      assert.equal(result.requests.filter(q=>q==='第一段').length,1);
    }
    const failure = await page.evaluate(async () => {
      const container = document.createElement('div');
      container.innerHTML = '<p>First</p><p>Second</p>';
      const original = container.innerHTML;
      window.fetch = async url => new URL(url).searchParams.get('q') === 'Second'
        ? {ok:false,status:429} : {ok:true,json:async()=>[[['Translated']]]};
      try {
        await translateHTMLContent(container,'en','zh-TW');
        return {rejected:false};
      } catch {
        return {rejected:true,unchanged:container.innerHTML===original};
      }
    });
    assert.deepEqual(failure,{rejected:true,unchanged:true});
    const blank = await page.evaluate(async () => {
      const container = document.createElement('div');
      container.innerHTML = '<p>   <br></p>';
      window.fetch = async () => {throw new Error('Blank content must not request translation');};
      return await translateHTMLContent(container,'en','zh-TW') === container.innerHTML;
    });
    assert.equal(blank,true);
    console.log('Passed: translations contain no internal separators, preserve paragraphs/inline spacing/links/excluded elements and original HTML, and reuse duplicate text.');
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
