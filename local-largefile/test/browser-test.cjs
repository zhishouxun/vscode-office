const { chromium } = require('playwright');
const http = require('http');
const fs = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const root=path.resolve(process.env.OFFICE_EXTENSION_ROOT || 'build');
const fixture=(process.env.MARKDOWN_TEST_FILE ? fs.readFileSync(process.env.MARKDOWN_TEST_FILE,'utf8') : '# 资源目录\n\n| 名称 | 链接 | 验证状态 | 备注 |\n|---|---|---|---|\n'+Array.from({length:1600},(_,i)=>`| 小说${i} | [打开](https://example.com/${i}) | 有效 | 示例 |\n`).join(''));
const server=http.createServer((req,res)=>{
 let pathname=decodeURIComponent(new URL(req.url,'http://localhost').pathname);
 const filename=path.join(root,pathname);
 if(!filename.startsWith(root+path.sep)){res.writeHead(403);res.end();return;}
 try{
  let data=fs.readFileSync(filename);
  if(pathname.endsWith('index.html'))data=Buffer.from(data.toString().replace('{{baseUrl}}',`http://127.0.0.1:${server.address().port}/resource/markdown`));
  const mime={'.html':'text/html','.js':'text/javascript','.mjs':'text/javascript','.css':'text/css','.svg':'image/svg+xml','.woff2':'font/woff2','.wasm':'application/wasm'}[path.extname(filename)]||'application/octet-stream';
  res.writeHead(200,{'Content-Type':mime});res.end(data);
 }catch{res.writeHead(404);res.end();}
});
(async()=>{
 await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
 const origin=`http://127.0.0.1:${server.address().port}`;
 const browser=await chromium.launch({executablePath:process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',headless:true});
 const page=await browser.newPage({viewport:{width:1440,height:950}});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 page.on('console',m=>{if(m.type()==='error')console.log('BROWSER',m.text().slice(0,200));});
 await page.addInitScript(({origin,fixture})=>{
  window.__messages=[];
  window.acquireVsCodeApi=()=>({postMessage(m){
   window.__messages.push(m);
   if(m.type==='init')setTimeout(()=>window.dispatchEvent(new MessageEvent('message',{data:{type:'open',content:{content:fixture,rootPath:origin+'/resource/markdown',workspaceBaseUrl:'',documentCacheId:'large-file-test',fileName:'有声书',config:{language:'zh-cn',isWeb:false,isDev:false,isPro:false,isProCancelled:false,editMode:'wysiwyg',editorTheme:'Auto',codeMirrorTheme:'Auto',mermaidTheme:'Auto',preventMacOptionKey:true},viewerSettings:{enabled:false}}}})),0);
   if(m.type==='queryAIAvailable')setTimeout(()=>window.dispatchEvent(new MessageEvent('message',{data:{type:'aiAvailable',content:false}})),0);
  },getState(){return{}},setState(){}});
 },{origin,fixture});
 const started=Date.now();await page.goto(origin+'/resource/markdown/index.html');
 await page.waitForSelector('.vditor-wysiwyg [contenteditable="true"], .vditor-wysiwyg[contenteditable="true"]',{timeout:30000});
 await page.waitForFunction(()=>document.querySelectorAll('.vditor-wysiwyg tbody tr').length>0,{timeout:30000});
 const firstRenderMs=Date.now()-started;
 const before=await page.locator('.vditor-wysiwyg tbody tr').count();assert.ok(before>80&&before<=320,`rendered ${before} rows`);
 const totalPages=await page.getByRole('combobox',{name:'选择页码'}).locator('option').count();assert.ok(totalPages>1);
 // Visiting another page must not normalize or rewrite the document.
 await page.getByRole('button',{name:'下一页',exact:true}).click();
 await page.waitForFunction(()=>document.querySelector('select[aria-label="选择页码"]').value==='1');
 const unchanged=await page.evaluate(()=>window.__messages.filter(m=>m.type==='save').at(-1)?.content);
 assert.equal(unchanged,fixture,'Navigation must preserve the complete original source');
 const headerLocked=await page.locator('.vditor-wysiwyg table thead').getAttribute('contenteditable');assert.equal(headerLocked,'false');
 // Edit a real resource name in page 2, then save with the actual keyboard shortcut.
 const cell=page.locator('.vditor-wysiwyg tbody tr').first().locator('td').first();
 const oldText=await cell.innerText();await cell.click();await page.keyboard.press('Home');await page.keyboard.type('本地测试');
 await page.waitForFunction(()=>window.__messages.some(m=>m.type==='save'&&m.content.includes('本地测试')),{timeout:10000});
 await page.keyboard.press('Meta+s');
 await page.waitForFunction(()=>window.__messages.some(m=>m.type==='doSave'),{timeout:10000});
 const saved=await page.evaluate(()=>window.__messages.filter(m=>m.type==='doSave').at(-1).content);
 fs.writeFileSync('browser-saved-fixture.md',saved);
 assert.ok(saved.includes('本地测试'));assert.ok(saved.length>fixture.length*.95,'Save must contain the entire document');
 assert.equal((saved.match(/\| 名称 \| 链接 \| 验证状态 \| 备注 \|/g)||[]).length,(fixture.match(/\| 名称 \| 链接 \| 验证状态 \| 备注 \|/g)||[]).length);
 // Cross-page search must reach the last page and keep the change.
 const lastName=fixture.split('\n').filter(s=>s.startsWith('| ')).at(-1).split('|')[1].trim();
 await page.getByRole('searchbox',{name:'搜索整篇文件'}).fill(lastName);
 await page.getByRole('button',{name:'查找',exact:true}).click();
 await page.waitForFunction(last=>document.querySelector('select[aria-label="选择页码"]').value===String(last),totalPages-1,{timeout:10000});
 const lastSave=await page.evaluate(()=>window.__messages.filter(m=>m.type==='save').at(-1).content);assert.ok(lastSave.includes('本地测试'));
 const navRect=await page.locator('.office-large-document-bar').boundingBox();assert.ok(navRect.y>=0&&navRect.y<10,'Paging toolbar must remain visible after search');
 await page.screenshot({path:'office-largefile-preview.png'});
 assert.deepEqual(errors,[]);
 console.log(JSON.stringify({firstRenderMs,firstPageRows:before,totalPages,fullDocumentSave:true,preserveUneditedText:true,tableHeaderNotDuplicated:true,fullDocumentSearch:true,browserErrors:errors,editedResource:oldText},null,2));
 await browser.close();server.close();
})().catch(e=>{console.error(e);server.close();process.exitCode=1;setTimeout(()=>process.exit(1),1000)});
