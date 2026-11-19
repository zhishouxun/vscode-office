const { chromium } = require('playwright');
const http = require('http');
const fs = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const root=path.resolve(process.env.OFFICE_EXTENSION_ROOT || 'build');
const fixture='# 导航测试\n\n'+Array.from({length:1000},(_,i)=>`段落${String(i).padStart(4,'0')}：同词，第一个；同词，第二个。\n\n`).join('')+'\n[隐藏匹配](https://example.com/隐藏匹配)\n';
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
   if(m.type==='init')setTimeout(()=>window.dispatchEvent(new MessageEvent('message',{data:{type:'open',content:{content:fixture,rootPath:origin+'/resource/markdown',workspaceBaseUrl:'',documentCacheId:localStorage.getItem('test.documentId')||'navigation-A',fileName:'有声书',config:{language:'zh-cn',isWeb:false,isDev:false,isPro:false,isProCancelled:false,editMode:'wysiwyg',editorTheme:'Auto',codeMirrorTheme:'Auto',mermaidTheme:'Auto',preventMacOptionKey:true},shouldRestoreFocus:true,viewerSettings:{enabled:false}}}})),0);
   if(m.type==='queryAIAvailable')setTimeout(()=>window.dispatchEvent(new MessageEvent('message',{data:{type:'aiAvailable',content:false}})),0);
  },getState(){return{}},setState(){}});
 },{origin,fixture});
 await page.goto(origin+'/resource/markdown/index.html');
 await page.waitForSelector('.office-large-document-bar[aria-busy="false"]');
 const query=page.getByRole('searchbox',{name:'搜索整篇文件'});
 const currentPage=()=>page.getByRole('combobox',{name:'选择页码'}).inputValue();
 const selection=()=>page.evaluate(()=>{const s=window.getSelection();return {text:s.toString(),node:s.anchorNode?.textContent,offset:s.anchorOffset};});
 await query.fill('同词');await query.press('Enter');
 const first=await selection();assert.equal(first.text,'同词');
 await query.press('Enter');const second=await selection();
 assert.equal(second.text,'同词');assert.equal(first.node,second.node);assert.ok(second.offset>first.offset);
 assert.match(await page.locator('.office-large-document-status').innerText(),/第 2 \/ 2000 处/);
 await query.press('Enter');const third=await selection();assert.notEqual(third.node,second.node);
 await query.press('Shift+Enter');assert.deepEqual(await selection(),second);
 // Reading at a unique paragraph restores its page and pixel position on reload.
 await query.fill('段落0500');await query.press('Enter');
 const beforePage=await currentPage();assert.notEqual(beforePage,'0');
 const geometry=()=>page.evaluate(()=>{
  const root=document.querySelector('.vditor-wysiwyg [contenteditable="true"], .vditor-wysiwyg[contenteditable="true"]');
  const target=[...root.querySelectorAll('p')].find(p=>p.textContent.startsWith('段落0500'));
  return target?.getBoundingClientRect().top;
 });
 const beforeY=await geometry();
 await page.waitForFunction(()=>localStorage.getItem('office.largeDocument.position:navigation-A'));
 // Flush the final scroll event using pagehide, just as a closing webview would.
 await page.evaluate(()=>window.dispatchEvent(new Event('pagehide')));
 const savedPosition=await page.evaluate(()=>JSON.parse(localStorage.getItem('office.largeDocument.position:navigation-A')));
 assert.ok(savedPosition.anchor.offset>0);
 await page.reload();await page.waitForSelector('.office-large-document-bar[aria-busy="false"]');
 await page.waitForFunction(()=>document.querySelector('select[aria-label="选择页码"]').value!== '0');
 await page.waitForFunction(y=>{const p=[...document.querySelectorAll('.vditor-wysiwyg p')].find(p=>p.textContent.startsWith('段落0500'));return p&&Math.abs(p.getBoundingClientRect().top-y)<12;},beforeY);
 assert.equal(await currentPage(),beforePage);
 // Resizing uses a source anchor instead of returning to the larger page's beginning.
 await page.getByRole('combobox',{name:'每页容量'}).selectOption('640');
 assert.ok(Math.abs(await geometry()-beforeY)<12);
 await page.keyboard.press('Meta+s');
 await page.waitForFunction(()=>window.__messages.some(m=>m.type==='doSave'));
 assert.equal(await page.evaluate(()=>window.__messages.filter(m=>m.type==='doSave').at(-1).content),fixture);
 // A collapsed caret restores without moving the reading viewport.
 await page.evaluate(()=>{
  const root=document.querySelector('.vditor-wysiwyg [contenteditable="true"], .vditor-wysiwyg[contenteditable="true"]');
  const paragraph=[...root.querySelectorAll('p')].find(p=>p.textContent.startsWith('段落0500'));
  root.focus({preventScroll:true});const range=document.createRange();range.setStart(paragraph.firstChild,12);range.collapse(true);
  const selection=window.getSelection();selection.removeAllRanges();selection.addRange(range);
  window.dispatchEvent(new Event('pagehide'));
 });
 await page.reload();await page.waitForSelector('.office-large-document-bar[aria-busy="false"]');
 await page.waitForFunction(()=>{const s=window.getSelection();return s?.isCollapsed&&s.anchorNode?.textContent.startsWith('段落0500')&&s.anchorOffset===12;});
 assert.ok(Math.abs(await geometry()-beforeY)<12);
 // External insertions relocate the same source anchor rather than retaining a stale page number.
 const external='# 新插入标题\n\n'+fixture;
 await page.evaluate(content=>window.dispatchEvent(new MessageEvent('message',{data:{type:'update',content}})),external);
 assert.ok(Math.abs(await geometry()-beforeY)<12);
 await page.keyboard.press('Meta+s');
 await page.waitForFunction(()=>window.__messages.some(m=>m.type==='doSave'));
 assert.equal(await page.evaluate(()=>window.__messages.filter(m=>m.type==='doSave').at(-1).content),external);
 // A separate file must not inherit the previous file's page or caret bookmark.
 await page.evaluate(()=>localStorage.setItem('test.documentId','navigation-B'));
 await page.reload();await page.waitForSelector('.office-large-document-bar[aria-busy="false"]');
 assert.equal(await currentPage(),'0');
 // Source-only occurrences must not highlight the wrong visible occurrence.
 await page.getByRole('searchbox',{name:'搜索整篇文件'}).fill('隐藏匹配');
 await page.getByRole('searchbox',{name:'搜索整篇文件'}).press('Enter');
 assert.match(await page.locator('.office-large-document-status').innerText(),/无法在预览中准确高亮/);
 assert.equal((await selection()).text,'');
 assert.deepEqual(errors,[]);
 console.log('PASS: repeated search positions, match count, reading restore, resize anchoring, per-file isolation and complete save');
 await browser.close();server.close();
})().catch(e=>{console.error(e);server.close();process.exitCode=1;setTimeout(()=>process.exit(1),1000)});
