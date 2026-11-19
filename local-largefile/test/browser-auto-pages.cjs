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
 await page.goto(origin+'/resource/markdown/index.html');
 await page.waitForSelector('.vditor-wysiwyg tbody tr');await page.waitForSelector('.office-large-document-bar[aria-busy="false"]');
 const pageNumber=()=>page.locator('select[aria-label="选择页码"]').inputValue();
 const count=await page.locator('.vditor-wysiwyg tbody tr').count();
 assert.ok(count>80&&count<=320,`default page has ${count} rows`);
 const pageCount=await page.locator('select[aria-label="选择页码"] option').count();
 const edge=async bottom=>page.evaluate(bottom=>{
  const root=document.querySelector('.vditor-wysiwyg [contenteditable="true"], .vditor-wysiwyg[contenteditable="true"]');
  root.focus({preventScroll:true});
  const range=document.createRange();range.selectNodeContents(root);range.collapse(!bottom);
  const selection=window.getSelection();selection.removeAllRanges();selection.addRange(range);
  let scroller=root;
  while(scroller!==document.body){
   if(/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)&&scroller.scrollHeight>scroller.clientHeight+2)break;
   scroller=scroller.parentElement;
  }
  scroller.scrollTop=bottom?scroller.scrollHeight:0;
  const r=scroller.getBoundingClientRect();
  return {x:r.right-50,y:bottom?r.bottom-40:r.top+50,top:scroller.scrollTop,height:scroller.clientHeight};
 },bottom);
 let point=await edge(true);await page.mouse.move(point.x,point.y);await page.mouse.wheel(0,160);
 await page.waitForFunction(()=>document.querySelector('select[aria-label="选择页码"]').value==='1');
 // The remainder of one trackpad gesture must not skip multiple pages.
 await page.evaluate(()=>{
  const root=document.querySelector('.vditor-wysiwyg [contenteditable="true"], .vditor-wysiwyg[contenteditable="true"]');
  root.scrollTop=root.scrollHeight;
  for(let i=0;i<10;i++)root.dispatchEvent(new WheelEvent('wheel',{bubbles:true,cancelable:true,deltaY:20000}));
 });
 assert.equal(await pageNumber(),'1');
 // A fresh upward gesture at the start returns to the previous page's bottom.
 await page.waitForTimeout(550);
 point=await edge(false);await page.mouse.move(point.x,point.y);await page.mouse.wheel(0,-160);
 await page.waitForFunction(()=>document.querySelector('select[aria-label="选择页码"]').value==='0');
 const previousEdge=await page.evaluate(()=>{
  let n=document.querySelector('.vditor-wysiwyg [contenteditable="true"], .vditor-wysiwyg[contenteditable="true"]');
  while(n!==document.body){if(/(auto|scroll)/.test(getComputedStyle(n).overflowY)&&n.scrollHeight>n.clientHeight+2)return n.scrollTop+n.clientHeight>=n.scrollHeight-3;n=n.parentElement;}
  return false;
 });assert.equal(previousEdge,true,'backwards paging must land at the bottom');
 await edge(true);await page.keyboard.press('PageDown');
 await page.waitForFunction(()=>document.querySelector('select[aria-label="选择页码"]').value==='1');
 await edge(false);await page.keyboard.press('PageUp');
 await page.waitForFunction(()=>document.querySelector('select[aria-label="选择页码"]').value==='0');
 // Turning off automatic paging restores ordinary scrolling.
 await page.getByRole('checkbox',{name:'自动翻页'}).uncheck();
 point=await edge(true);await page.mouse.move(point.x,point.y);await page.mouse.wheel(0,200);
 assert.equal(await pageNumber(),'0');
 await page.getByRole('checkbox',{name:'自动翻页'}).check();
 // Edit before resizing: the full source must retain the change after repartitioning.
 const cell=page.locator('.vditor-wysiwyg tbody tr').first().locator('td').first();
 await cell.click();await page.keyboard.press('Home');await page.keyboard.type('容量测试');
 await page.waitForFunction(()=>window.__messages.some(m=>m.type==='save'&&m.content.includes('容量测试')));
 const expectedSource=await page.evaluate(()=>window.__messages.filter(m=>m.type==='save').at(-1).content);
 await page.getByRole('combobox',{name:'每页容量'}).selectOption('640');
 assert.ok(await page.locator('select[aria-label="选择页码"] option').count()<pageCount);
 await page.keyboard.press('Meta+s');
 await page.waitForFunction(()=>window.__messages.some(m=>m.type==='doSave'&&m.content.includes('容量测试')));
 const saved=await page.evaluate(()=>window.__messages.filter(m=>m.type==='doSave').at(-1).content);
 assert.ok(saved.includes('小说1599'),'last page must remain in saved source');
 assert.equal(saved,expectedSource,'changing capacity must preserve the entire edited source');
 assert.equal((saved.match(/\|\s*名称\s*\|\s*链接\s*\|\s*验证状态\s*\|\s*备注\s*\|/g)||[]).length,1);
 await page.getByRole('checkbox',{name:'自动翻页'}).uncheck();
 await page.reload();await page.waitForSelector('.vditor-wysiwyg tbody tr');await page.waitForSelector('.office-large-document-bar[aria-busy="false"]');
 assert.equal(await page.getByRole('combobox',{name:'每页容量'}).inputValue(),'640');
 assert.equal(await page.getByRole('checkbox',{name:'自动翻页'}).isChecked(),false);
 assert.deepEqual(errors,[]);
 console.log(JSON.stringify({defaultPageRows:count,defaultPages:pageCount,wheelForward:true,wheelBack:true,momentumGuard:true,keyboardPaging:true,manualToggle:true,resizingPreservesEdits:true,preferencesPersist:true}));
 await browser.close();server.close();
})().catch(e=>{console.error(e);server.close();process.exitCode=1;setTimeout(()=>process.exit(1),1000)});
