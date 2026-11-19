const { chromium } = require('playwright');
const http = require('http');
const fs = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const root=path.resolve(process.env.OFFICE_EXTENSION_ROOT || 'build');
const fixture='# 测试书单\n\n- [ ] 第一本\n- [x] 第二本\n';
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
 assert.equal(await page.locator('.office-large-document-bar').count(),0);
 await page.keyboard.press('Meta+s');
 await page.waitForFunction(()=>window.__messages.some(m=>m.type==='doSave'));
 assert.equal(await page.evaluate(()=>window.__messages.filter(m=>m.type==='doSave').at(-1).content),fixture);
 const editable=page.locator('.vditor-wysiwyg [contenteditable="true"], .vditor-wysiwyg[contenteditable="true"]').first();
 await editable.click();await page.keyboard.press('End');await page.keyboard.type(' 新增文字');
 await page.waitForFunction(()=>window.__messages.some(m=>m.type==='save'&&m.content.includes('新增文字')));
 await page.keyboard.press('Meta+s');
 await page.waitForFunction(()=>window.__messages.filter(m=>m.type==='doSave').at(-1)?.content.includes('新增文字'));
 assert.deepEqual(errors,[]);
 console.log('PASS: small file stays unpaginated, unchanged save preserves source, edit/save works');
 await browser.close();server.close();
})().catch(e=>{console.error(e);server.close();process.exitCode=1;setTimeout(()=>process.exit(1),1000)});
