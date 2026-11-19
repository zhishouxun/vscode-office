const { chromium } = require('playwright');
const http = require('http');
const fs = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const root=path.resolve(process.env.OFFICE_EXTENSION_ROOT || 'build');
const fixture='# 测试书单\n\n## 第一本\n\n- [ ] 第一章\n- [x] 第二章\n';
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
   if(m.type==='init')setTimeout(()=>window.dispatchEvent(new MessageEvent('message',{data:{type:'open',content:{content:fixture,rootPath:origin+'/resource/markdown',workspaceBaseUrl:'',documentCacheId:'large-file-test',fileName:'有声书',config:{language:'zh-cn',isWeb:false,isDev:false,isPro:false,isProCancelled:true,editMode:'wysiwyg',editorTheme:'Auto',codeMirrorTheme:'Auto',mermaidTheme:'Auto',preventMacOptionKey:true},viewerSettings:{enabled:false}}}})),0);
   if(m.type==='getExportDirectory')setTimeout(()=>window.dispatchEvent(new MessageEvent('message',{data:{type:'exportDirectory',content:{requestId:m.content,directory:'/tmp',isDefault:true}}})),0);
   if(m.type==='queryAIAvailable')setTimeout(()=>window.dispatchEvent(new MessageEvent('message',{data:{type:'aiAvailable',content:false}})),0);
  },getState(){return{}},setState(){}});
 },{origin,fixture});
 const started=Date.now();await page.goto(origin+'/resource/markdown/index.html');
 await page.waitForSelector('.vditor-wysiwyg [contenteditable="true"], .vditor-wysiwyg[contenteditable="true"]',{timeout:30000});
 // Even missing/cancelled account states must expose the local capabilities.
 assert.equal(await page.locator('body.is-pro').count(),1);
 assert.equal(await page.locator('body.is-pro-cancelled').count(),0);
 assert.equal(await page.locator('[data-type="pro-upgrade"]').count(),0);
 console.log('CHECK keymap');
 await page.locator('[data-type="keymap"]').click();
 await page.waitForSelector('.vditor-keymap-dialog');
 assert.ok(await page.locator('.vditor-keymap-dialog').isVisible());
 await page.keyboard.press('Escape');
 await page.evaluate(()=>window.dispatchEvent(new MessageEvent('message',{
   data:{type:'proStatus',content:{isPro:false,isProCancelled:true}}
 })));
 assert.equal(await page.locator('body.is-pro-cancelled').count(),0);
 assert.equal(await page.locator('body.is-pro').count(),1);
 console.log('CHECK export');
 await page.locator('[data-type="export"]').click();
 await page.waitForSelector('[data-export-action="export"]');
 const font=page.locator('[data-export-option="fontSize"]');
 assert.equal(await font.isEnabled(),true);
 await font.selectOption('18');
 await page.locator('label').filter({has:page.locator('[data-export-option="type"][value="html"]')}).click();
 await page.locator('[data-export-action="export"]').click();
 await page.waitForFunction(()=>window.__messages.some(m=>m.type==='export'));
 const exported=await page.evaluate(()=>window.__messages.find(m=>m.type==='export').content);
 assert.equal(exported.useProExport,true);
 assert.equal(exported.type,'html');
 assert.equal(exported.exportTheme.layout.fontSize,'18px','custom export settings must reach the host');
 console.log('CHECK markmap');
 await page.locator('[data-type="markmap"]').click();
 await page.waitForSelector('#office-markmap-overlay');
 assert.equal(await page.locator('.office-markmap-lock').isVisible(),false);
 assert.equal(await page.evaluate(()=>window.__messages.some(m=>m.type==='openProPanel')),false);
 assert.deepEqual(errors,[]);
 console.log('PASS: local features survive missing/cancelled licenses: keymap, themed export and markmap');
 await browser.close();server.close();
})().catch(e=>{console.error(e);server.close();process.exitCode=1;setTimeout(()=>process.exit(1),1000)});
