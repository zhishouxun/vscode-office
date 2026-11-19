import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { LargeDocument, splitDocument, PAGE_LINES, PAGE_CHARS } from '../largeDocument.mjs';
import { literalMatches, resolveAnchor } from '../resource/markdown/documentNavigation.mjs';
const table = '| 名称 | 链接 |\n|---|---|\n' + Array.from({length:1000},(_,i)=>`| 小说${i} | https://example.com/${i} |\n`).join('');
test('page ranges reproduce all source bytes and tables have bounded pages',()=>{
 const m = new LargeDocument(table,{force:true,maxLines:50});
 assert.equal(m.spans.map(s=>table.slice(s.start,s.end)).join(''),table);
 assert.ok(m.spans.length>=20);
 assert.ok(m.page(3).startsWith('| 名称 | 链接 |\n|---|---|\n'));
 assert.ok(m.page(3).split('\n').length<55);
});
test('middle-page edit keeps prior/later pages byte-identical, no duplicated headers',()=>{
 const m=new LargeDocument(table,{force:true,maxLines:50});
 const span={...m.spans[5]},head=table.slice(0,span.start),tail=table.slice(span.end);
 const page=m.page(5);const edited=page.replace('小说248','修改248');
 assert.notEqual(page,edited);m.commit(edited);
 assert.ok(m.text.startsWith(head));assert.ok(m.text.endsWith(tail));
 assert.equal((m.text.match(/\| 名称 \| 链接 \|/g)||[]).length,1);
 assert.equal(m.text,table.replace('小说248','修改248'));
});
test('deleting synthetic header rejects the edit without corrupting source',()=>{
 const m=new LargeDocument(table,{force:true,maxLines:50});m.page(5);
 assert.throws(()=>m.commit('not a table\n'),/表头/);assert.equal(m.text,table);
});
test('navigation and later edits retain previous edits with changed lengths',()=>{
 const m=new LargeDocument(table,{force:true,maxLines:50});
 m.commit(m.page(2).replace('小说100','长长的更新100'));
 m.commit(m.page(8).replace('小说400','更新400'));
 assert.equal(m.text,table.replace('小说100','长长的更新100').replace('小说400','更新400'));
});
test('fenced code and YAML frontmatter are not split, CRLF is preserved',()=>{
 const text='---\ntitle: hello\n---\n\n```js\n'+ 'x\n'.repeat(120)+'```\n\n'+table;
 const spans=splitDocument(text,50,5);assert.ok(spans.some(s=>text.slice(s.start,s.end).includes('```js\n'+'x\n'.repeat(120)+'```')));
 const crlf=table.replaceAll('\n','\r\n');const m=new LargeDocument(crlf,{force:true,maxLines:50});
 m.commit(m.page(4).replace('小说200','更新200').replaceAll('\r\n','\n'));
 assert.equal(m.text,crlf.replace('小说200','更新200'));
});
test('search traverses source, wraps, and locates the target page',()=>{
 const m=new LargeDocument(table,{force:true,maxLines:50});
 const pos=m.find('小说999');assert.ok(pos>0);assert.ok(m.indexAt(pos)>0);
 assert.equal(m.find('小说999',pos+1),pos);assert.equal(m.find('小说0',0,true),table.indexOf('小说0'));
 assert.equal(m.find('不存在'),-1);
});
test('small files remain one page',()=>{const m=new LargeDocument('# hi\n');assert.equal(m.enabled,false);assert.equal(m.page(),'# hi\n');});
test('larger default pages preserve the full resource catalogue',()=>{
 const text=(process.env.MARKDOWN_TEST_FILE ? readFileSync(process.env.MARKDOWN_TEST_FILE,'utf8') : table.repeat(2));
 const m=new LargeDocument(text);assert.equal(m.enabled,true);assert.ok(m.spans.length>1);
 assert.equal(PAGE_LINES,320);assert.equal(PAGE_CHARS,48000);
 for(const span of m.spans) assert.ok(text.slice(span.start,span.end).split('\n').length<=PAGE_LINES+1);
 assert.equal(m.spans.map(s=>text.slice(s.start,s.end)).join(''),text);
 console.log(JSON.stringify({sourceChars:text.length,pages:m.spans.length,maxPageChars:Math.max(...m.spans.map((s,i)=>m.page(i).length))}));
});

test('resizing after a middle-page edit retains source, location and synthetic headers',()=>{
 const m=new LargeDocument(table,{force:true,maxLines:50});
 const page=m.page(5),originalStart=m.spans[5].start;
 m.commit(page.replace('小说248','修改248'));
 const expected=table.replace('小说248','修改248');
 for (const lines of [640,160,320]) {
  const anchor=m.spans[m.index].start;
  const index=m.setPageSize(lines*150,lines);
  assert.equal(m.text,expected);
  assert.ok(m.spans[index].start<=anchor);
  assert.ok(m.spans[index].end>anchor);
  assert.equal(m.spans.map(s=>m.text.slice(s.start,s.end)).join(''),expected);
  m.page(m.indexAt(originalStart));
 }
 assert.throws(()=>m.setPageSize(0,NaN),RangeError);
});

test('changing a repeated header is rejected instead of silently discarding user input',()=>{
 const m=new LargeDocument(table,{force:true,maxLines:50});const page=m.page(3);
 assert.throws(()=>m.commit(page.replace('名称','changed')),/表头/);assert.equal(m.text,table);
});

test('prepared neighbors are bounded and do not change the active page or duplicate headers',()=>{
 const m=new LargeDocument(table,{force:true,maxLines:50});
 m.page(4);const prefix=m.prefix,index=m.index;
 const prepared=m.preparePage(5);m.preparePage(3);
 assert.equal(m.index,index);assert.equal(m.prefix,prefix);
 assert.equal(m.page(5),prepared.value);
 assert.equal(m.pageCache.size,3);
 const prefixIndex=m.tablePrefixes;
 for(let i=0;i<m.spans.length;i++)m.page(i);
 assert.equal(m.tablePrefixes,prefixIndex);assert.ok(m.pageCache.size<=3);
 m.commit(m.page(0).replace('名称','书名'));
 assert.equal(m.pageCache.size,0);
 assert.ok(m.page(5).startsWith('| 书名 | 链接 |'));
 m.setPageSize(48000,320);assert.equal(m.pageCache.size,0);
});

test('search indexes preserve original Unicode offsets, overlapping matches and invalidation',()=>{
 const m=new LargeDocument('İ x 😀😀 aaaa [a.b] A.B\n');
 const positions=m.matches('😀');assert.deepEqual(positions,[4,6]);
 assert.equal(m.matches('😀'),positions);
 assert.deepEqual(m.matches('aa'),[9,10,11]);
 assert.deepEqual(m.matches('a.b'),[15,20]);
 assert.deepEqual(literalMatches(m.text,'😀'),[4,6]);
 m.page();m.commit('new 😀\n');assert.deepEqual(m.matches('😀'),[4]);
 assert.equal(m.find('😀',5),4);assert.equal(m.find('😀',0,true),4);
});

test('reading anchors relocate after insertions and prefer their source context',()=>{
 const original='before\n相同词\nmarker\n相同词\nafter\n';
 const offset=original.lastIndexOf('相同词');
 const anchor={offset,quote:'相同词',before:original.slice(Math.max(0,offset-10),offset),after:original.slice(offset+3)};
 const changed='inserted\n'+original;
 assert.equal(resolveAnchor(changed,anchor),offset+9);
 assert.equal(resolveAnchor('short', {offset:999,quote:'missing'}),4);
 assert.equal(resolveAnchor(original,{offset:NaN}),0);
});
