from pathlib import Path
import argparse, json, shutil
from local_features import patch_host
p=argparse.ArgumentParser();p.add_argument('--base',required=True,help='Official Office Viewer 4.2.3 extension directory');p.add_argument('--output',required=True);args=p.parse_args()
base=Path(args.base).resolve();out=Path(args.output).resolve();source=Path(__file__).parent
assert json.loads((base/'package.json').read_text())['version']=='4.2.3'
assert not out.exists(),'Use a fresh output directory'
shutil.copytree(base,out)
for name in ('extension.js', 'extension.web.js'):
 target=out/'out'/name
 target.write_text(patch_host(target.read_text(), web=name=='extension.web.js'))
for f in (source/'resource').rglob('*'):
 if f.is_file():
  dest=out/f.relative_to(source);dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(f,dest)
a=json.loads((out/'package.json').read_text());a.update(name='office-largefile-local',publisher='local',version='4.2.3-local.3',displayName='Office Viewer · 大文件免费版',description='Local modification: free bundled features, automatic page turning and adjustable Markdown page capacity');a.pop('__metadata',None);a['scripts']={}
for command in a.get('contributes',{}).get('commands',[]):
 if command['command']=='office.pro.activate': command['title']='Office Viewer: 本地免费版功能说明'
(out/'package.json').write_text(json.dumps(a,ensure_ascii=False,indent=2)+'\n')
(out/'.vscodeignore').write_text('.github/**\n**/*.map\n.vsixmanifest\n')
shutil.copy2(source/'README.md',out/'README-LOCAL.md')
print('Run: npx @vscode/vsce package --no-dependencies (inside '+str(out)+')')
