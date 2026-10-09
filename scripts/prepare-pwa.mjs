import fs from 'node:fs';
import crypto from 'node:crypto';
const required=['index.html','style.css','app.js','pwa.js','data.json','sources.json','manifest.webmanifest'];
function walk(dir){return fs.readdirSync(dir,{withFileTypes:true}).flatMap(e=>e.isDirectory()?walk(dir+'/'+e.name):[dir+'/'+e.name]);}
const assets=[...required,...walk('assets'),...walk('vendor'),...walk('downloads')].sort();
const hash=crypto.createHash('sha256');for(const p of assets){hash.update(p);hash.update(fs.readFileSync(p));}const template=fs.readFileSync('scripts/sw-template.js','utf8');hash.update(template);
const version='20261009-'+hash.digest('hex').slice(0,14);
fs.writeFileSync('sw.js',template.replace('__VERSION__',JSON.stringify(version)).replace('__ASSETS__',JSON.stringify(assets)));
console.log('Prepared offline bundle:',assets.length,'files,',Math.round(assets.reduce((n,p)=>n+fs.statSync(p).size,0)/1024/1024*10)/10,'MB,',version);
