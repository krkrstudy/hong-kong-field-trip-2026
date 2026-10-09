import {mkdir,rm,cp} from 'node:fs/promises';
await import('./prepare-pwa.mjs');await import('./check.mjs');
await rm('dist',{recursive:true,force:true});await mkdir('dist');for(const f of ['index.html','style.css','app.js','pwa.js','sw.js','manifest.webmanifest','data.json','sources.json','assets','vendor','downloads'])await cp(f,'dist/'+f,{recursive:true});console.log('Static PWA built in dist/');
