import {mkdir,rm,cp} from 'node:fs/promises';
await rm('dist',{recursive:true,force:true});await mkdir('dist');for(const f of ['index.html','style.css','app.js','data.json','sources.json','assets','vendor','downloads'])await cp(f,'dist/'+f,{recursive:true});await import('./check.mjs');console.log('Static site built in dist/');
