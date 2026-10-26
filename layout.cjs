// 從網頁簡報抽出每頁的文字區塊、卡片、影片框位置與樣式，給 build_pptx.py 組成可編輯的 PowerPoint
const puppeteer=require('/Users/matt/Claude/影片/功能性訓練/node_modules/puppeteer-core');
const fs=require('fs'),path=require('path');
(async()=>{
  const b=await puppeteer.launch({executablePath:'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',headless:'new'});
  const p=await b.newPage();await p.setViewport({width:1280,height:720});
  await p.goto('file://'+path.resolve('index.html'),{waitUntil:'networkidle0'});
  await p.evaluate(()=>document.fonts.ready);
  await p.addStyleTag({content:'#hud{display:none!important}'});
  const n=await p.evaluate(()=>slides.length),out=[];
  for(let i=0;i<n;i++){
    await p.evaluate(i=>go(i),i);await new Promise(r=>setTimeout(r,250));
    out.push(await p.evaluate(()=>{
      const s=document.querySelector('.slide.on'),sr=s.getBoundingClientRect(),k=sr.width/1280;
      const R=e=>{const r=e.getBoundingClientRect();return {x:(r.left-sr.left)/k,y:(r.top-sr.top)/k,w:r.width/k,h:r.height/k}};
      const BLOCK=new Set(['DIV','P','LI','H1','H2','H3','TD','TH','UL','OL','TABLE','TR','TBODY','SECTION']);
      const boxes=[],texts=[];
      s.querySelectorAll('.card,.quote,.deid').forEach(e=>{const cs=getComputedStyle(e);
        boxes.push({...R(e),bg:cs.backgroundColor,line:cs.borderTopColor,lineW:parseFloat(cs.borderTopWidth),lineStyle:cs.borderTopStyle,
          left:cs.borderLeftColor,leftW:parseFloat(cs.borderLeftWidth),radius:parseFloat(cs.borderTopLeftRadius)});});
      const isLeaf=e=>BLOCK.has(e.tagName)&&e.textContent.trim()&&![...e.querySelectorAll('*')].some(c=>BLOCK.has(c.tagName));
      const runsOf=(node,inh,acc)=>{node.childNodes.forEach(c=>{
        if(c.nodeType===3){if(c.textContent.replace(/\s+/g,'')||c.textContent===' ')acc.push({t:c.textContent.replace(/\s*\n\s*/g,''),...inh});}
        else if(c.nodeType===1){if(c.tagName==='BR'){acc.push({br:1});return;}
          const cs=getComputedStyle(c);runsOf(c,{b:parseInt(cs.fontWeight)>=600,c:cs.color,sz:parseFloat(cs.fontSize),f:cs.fontFamily},acc);}});return acc;};
      s.querySelectorAll('.main *').forEach(e=>{if(!isLeaf(e))return;const cs=getComputedStyle(e);
        if(cs.display==='none')return;
        const base={b:parseInt(cs.fontWeight)>=600,c:cs.color,sz:parseFloat(cs.fontSize),f:cs.fontFamily};
        texts.push({...R(e),runs:runsOf(e,base,[]),lh:parseFloat(cs.lineHeight)/parseFloat(cs.fontSize)||1.3,
          li:e.tagName==='LI',align:cs.textAlign,pl:parseFloat(cs.paddingLeft),pt:parseFloat(cs.paddingTop),font:cs.fontFamily,
          sz:parseFloat(cs.fontSize),c:cs.color,b:base.b});});
      const c=s.querySelector('.clip');
      return {title:s.dataset.title,clip:s.dataset.clip,cover:s.classList.contains('cover'),num:s.dataset.n,
        note:(s.querySelector('.note')||{}).textContent||'',boxes,texts,clipRect:c?R(c):null};
    }));
  }
  fs.mkdirSync('pptx_build',{recursive:true});
  fs.writeFileSync('pptx_build/layout.json',JSON.stringify(out));
  await b.close();console.log('layout',n,out.map(o=>o.texts.length).join(','));
})();
