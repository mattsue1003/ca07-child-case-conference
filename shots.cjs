// 每頁截圖（隱藏影片、保留手機框）＋輸出影片位置與講稿，給 build_pptx.py 用
const puppeteer=require('/Users/matt/Claude/影片/功能性訓練/node_modules/puppeteer-core');
const fs=require('fs'),path=require('path');
(async()=>{
  const b=await puppeteer.launch({executablePath:'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',headless:'new'});
  const p=await b.newPage();await p.setViewport({width:1280,height:720,deviceScaleFactor:1.5});
  await p.goto('file://'+path.resolve('index.html'),{waitUntil:'networkidle0'});
  await p.evaluate(()=>document.fonts.ready);
  await p.addStyleTag({content:'.clip video{visibility:hidden}#hud{display:none!important}'});
  fs.mkdirSync('pptx_build',{recursive:true});
  const n=await p.evaluate(()=>slides.length),meta=[];
  for(let i=0;i<n;i++){
    await p.evaluate(i=>go(i),i);await new Promise(r=>setTimeout(r,500));
    const m=await p.evaluate(()=>{const s=document.querySelector('.slide.on'),sr=s.getBoundingClientRect(),k=sr.width/1280;
      const c=s.querySelector('.clip');const cr=c?c.getBoundingClientRect():null;
      return {clip:s.dataset.clip,title:s.dataset.title,note:(s.querySelector('.note')||{}).textContent||'',
        rect:cr?{x:(cr.left-sr.left)/k,y:(cr.top-sr.top)/k,w:cr.width/k,h:cr.height/k}:null};});
    const f=`pptx_build/s${String(i+1).padStart(2,'0')}.png`;
    await p.screenshot({path:f,clip:{x:0,y:0,width:1280,height:720}});
    meta.push({...m,img:f});
  }
  fs.writeFileSync('pptx_build/meta.json',JSON.stringify(meta,null,1));
  await b.close();console.log('shots',n);
})();
