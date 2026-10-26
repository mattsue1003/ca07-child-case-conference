const puppeteer=require('/Users/matt/Claude/影片/功能性訓練/node_modules/puppeteer-core');
const {spawn}=require('child_process');const fs=require('fs');const path=require('path');
const FF='/Users/matt/Documents/short/ffmpeg-bin/ffmpeg';
const mode=process.argv[2]||'sheet',only=process.argv[3];
(async()=>{
  const b=await puppeteer.launch({executablePath:'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',headless:'new'});
  const p=await b.newPage();await p.setViewport({width:540,height:960});
  await p.goto('file://'+path.resolve('clips.html')+'?render=1',{waitUntil:'networkidle0'});
  await p.evaluate(()=>document.fonts.ready);
  const clips=await p.evaluate(()=>window.CLIPS);
  const el=await p.$('#v');
  fs.mkdirSync('video',{recursive:true});fs.mkdirSync('sheet',{recursive:true});
  for(const [name,L] of Object.entries(clips)){
    if(only&&!only.split(',').includes(name))continue;
    if(mode==='sheet'){
      for(const f of [.2,.45,.7,.92]){await p.evaluate((n,t)=>draw(n,t),name,L*f);await el.screenshot({path:`sheet/${name}_${Math.round(f*100)}.jpg`,type:'jpeg',quality:70});}
      continue;
    }
    const fps=24,N=Math.round(L*fps);
    const ff=spawn(FF,['-y','-loglevel','error','-f','image2pipe','-framerate',String(fps),'-c:v','mjpeg','-i','-','-c:v','libx264','-pix_fmt','yuv420p','-crf','24','-preset','medium','-movflags','+faststart','-an',`video/${name}.mp4`]);
    for(let i=0;i<N;i++){await p.evaluate((n,t)=>draw(n,t),name,i/fps);const buf=await el.screenshot({type:'jpeg',quality:92});if(!ff.stdin.write(buf))await new Promise(r=>ff.stdin.once('drain',r));}
    ff.stdin.end();await new Promise(r=>ff.on('close',r));
    await p.evaluate((n,t)=>draw(n,t),name,L*.72);await el.screenshot({path:`video/${name}.jpg`,type:'jpeg',quality:80});
    console.log('done',name,N);
  }
  await b.close();
})();
