# 由 pptx_build/ 的截圖與 meta.json 組成 PowerPoint，每頁右側嵌入 9:16 影片（自動播放、循環、靜音）
import json, copy
from pptx import Presentation
from pptx.util import Emu
from lxml import etree
PX=9525  # 1280px = 13.333in
meta=json.load(open('pptx_build/meta.json',encoding='utf-8'))
prs=Presentation(); prs.slide_width=Emu(1280*PX); prs.slide_height=Emu(720*PX)
blank=prs.slide_layouts[6]
NS='http://schemas.openxmlformats.org/presentationml/2006/main'
def timing(spid):
    return f'''<p:timing xmlns:p="{NS}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
<p:par><p:cTn id="2" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="4" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="5" dur="indefinite" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:cmd>
</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
<p:video><p:cMediaNode vol="80000" mute="1" numSld="999" showWhenStopped="1"><p:cTn id="6" repeatCount="indefinite" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cMediaNode></p:video>
</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'''
for m in meta:
    s=prs.slides.add_slide(blank)
    s.shapes.add_picture(m['img'],0,0,prs.slide_width,prs.slide_height)
    if m['clip'] and m['rect']:
        r=m['rect']; inset=10
        iw=r['w']-2*inset; ih=iw*16/9
        if ih>r['h']-2*inset: ih=r['h']-2*inset; iw=ih*9/16
        x=r['x']+(r['w']-iw)/2; y=r['y']+(r['h']-ih)/2
        mv=s.shapes.add_movie(f"video/{m['clip']}.mp4",Emu(int(x*PX)),Emu(int(y*PX)),Emu(int(iw*PX)),Emu(int(ih*PX)),
                              poster_frame_image=f"video/{m['clip']}.jpg",mime_type='video/mp4')
        el=s.element
        for t in el.findall(f'{{{NS}}}timing'): el.remove(t)
        el.append(etree.fromstring(timing(mv.shape_id)))
    s.notes_slide.notes_text_frame.text=m['note'].strip()
prs.save('個案研討簡報.pptx'); print('slides',len(meta))
