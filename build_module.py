#!/usr/bin/env python3
"""
EM2 MODULE BUILDER — auto-loads a full interactive lesson deck from 4 source files.
See HANDOFF.md for usage. Run with --help for arguments.
"""
import os, re, sys, glob, shutil, subprocess, argparse, json, tempfile
import numpy as np
from pptx import Presentation
from pptx.util import Inches as I, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn
from PIL import Image, ImageChops

HERE = os.path.dirname(os.path.abspath(__file__))

NAVY=RGBColor(0x18,0x30,0x5A); TEAL=RGBColor(0x2C,0x7A,0x7B)
RED=RGBColor(0xC0,0x39,0x2B); ORANGE=RGBColor(0xC9,0x82,0x1B); GOLD=RGBColor(0xB8,0x90,0x1A)
GREEN=RGBColor(0x2E,0x7D,0x32); BLUE=RGBColor(0x2F,0x6F,0xE0)
INK=RGBColor(0x27,0x31,0x3B); MUTED=RGBColor(0x7A,0x87,0x94); BODY=RGBColor(0x46,0x53,0x5F)
CARD=RGBColor(0xEE,0xF3,0xFA); LINE=RGBColor(0xD8,0xE0,0xEC); WHITE=RGBColor(0xFF,0xFF,0xFF)
CHEV=RGBColor(0xC2,0xCC,0xD8); TEAL_BAR=RGBColor(0x3F,0xB6,0xA8); CLOCK=RGBColor(0xCF,0xE0,0xFF)
SERIF="Georgia"; SANS="Calibri"
SW, SH = 10.0, 5.625
TOPIC_COLORS=[RED,ORANGE,GOLD,GREEN,BLUE]; TOPIC_NAMES=["TOPIC A","TOPIC B","TOPIC C","TOPIC D","TOPIC E"]
AG_CROP=(0.020,0.193,0.980,0.952)

# ---------- source reading ----------
def shape_text(shapes):
    out=[]
    for sh in shapes:
        if sh.shape_type==6: out.append(shape_text(sh.shapes))
        if sh.has_table:
            for r in sh.table.rows:
                for c in r.cells: out.append(c.text)
        if sh.has_text_frame: out.append(sh.text_frame.text)
    return "\n".join(out)
def slide_texts(pptx): return [shape_text(s.shapes) for s in Presentation(pptx).slides]

def detect_dares(dares_pptx):
    texts=slide_texts(dares_pptx); grade=module=None; lessons={}
    for t in texts:
        gm=re.search(r'Grade\s*(\d+),\s*Module\s*(\d+)',t)
        if gm and grade is None: grade,module=int(gm.group(1)),int(gm.group(2))
        # Lesson labels appear in two conventions across our source decks:
        #   long form  -> "Lesson 1", "Lesson 2", ...
        #   short form -> "L1 Question:", "L2 Question:", ...
        # Match the long form first (existing behaviour), then fall back to the
        # short form. The short-form pattern is anchored to "Question" so a bare
        # token like an L in a standard code can't be mistaken for a lesson.
        ln=re.search(r'Lesson\s*(\d+)',t) or re.search(r'\bL(\d+)\s*Question',t)
        if not ln: continue
        n=int(ln.group(1))
        q=re.search(r'L\d+\s*Question:\s*(.*?)\s*Answer Statement:',t,re.S)
        w=re.search(r'Words:\s*(.*?)\s*(?:CCSS|DARE\s*©|\Z)',t,re.S)
        ccss=re.search(r'CCSS[.\s]*MATH[.\s]*CONTENT[.\s]*([0-9A-Z.]+)',t) or re.search(r'\b(\d[A-Z]{1,3}\.[A-Z0-9.]+)\b',t)
        lessons[n]={'n':n,'question':(q.group(1).strip() if q else ""),
                    'words':(re.sub(r'\s+',' ',w.group(1)).strip() if w else ""),
                    'ccss':(ccss.group(1).rstrip('.') if ccss else "")}
    return grade, module, [lessons[k] for k in sorted(lessons)]

def detect_sort_lessons(sorts_pptx):
    texts=slide_texts(sorts_pptx); mapping={}; i=0
    while i<len(texts):
        lab=re.search(r'\bL(\d+)\b',texts[i])
        if lab and texts[i].count('L'+lab.group(1))>=3 and i+1<len(texts):
            mapping[int(lab.group(1))]=i+2; i+=2
        else: i+=1
    return mapping

# ---------- rendering / image ops ----------
# Optional map of {source_pptx_path: pre_rendered_pdf_path}. When a source file
# has an entry here, we skip the memory-heavy LibreOffice step and rasterize the
# supplied PDF directly with pdftoppm (poppler), which uses far less memory.
PRERENDERED_PDFS = {}

def render_to_pngs(pptx,outdir,prefix,dpi=150):
    os.makedirs(outdir,exist_ok=True); work=tempfile.mkdtemp()
    supplied=PRERENDERED_PDFS.get(pptx)
    if supplied and os.path.exists(supplied):
        # Use the pre-rendered PDF; no LibreOffice needed.
        pdf=supplied
    else:
        # Fall back to converting the pptx with LibreOffice (original behaviour).
        shutil.copy(pptx,work)
        base=os.path.join(work,os.path.basename(pptx))
        subprocess.run(["soffice","--headless","--convert-to","pdf","--outdir",work,base],
                       check=True,capture_output=True,timeout=300,env={**os.environ,"HOME":work})
        pdf=base.rsplit('.',1)[0]+'.pdf'
    subprocess.run(["pdftoppm","-png","-r",str(dpi),pdf,os.path.join(outdir,prefix)],check=True)
    return sorted(glob.glob(os.path.join(outdir,prefix+"*.png")))

def trim(path,pad=8):
    im=Image.open(path).convert('RGB')
    bb=ImageChops.difference(im,Image.new('RGB',im.size,(255,255,255))).getbbox()
    if bb:
        l,t,r,b=bb; im=im.crop((max(0,l-pad),max(0,t-pad),min(im.width,r+pad),min(im.height,b+pad)))
    im.save(path); return path

def crop_box(path,frac,out):
    im=Image.open(path).convert('RGB'); W,H=im.size; l,t,r,b=frac
    im.crop((int(W*l),int(H*t),int(W*r),int(H*b))).save(out); return out

def _trim_im(im,pad=6):
    bb=ImageChops.difference(im.convert('RGB'),Image.new('RGB',im.size,(255,255,255))).getbbox()
    if bb:
        l,t,r,b=bb; im=im.crop((max(0,l-pad),max(0,t-pad),min(im.width,r+pad),min(im.height,b+pad)))
    return im

def crop_sort_cells(path,outdir,tag):
    """Cut the FIRST ROW of the 3-column sort sheet into 3 card images (col1,col2,col3)."""
    im=Image.open(path).convert('RGB'); g=np.array(im.convert('L')); H,W=g.shape
    rowfrac=(g<180).mean(axis=1)
    cand=[y for y in range(int(H*0.04),int(H*0.6)) if rowfrac[y]>0.35]
    row_bottom=cand[0] if cand else H//7
    colw=W/3; cells=[]
    for c in range(3):
        x0=int(c*colw)+8; x1=int((c+1)*colw)-8
        cell=_trim_im(im.crop((x0,8,x1,row_bottom-4)))
        p=os.path.join(outdir,f'{tag}_cell{c}.png'); cell.save(p); cells.append(p)
    return cells

def detect_mathtalks(file,img):
    """Group math-talk pages by lesson (handles 2 pages, or 4 if Math Talk A & B)."""
    texts=slide_texts(file); pngs=render_to_pngs(file,img,'mt'); by={}
    for i,t in enumerate(texts):
        m=re.search(r'Lesson\s*(\d+)',t)
        if m and i<len(pngs): by.setdefault(int(m.group(1)),[]).append(trim(pngs[i]))
    return by,pngs

# ---------- drawing helpers ----------
def A(name): return os.path.join(ASSETS,name)
def _set_font(r,size,color,bold=False,italic=False,font=SANS):
    r.font.size=Pt(size); r.font.bold=bold; r.font.italic=italic; r.font.name=font; r.font.color.rgb=color
def _sup(r): r.font._rPr.set('baseline','30000')
def fmt(sec): return f"{sec//60}:{sec%60:02d}"

def text(s,l,t,w,h,runs,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,sp_after=2):
    tb=s.shapes.add_textbox(I(l),I(t),I(w),I(h)); tf=tb.text_frame
    tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=0;tf.margin_right=0;tf.margin_top=0;tf.margin_bottom=0
    first=True
    for line in runs:
        p=tf.paragraphs[0] if first else tf.add_paragraph()
        p.alignment=align; p.space_after=Pt(sp_after); p.space_before=Pt(0); first=False
        for seg in line:
            r=p.add_run(); r.text=seg[0]
            _set_font(r,seg[1],seg[2],seg[3] if len(seg)>3 else False,seg[4] if len(seg)>4 else False,seg[5] if len(seg)>5 else SANS)
    return tb

def rect(s,l,t,w,h,fill=None,line_c=None,line_w=1.0,round_=False,shadow=False,radius=0.08):
    shp=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE,I(l),I(t),I(w),I(h))
    if round_:
        try: shp.adjustments[0]=radius
        except Exception: pass
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb=fill
    if line_c is None: shp.line.fill.background()
    else: shp.line.color.rgb=line_c; shp.line.width=Pt(line_w)
    shp.shadow.inherit=False
    if shadow:
        sp=shp._element.spPr; el=sp.makeelement(qn('a:effectLst'),{}); sp.append(el)
        sh=el.makeelement(qn('a:outerShdw'),{'blurRad':'90000','dist':'40000','dir':'5400000','rotWithShape':'0'}); el.append(sh)
        clr=sh.makeelement(qn('a:srgbClr'),{'val':'18305A'}); sh.append(clr); clr.append(clr.makeelement(qn('a:alpha'),{'val':'16000'}))
    return shp

def pill(s,l,t,w,h,fill,label,size=11,color=WHITE):
    shp=rect(s,l,t,w,h,fill=fill,round_=True,radius=0.5)
    tf=shp.text_frame; tf.word_wrap=False; tf.margin_top=0; tf.margin_bottom=0
    p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=label; _set_font(r,size,color,bold=True); return shp

def pic(s,path,l,t,w=None,h=None):
    kw={}
    if w: kw['width']=I(w)
    if h: kw['height']=I(h)
    return s.shapes.add_picture(path,I(l),I(t),**kw)

def fit_pic(s,path,bl,bt,bw,bh,frame=False):
    iw,ih=Image.open(path).size; ar=iw/ih; bar=bw/bh
    if ar>bar: w=bw; h=bw/ar
    else: h=bh; w=bh*ar
    l=bl+(bw-w)/2; t=bt+(bh-h)/2
    if frame: rect(s,l-0.06,t-0.06,w+0.12,h+0.12,fill=WHITE,line_c=LINE,line_w=1,round_=True,radius=0.03,shadow=True)
    return s.shapes.add_picture(path,I(l),I(t),width=I(w),height=I(h))

def logo(s): pic(s,A('logo.png'),SW-1.45,0.18,w=1.15)
def section_pill(s,label,fill): pill(s,0.42,0.30,1.4,0.36,fill,label,size=11)
def title(s,txt,color,size=30): text(s,1.0,0.30,8.0,0.6,[[(txt,size,color,True,False,SERIF)]],align=PP_ALIGN.CENTER)

# ---------- EDITABLE timer: navy pill + drawn clock + clickable time text + depleting bar ----------
def timer(s,seconds,pos='tr'):
    pw=1.45; ph=0.56
    pl=(SW-pw-1.55) if pos=='tr' else 0.42; pt=0.30
    rect(s,pl,pt,pw,ph,fill=NAVY,round_=True,radius=0.22)
    cd=0.26; cx=pl+0.15; cy=pt+(ph-cd)/2
    o=s.shapes.add_shape(MSO_SHAPE.OVAL,I(cx),I(cy),I(cd),I(cd))
    o.fill.background(); o.line.color.rgb=CLOCK; o.line.width=Pt(1.5); o.shadow.inherit=False
    ccx=cx+cd/2; ccy=cy+cd/2
    for (x2,y2) in [(ccx,ccy-cd*0.34),(ccx+cd*0.27,ccy)]:
        h=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,I(ccx),I(ccy),I(x2),I(y2))
        h.line.color.rgb=CLOCK; h.line.width=Pt(1.25); h.shadow.inherit=False
    # editable time text (teacher can click and type a new value)
    tb=s.shapes.add_textbox(I(pl+0.40),I(pt),I(pw-0.46),I(ph)); tf=tb.text_frame
    tf.vertical_anchor=MSO_ANCHOR.MIDDLE; tf.word_wrap=False
    tf.margin_left=0;tf.margin_right=0;tf.margin_top=0;tf.margin_bottom=0
    p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=fmt(seconds); _set_font(r,22,WHITE,bold=True)
    tb.name="TimerText"
    bw=pw*0.86; bh=0.085; bl=pl+(pw-bw)/2; bt=pt+ph+0.05
    bar=rect(s,bl,bt,bw,bh,fill=TEAL_BAR,round_=True,radius=0.5); bar.name=f"TimerBar::{seconds}"
    return bar

def goto(shp,idx): shp.name=f"GOTO::{idx}"

def footer(s,ccss,grade,module,lesson):
    text(s,0.42,SH-0.42,4.5,0.3,[[("CCSS  ",9,INK,True),(ccss or "—",9,MUTED)]])
    text(s,SW-4.3,SH-0.42,3.88,0.3,[[(f"Grade {grade}   ·   Module {module}   ·   Lesson {lesson}",9,MUTED)]],align=PP_ALIGN.RIGHT)

def routine_card(s,l,t,w,h,color,num,img,title_,desc):
    rect(s,l,t,w,h,fill=WHITE,line_c=LINE,line_w=1,round_=True,radius=0.06,shadow=True)
    c=rect(s,l+0.18,t+0.16,0.44,0.44,fill=color,round_=True,radius=0.5)
    tf=c.text_frame;tf.margin_top=0;tf.margin_bottom=0;pr=tf.paragraphs[0];pr.alignment=PP_ALIGN.CENTER
    rr=pr.add_run();rr.text=str(num);_set_font(rr,18,WHITE,bold=True)
    fit_pic(s,img,l+0.25,t+0.52,w-0.5,1.30)
    text(s,l+0.12,t+2.02,w-0.24,0.4,[[(title_,15.5,NAVY,True)]],align=PP_ALIGN.CENTER)
    text(s,l+0.2,t+2.46,w-0.4,0.95,[[(desc,12.5,BODY)]],align=PP_ALIGN.CENTER)

def reflected_callout(s,l,t,w,h):
    rect(s,l,t,w,h,fill=CARD,line_c=GREEN,line_w=2.0,round_=True,radius=0.14)
    b=s.shapes[-1]; tf=b.text_frame; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text="My edit was\u2026"; _set_font(r,18,GREEN,italic=True,font=SERIF)
    bx=l+0.45; by=t+h
    tri=s.shapes.build_freeform(I(bx),I(by),scale=I(1)/914400)
    tri.add_line_segments([(I(bx+0.30),I(by)),(I(bx),I(by+0.40)),(I(bx),I(by))],close=True)
    sh=tri.convert_to_shape(); sh.fill.solid(); sh.fill.fore_color.rgb=CARD
    sh.line.color.rgb=GREEN; sh.line.width=Pt(2.0); sh.shadow.inherit=False

def dropzone(s,l,t,w,h,label,field,fs=14):
    shp=rect(s,l,t,w,h,fill=RGBColor(0xF7,0xFA,0xFE),line_c=RGBColor(0xB7,0xC4,0xD6),line_w=1.75,round_=True,radius=0.04)
    ln=shp._element.spPr.find(qn('a:ln')); ln.append(ln.makeelement(qn('a:prstDash'),{'val':'dash'}))
    tf=shp.text_frame; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=label; _set_font(r,fs,RGBColor(0x88,0x95,0xA6),bold=True)
    if field:
        p2=tf.add_paragraph(); p2.alignment=PP_ALIGN.CENTER
        r2=p2.add_run(); r2.text=field; _set_font(r2,10,RGBColor(0xA6,0xB2,0xC0),italic=True)

SCHED=[('Math Talk','sched_2.png',RED),('Randomizer','sched_3.png',ORANGE),
       ('Sort','sched_4.png',GOLD),('DARE','sched_5.png',GREEN),('Game','sched_6.png',BLUE)]

# ---------- slide builders ----------
def b_toc(prs,MT):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s)
    text(s,0.6,0.60,8.8,0.8,[[("EM",34,NAVY,True,False,SERIF),("2",20,NAVY,True,False,SERIF),
        (f"   Grade {MT['grade']} · Module {MT['module']}",34,NAVY,True,False,SERIF)]],align=PP_ALIGN.CENTER)
    for r in s.shapes[-1].text_frame.paragraphs[0].runs:
        if r.text=="2": _sup(r)
    text(s,0.6,1.36,8.8,0.4,[[(f"{MT['title']} — Table of Contents",15,TEAL,False,True,SERIF)]],align=PP_ALIGN.CENTER)
    topics=MT['topics']; n=len(topics); gap=0.16; mL=0.5; cw=(SW-2*mL-(n-1)*gap)/n; top=1.95; ch=2.95
    for i,tp in enumerate(topics):
        l=mL+i*(cw+gap); rect(s,l,top,cw,ch,fill=CARD,line_c=LINE,line_w=1,round_=True,radius=0.05)
        pill(s,l+0.12,top+0.14,cw-0.24,0.32,tp['color'],tp['name'],size=10.5); y=top+0.62
        for Ln in tp['lessons']:
            tb=text(s,l+0.1,y,cw-0.2,0.26,[[(f"Lesson {Ln}",11.5,NAVY)]],align=PP_ALIGN.CENTER)
            tb.text_frame.paragraphs[0].runs[0].font.underline=True; goto(tb,MT['welcome_idx'][Ln]); y+=0.275
    text(s,0.42,SH-0.42,5,0.3,[[("CCSS  ",9,INK,True),(MT['ccss_range'],9,MUTED)]])
    text(s,SW-4.0,SH-0.42,3.58,0.3,[[(f"Grade {MT['grade']}   ·   Module {MT['module']}",9,MUTED)]],align=PP_ALIGN.RIGHT)

def b_welcome(prs,MT,L,has_sort,chip):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s)
    text(s,1.0,0.55,8.0,1.0,[[("Welcome!",54,NAVY,True,False,SERIF)]],align=PP_ALIGN.CENTER)
    text(s,1.0,1.62,8.0,0.4,[[(f"Grade {MT['grade']} · Module {MT['module']} · Lesson {L}",20,TEAL,False,True,SERIF)]],align=PP_ALIGN.CENTER)
    text(s,1.0,2.28,8.0,0.3,[[("TODAY'S SCHEDULE",12,MUTED,True)]],align=PP_ALIGN.CENTER)
    cards=[c for c in SCHED if has_sort or c[0] not in ('Randomizer','Sort')]
    n=len(cards); cw=1.45; gap=0.30; startl=(SW-(n*cw+(n-1)*gap))/2; top=2.7; ch=2.05
    for i,(name,icon,color) in enumerate(cards):
        l=startl+i*(cw+gap); rect(s,l,top,cw,ch,fill=WHITE,line_c=LINE,line_w=1,round_=True,radius=0.06,shadow=True)
        fit_pic(s,A(icon),l+0.1,top+0.12,cw-0.2,1.05)
        chs=pill(s,l+0.18,top+ch-0.5,cw-0.36,0.36,color,name,size=11)
        if name in chip: goto(chs,chip[name])
        if i<n-1: text(s,l+cw-0.02,top+0.7,gap+0.04,0.4,[[("›",24,CHEV)]],align=PP_ALIGN.CENTER)
    footer(s,MT['lesson_ccss'][L],MT['grade'],MT['module'],L)

def b_mt_routine(prs,MT,L):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s); section_pill(s,"MATH TALK",RED)
    title(s,"Math Talk Routine",RED,size=25); timer(s,120)
    cards=[("mtr_1.png","Walk to the Math Talk","Gather and stand around the projected Math Talk."),
           ("mtr_2.png","Notice & Wonder","What do you notice? What do you wonder?"),
           ("mtr_3.png","Share, Listen, Connect","Share your thinking and build on others' ideas.")]
    cw=2.7; gap=0.3; startl=(SW-(3*cw+2*gap))/2; top=1.42; ch=3.6
    for i,(img,tt,d) in enumerate(cards): routine_card(s,startl+i*(cw+gap),top,cw,ch,RED,i+1,A(img),tt,d)
    footer(s,MT['lesson_ccss'][L],MT['grade'],MT['module'],L)

def b_mt(prs,MT,L,img):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s); section_pill(s,"MATH TALK",RED)
    title(s,"Math Talk",RED); timer(s,420)
    if img and os.path.exists(img): fit_pic(s,img,1.0,1.2,8.0,3.7,frame=True)
    else: dropzone(s,1.0,1.25,8.0,3.6,"Drop the Math Talk image here","mathTalkImg")
    footer(s,MT['lesson_ccss'][L],MT['grade'],MT['module'],L)

def b_sort(prs,MT,L,cells):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s); timer(s,120,pos='tl')
    text(s,1.8,0.32,6.5,0.5,[[("Do these cards match mathematically?",20,NAVY,True,False,SERIF)]],align=PP_ALIGN.CENTER)
    text(s,1.8,0.84,6.5,0.35,[[("Why or why not?",15,TEAL,False,True,SERIF)]],align=PP_ALIGN.CENTER)
    n=3; bw=2.5; gap=0.4; startl=(SW-(n*bw+(n-1)*gap))/2; top=1.5; bh=2.15
    for i in range(3):
        l=startl+i*(bw+gap)
        rect(s,l,top,bw,bh,fill=WHITE,line_c=BLUE,line_w=2.0,round_=True,radius=0.04)
        if cells and i<len(cells) and os.path.exists(cells[i]):
            fit_pic(s,cells[i],l+0.18,top+0.18,bw-0.36,bh-0.36)
        else:
            dropzone(s,l+0.12,top+0.12,bw-0.24,bh-0.24,f"Drop sort card {i+1}",f"sortCard{i+1}",fs=12)
    text(s,1.0,3.95,8.0,1.0,[[("Tell someone near you:",13,NAVY,True)],
        [("The cards match because…      The cards do ",12.5,BODY),("NOT",12.5,BODY,True),(" match because…",12.5,BODY)],
        [("Who wants to share their thinking?",12.5,TEAL,False,True)]],align=PP_ALIGN.CENTER,sp_after=4)
    footer(s,MT['lesson_ccss'][L],MT['grade'],MT['module'],L)

def b_randroutine(prs,MT,L):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s)
    pill(s,0.42,0.30,1.6,0.36,ORANGE,"RANDOMIZER",size=10.5)
    text(s,2.1,0.32,0.4,0.34,[[("→",18,CHEV)]]); pill(s,2.5,0.30,1.0,0.36,GOLD,"SORT",size=10.5); timer(s,600)
    pw=4.25; ph=3.35; top=1.15
    rect(s,0.5,top,pw,ph,fill=CARD,line_c=LINE,line_w=1,round_=True,radius=0.05)
    text(s,0.75,top+0.14,pw-0.5,0.4,[[("Randomizer Routine",17,ORANGE,True,False,SERIF)]])
    text(s,0.82,top+0.6,pw-0.6,1.6,[[("1.  Choose a card from the basket.",12.5,BODY)],
        [("2.  Find two people whose cards mathematically match yours.",12.5,BODY)],
        [("3.  Bring the match to the teacher: \u201cOur cards match because\u2026\u201d",12.5,BODY)]],sp_after=6)
    for i,im in enumerate(["rand_1.png","rand_2.png","rand_3.png"]): fit_pic(s,A(im),0.7+i*1.22,top+2.15,1.12,1.0)
    rect(s,5.0,top,pw,ph,fill=CARD,line_c=LINE,line_w=1,round_=True,radius=0.05)
    text(s,5.25,top+0.14,pw-0.5,0.4,[[("Sort Routine",17,GOLD,True,False,SERIF)]])
    text(s,5.32,top+0.6,pw-0.6,1.5,[[("1.  With your group, prove your cards match and get the full sort from your teacher.",12.5,BODY)],
        [("2.  Work as a team to sort the cards in a way that makes sense to you.",12.5,BODY)]],sp_after=6)
    fit_pic(s,A('work_team.png'),6.55,top+2.05,1.6,1.15)
    footer(s,MT['lesson_ccss'][L],MT['grade'],MT['module'],L)

def b_dare_routine(prs,MT,L,question,words):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s); section_pill(s,"DARE",GREEN)
    title(s,"DARE Routine",GREEN); timer(s,420)
    fit_pic(s,A('dare_icons.png'),3.0,1.05,4.0,0.95)
    rect(s,0.7,2.1,8.6,2.95,fill=CARD,line_c=LINE,line_w=1,round_=True,radius=0.04)
    text(s,1.0,2.32,8.0,2.55,[
        [("Question:  ",13.5,NAVY,True),(question or "—",13.5,BODY)],
        [(" ",6,BODY)],
        [("Words:  ",13.5,NAVY,True),(words or "—",13.5,BODY)]],sp_after=4)
    footer(s,MT['lesson_ccss'][L],MT['grade'],MT['module'],L)

def b_dareguide(prs,MT,L,ag):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s); section_pill(s,"DARE",GREEN)
    title(s,"DARE Routine",GREEN); timer(s,360)
    fit_pic(s,A('dare_icons.png'),0.5,1.0,3.4,0.92)
    text(s,0.5,2.0,4.5,0.7,[[("Share & critique strategies.",16.5,NAVY,True,False,SERIF)],[("Look for your one edit!",16.5,NAVY,True,False,SERIF)]],sp_after=2)
    fit_pic(s,A('slc.png'),0.55,2.95,1.7,1.55)
    text(s,2.45,3.35,2.6,0.9,[[("What do you notice?",15.5,TEAL,True)],[("What do you wonder?",15.5,TEAL,True)]],sp_after=2)
    if ag and os.path.exists(ag): fit_pic(s,ag,5.13,0.95,4.55,4.05,frame=True)
    else: dropzone(s,5.13,0.95,4.55,4.05,"Drop the DARE answer-guide image here","dareAnswerImg")
    footer(s,MT['lesson_ccss'][L],MT['grade'],MT['module'],L)

def b_dareedit(prs,MT,L,ag):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s); section_pill(s,"DARE",GREEN)
    title(s,"DARE Routine",GREEN); timer(s,120)
    text(s,0.5,1.25,4.5,0.5,[[("Time to Edit!",26,NAVY,True,False,SERIF)]])
    text(s,0.5,1.85,4.5,0.7,[[("Use an edit pen or a different color to record new thinking, connections, or strategies.",13,BODY)]])
    fit_pic(s,A('edit_pens.png'),0.45,2.95,1.85,1.2)
    reflected_callout(s,2.45,3.0,2.6,0.92)
    if ag and os.path.exists(ag): fit_pic(s,ag,5.13,0.95,4.55,4.05,frame=True)
    else: dropzone(s,5.13,0.95,4.55,4.05,"Drop the DARE answer-guide image here","dareAnswerImg")
    footer(s,MT['lesson_ccss'][L],MT['grade'],MT['module'],L)

def b_game(prs,MT,L):
    s=prs.slides.add_slide(prs.slide_layouts[6]); logo(s); section_pill(s,"GAME",BLUE)
    title(s,"Game / Task Routine",BLUE,size=24); timer(s,420)
    cw=2.7; gap=0.3; startl=(SW-(3*cw+2*gap))/2; top=1.42; ch=3.6
    routine_card(s,startl,top,cw,ch,BLUE,1,A('gather.png'),"Gather Materials","Get your whiteboard, marker, eraser, and any tools you need.")
    routine_card(s,startl+(cw+gap),top,cw,ch,BLUE,2,A('work_team.png'),"Work with Your Team","Play the game or complete the task together, taking turns.")
    l3=startl+2*(cw+gap); rect(s,l3,top,cw,ch,fill=CARD,line_c=LINE,line_w=1,round_=True,radius=0.06)
    text(s,l3+0.2,top+0.35,cw-0.4,0.4,[[("Today's Game / Task",15,BLUE,True,False,SERIF)]],align=PP_ALIGN.CENTER)
    text(s,l3+0.25,top+1.05,cw-0.5,1.6,[[("Add today's specific game directions or task here.",13,RGBColor(0x88,0x95,0xA6),False,True)]],align=PP_ALIGN.CENTER)
    footer(s,MT['lesson_ccss'][L],MT['grade'],MT['module'],L)

def auto_topics(L_nums):
    n=len(L_nums); k=min(5,n); size=-(-n//k); topics=[]
    for i in range(0,n,size):
        idx=len(topics)
        topics.append({'name':TOPIC_NAMES[idx] if idx<5 else f"TOPIC {idx+1}",'color':TOPIC_COLORS[idx%5],'lessons':L_nums[i:i+size]})
    return topics

def build(args):
    global ASSETS; ASSETS=args.assets
    # If pre-rendered PDFs were supplied (to skip LibreOffice / save memory),
    # register them keyed by their source pptx path.
    PRERENDERED_PDFS.clear()
    for src_attr, pdf_attr in (("mathtalks","mathtalks_pdf"),("sorts","sorts_pdf"),
                               ("answerguides","answerguides_pdf"),("dares","dares_pdf")):
        src=getattr(args, src_attr, None); pdf=getattr(args, pdf_attr, None)
        if src and pdf and os.path.exists(pdf):
            PRERENDERED_PDFS[src]=pdf
    tmp=tempfile.mkdtemp(); img=os.path.join(tmp,'img'); os.makedirs(img,exist_ok=True)
    print("Reading DARE problems…")
    grade,module,dares=detect_dares(args.dares)
    if args.grade: grade=args.grade
    if args.module: module=args.module
    L_nums=[d['n'] for d in dares]; N=len(L_nums)
    ccss_map={d['n']:d['ccss'] for d in dares}; q_map={d['n']:d['question'] for d in dares}; w_map={d['n']:d['words'] for d in dares}
    print(f"  Grade {grade}, Module {module}, {N} lessons")
    print("Rendering Math Talks…")
    mt_by,mt_png=detect_mathtalks(args.mathtalks,img)
    if not mt_by:  # fallback: 2 pages per lesson in order
        for i,n in enumerate(L_nums): mt_by[n]=[trim(p) for p in mt_png[2*i:2*i+2]] or [None,None]
    for n in L_nums: mt_by.setdefault(n,[None,None])
    print("Rendering Sorts…")
    sort_cells={}
    try:
        sp=detect_sort_lessons(args.sorts); sorts=render_to_pngs(args.sorts,img,'so')
        for ln,front in sp.items():
            if 1<=front<=len(sorts): sort_cells[ln]=crop_sort_cells(trim(sorts[front-1]),img,f'L{ln}')
    except Exception as e: print("  (no sorts:",e,")")
    print("Rendering DARE answer guides…")
    ag=render_to_pngs(args.answerguides,img,'ag'); ag_map={}
    for i,n in enumerate(L_nums):
        if i<len(ag): ag_map[n]=crop_box(ag[i],AG_CROP,os.path.join(img,f'agc_{n}.png'))
    if args.topics and os.path.exists(args.topics):
        tj=json.load(open(args.topics)); topics=[{'name':t['name'],'color':TOPIC_COLORS[i%5],'lessons':t['lessons']} for i,t in enumerate(tj)]
    else: topics=auto_topics(L_nums)
    # plan
    plan=[(0,'toc')]
    for n in L_nums:
        secs=['welcome','mtr']+[f'mtp{i}' for i in range(len(mt_by[n]))]
        if n in sort_cells: secs+=['sort','rand']
        secs+=['dareroutine','dareguide','dareedit','game']
        for sec in secs: plan.append((n,sec))
    idx={(L,sec):i+1 for i,(L,sec) in enumerate(plan)}
    welcome_idx={n:idx[(n,'welcome')] for n in L_nums}
    codes=[ccss_map[n] for n in L_nums if ccss_map[n]]
    crange=f"{min(codes)} – {max(codes)}" if codes else ""
    MT={'grade':grade,'module':module,'title':args.title or f"Grade {grade} · Module {module}",
        'topics':topics,'lesson_ccss':ccss_map,'welcome_idx':welcome_idx,'ccss_range':crange}
    print("Building slides…")
    prs=Presentation(); prs.slide_width=Emu(9144000); prs.slide_height=Emu(5143500)
    for (L,sec) in plan:
        if sec=='toc': b_toc(prs,MT)
        elif sec=='welcome':
            has=L in sort_cells
            chip={'Math Talk':idx[(L,'mtr')],'DARE':idx[(L,'dareroutine')],'Game':idx[(L,'game')]}
            if has: chip['Randomizer']=idx[(L,'rand')]; chip['Sort']=idx[(L,'sort')]
            b_welcome(prs,MT,L,has,chip)
        elif sec=='mtr': b_mt_routine(prs,MT,L)
        elif sec.startswith('mtp'): b_mt(prs,MT,L,mt_by[L][int(sec[3:])])
        elif sec=='sort': b_sort(prs,MT,L,sort_cells.get(L))
        elif sec=='rand': b_randroutine(prs,MT,L)
        elif sec=='dareroutine': b_dare_routine(prs,MT,L,q_map.get(L,''),w_map.get(L,''))
        elif sec=='dareguide': b_dareguide(prs,MT,L,ag_map.get(L))
        elif sec=='dareedit': b_dareedit(prs,MT,L,ag_map.get(L))
        elif sec=='game': b_game(prs,MT,L)
    raw=os.path.join(tmp,'raw.pptx'); prs.save(raw)
    print("Wiring timers & links…"); postprocess(raw,args.out); print("DONE →",args.out)

# ---------- post-process: timers + links ----------
TIMING='''<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="5" presetID="22" presetClass="exit" presetSubtype="0" fill="hold" grpId="0" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="{D}"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{S}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set><p:animEffect transition="out" filter="wipe(right)"><p:cBhvr><p:cTn id="7" dur="{D}"/><p:tgtEl><p:spTgt spid="{S}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="{S}" grpId="0"/></p:bldLst></p:timing>'''
REL='<Relationship Id="{r}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="{t}"/>'

def postprocess(src,out):
    import zipfile
    work=tempfile.mkdtemp(); unp=os.path.join(work,'u'); os.makedirs(unp)
    with zipfile.ZipFile(src) as z: z.extractall(unp)
    pres=open(f'{unp}/ppt/presentation.xml').read(); rels=open(f'{unp}/ppt/_rels/presentation.xml.rels').read()
    r2t={m.group(1):m.group(2) for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="slides/([^"]+)"',rels)}
    order=[r2t[m.group(1)] for m in re.finditer(r'<p:sldId[^>]*r:id="(rId\d+)"',pres)]
    pos2file={i+1:fn for i,fn in enumerate(order)}
    for pos,fn in pos2file.items():
        sp=f'{unp}/ppt/slides/{fn}'; xml=open(sp).read(); relp=f'{unp}/ppt/slides/_rels/{fn}.rels'
        rl=open(relp).read() if os.path.exists(relp) else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        mt=re.search(r'name="TimerBar::(\d+)"',xml)
        if mt:
            spid=re.search(r'<p:cNvPr id="(\d+)" name="TimerBar::\d+"',xml).group(1)
            xml=xml.replace('</p:sld>',TIMING.replace('{D}',str(int(mt.group(1))*1000)).replace('{S}',spid)+'</p:sld>',1)
        ex=[int(x) for x in re.findall(r'Id="rId(\d+)"',rl)]; nxt=[max(ex)+1 if ex else 1]; t2r={}; newr=[]
        def addlink(tp):
            tf=pos2file[tp]
            if tf in t2r: return t2r[tf]
            r=f'rId{nxt[0]}'; nxt[0]+=1; newr.append(REL.format(r=r,t=tf)); t2r[tf]=r; return r
        def repl(m):
            i,name=m.group(1),m.group(2); tp=int(name.split('::')[1])
            if tp==pos: return m.group(0)
            return f'<p:cNvPr id="{i}" name="{name}"><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{addlink(tp)}" action="ppaction://hlinksldjump"/></p:cNvPr>'
        xml=re.sub(r'<p:cNvPr id="(\d+)" name="(GOTO::\d+)"\s*/>',repl,xml)
        if newr: rl=rl.replace('</Relationships>',''.join(newr)+'</Relationships>'); open(relp,'w').write(rl)
        open(sp,'w').write(xml)
    if os.path.exists(out): os.remove(out)
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        for root,_,files in os.walk(unp):
            for f in files:
                full=os.path.join(root,f); z.write(full,os.path.relpath(full,unp))

if __name__=="__main__":
    ap=argparse.ArgumentParser()
    ap.add_argument('--mathtalks',required=True); ap.add_argument('--sorts',required=True)
    ap.add_argument('--dares',required=True); ap.add_argument('--answerguides',required=True)
    ap.add_argument('--out',required=True); ap.add_argument('--title',default=None)
    ap.add_argument('--topics',default=None); ap.add_argument('--grade',type=int,default=None)
    ap.add_argument('--module',type=int,default=None); ap.add_argument('--assets',default=os.path.join(HERE,'assets'))
    build(ap.parse_args())
