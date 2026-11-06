import re,shutil,os,zipfile,copy,sys,json
from lxml import etree
from xml.sax.saxutils import escape as X
from PIL import Image
SRC='x';OUT='y'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns={'w':W}
shutil.rmtree(OUT,ignore_errors=True);shutil.copytree(SRC,OUT)
doc=open(SRC+'/word/document.xml',encoding='utf8').read()
head=doc[:doc.index('<w:body>')+8]
for pfx,u in (('a','http://schemas.openxmlformats.org/drawingml/2006/main'),('pic','http://schemas.openxmlformats.org/drawingml/2006/picture'),('wp','http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')):
    if f'xmlns:{pfx}=' not in head: head=head.replace('<w:document ',f'<w:document xmlns:{pfx}="{u}" ',1)
pages=json.load(open('pages.json')) if os.path.exists('pages.json') else {}
IMG='/tmp/claude-0/-home-user-Rochdale-Spears/64423226-a456-5e96-8225-6a5ffe123efb/images/'
# ---------- images
for f in ('image1.png','image2.jpg','image3.png'):os.remove(f'{OUT}/word/media/{f}')
imgs={}
Image.open(IMG+'59.webp').convert('RGB').save(f'{OUT}/word/media/ncr.png')
shutil.copy(IMG+'60.png',f'{OUT}/word/media/bt.png');shutil.copy(IMG+'61.png',f'{OUT}/word/media/stop.png')
rels=open(OUT+'/word/_rels/document.xml.rels',encoding='utf8').read()
rels=re.sub(r'<Relationship Id="rId[789]"[^>]*/>','',rels)
rels=rels.replace('</Relationships>',''.join(f'<Relationship Id="{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{f}"/>' for i,f in (('rId21','ncr.png'),('rId22','bt.png'),('rId23','stop.png')))+'</Relationships>')
open(OUT+'/word/_rels/document.xml.rels','w',encoding='utf8').write(rels)
# ---------- numbering: restartable numbered lists
nx=open(OUT+'/word/numbering.xml',encoding='utf8').read()
add=''.join(f'<w:num w:numId="{n}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' for n in (20,21,22))
nx=nx.replace('</w:numbering>',add+'</w:numbering>');open(OUT+'/word/numbering.xml','w',encoding='utf8').write(nx)
# ---------- header
h=open(OUT+'/word/header1.xml',encoding='utf8').read()
hp=h.index('<w:p ');hp2=h.index('</w:p>',hp)
pp=re.search(r'<w:pPr>.*?</w:pPr>',h[hp:hp2],re.S)
rp=re.search(r'<w:rPr>.*?</w:rPr>',h[hp:hp2],re.S)
h=h[:hp]+h[hp:h.index('>',hp)+1]+(pp.group(0) if pp else '')+f'<w:r>{rp.group(0) if rp else ""}<w:t>Báo cáo thử việc</w:t></w:r>'+h[hp2:]
open(OUT+'/word/header1.xml','w',encoding='utf8').write(h)
# ---------- helpers
def runs(t,base=''):
    """**bold** markup -> runs"""
    out=''
    for k,seg in enumerate(re.split(r'\*\*',t)):
        if not seg:continue
        rpr=base+('<w:b/><w:bCs/>' if k%2 else '')
        out+=f'<w:r>{"<w:rPr>"+rpr+"</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{X(seg)}</w:t></w:r>'
    return out
BP='<w:pPr><w:spacing w:before="100" w:after="120" w:line="340" w:lineRule="auto"/><w:ind w:firstLine="567"/><w:jc w:val="both"/></w:pPr>'
def para(t):return f'<w:p>{BP}{runs(t)}</w:p>'
bm=[100]; toc=[]
def heading(level,t):
    bm[0]+=1;name=f'_Toc9{bm[0]:05d}';toc.append((level,t,name))
    return f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr><w:bookmarkStart w:id="{bm[0]}" w:name="{name}"/><w:r><w:t>{X(t)}</w:t></w:r><w:bookmarkEnd w:id="{bm[0]}"/></w:p>'
def item(t,num):return f'<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr><w:spacing w:after="80" w:line="320" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>{runs(t)}</w:p>'
def bullet(t):return item(t,3)
PB='<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
def tcap(t):return f'<w:p><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="100"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">{X(t)}</w:t></w:r></w:p>'
def fcap(t):return f'<w:p><w:pPr><w:spacing w:before="60" w:after="200"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/><w:iCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">{X(t)}</w:t></w:r></w:p>'
fid=[10]
def figure(rid,fn,cap):
    w,h=Image.open(f'{OUT}/word/media/{fn}').size;cx=5715000;cy=int(cx*h/w);fid[0]+=1
    return (f'<w:p><w:pPr><w:keepNext/><w:spacing w:before="120"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{fid[0]}" name="{fn}" descr="{X(cap)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')+fcap(cap)
BRD='<w:tcBorders>'+''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>' for s in('top','left','bottom','right'))+'</w:tcBorders>'
MAR='<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/>'
def table(rows,widths,center=(),sz=22,shade=None):
    g=''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    o=('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblBorders>'%sum(widths)+''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>' for s in('top','left','bottom','right','insideH','insideV'))
       +'</w:tblBorders><w:tblCellMar><w:left w:w="10" w:type="dxa"/><w:right w:w="10" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>'+f'<w:tblGrid>{g}</w:tblGrid>')
    for ri,r in enumerate(rows):
        hd=ri==0
        o+='<w:tr><w:trPr>'+('<w:tblHeader/>' if hd else '<w:cantSplit/>')+'</w:trPr>'
        for ci,(c,w) in enumerate(zip(r,widths)):
            fill='<w:shd w:val="clear" w:color="auto" w:fill="1F4E79"/>' if hd else (f'<w:shd w:val="clear" w:color="auto" w:fill="{shade(ri,ci,c)}"/>' if shade and shade(ri,ci,c) else '')
            kn='<w:keepNext/>' if (len(rows)<=8 and ri<len(rows)-1) else ''
            jc=f'<w:pPr>{kn}<w:jc w:val="center"/></w:pPr>' if (hd or ci in center) else (f'<w:pPr>{kn}</w:pPr>' if kn else '')
            base=('<w:b/><w:bCs/><w:color w:val="FFFFFF"/>' if hd else '')+f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
            o+=f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{BRD}{fill}{MAR}</w:tcPr><w:p>{jc}{runs(c,base)}</w:p></w:tc>'
        o+='</w:tr>'
    return o+'</w:tbl>'
def cover(t,b=False):pass
# ---------- cover from template (edit text)
body=etree.fromstring(doc.encode('utf8')).find('w:body',ns)
def settext(p,text):
    rs=p.findall('.//w:r',ns);first=True
    for r in rs:
        tt=r.find('w:t',ns)
        if tt is None:continue
        if first:tt.text=text;tt.set('{http://www.w3.org/XML/1998/namespace}space','preserve');first=False
        else:r.getparent().remove(r)
    for pe in p.findall('w:proofErr',ns):p.remove(pe)
settext(body[1],'THỬ VIỆC (1 THÁNG)')
# bold-label paragraphs (3..6): label run stays; fix value run
def setval(p,val):
    ts=[r for r in p.findall('.//w:r',ns) if r.find('w:t',ns) is not None]
    for r in ts:
        if r.find('w:rPr/w:b',ns) is None:
            r.find('w:t',ns).text=val;return
cv=''.join(etree.tostring(body[i]).decode() for i in range(0,8))
# rebuild cover paragraph 6 properly
p6=body[6]
labels=[r for r in p6.findall('.//w:r',ns) if r.find('w:rPr/w:b',ns) is not None]
nb=[r for r in p6.findall('.//w:r',ns) if r.find('w:rPr/w:b',ns) is None]
for r in nb[1:]:r.getparent().remove(r)
nb[0].find('w:t',ns).text='5/10/2026 – 6/11/2026'
# label text: "Thời gian: "
for r in labels[1:]:r.getparent().remove(r)
labels[0].find('w:t',ns).text='Thời gian: ';labels[0].find('w:t',ns).set('{http://www.w3.org/XML/1998/namespace}space','preserve')
for pe in p6.findall('w:proofErr',ns):p6.remove(pe)
cover_xml=''.join(re.sub(r' xmlns:\w+="[^"]+"','',etree.tostring(body[i]).decode()) for i in range(0,9))
sect_para=re.sub(r' xmlns:\w+="[^"]+"','',etree.tostring(body[11]).decode())
final_sect=re.sub(r' xmlns:\w+="[^"]+"','',etree.tostring(body[102]).decode())
exec(open('content.py',encoding='utf8').read())
# ---------- TOC
def tocp(level,t,name):
    pg=str(pages.get(name,1))
    return (f'<w:p><w:pPr><w:pStyle w:val="TOC{level}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9061"/></w:tabs></w:pPr><w:hyperlink w:anchor="{name}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>{X(t)}</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>{pg}</w:t></w:r></w:hyperlink></w:p>')
toc_xml='<w:sdt><w:sdtPr><w:alias w:val="Mục lục"/><w:id w:val="-386885119"/></w:sdtPr><w:sdtContent>'+''.join(tocp(*t) for t in toc)+'</w:sdtContent></w:sdt>'
full=head+cover_xml+toc_xml+PB+sect_para+BODY+final_sect+'</w:body></w:document>'
# restore namespaces on stripped fragments: parse-check
full=full.replace('<w:p>','<w:p>')
open(OUT+'/word/document.xml','w',encoding='utf8').write(full)
json.dump([(l,t,n) for l,t,n in toc],open('toc.json','w'),ensure_ascii=False)
fn='Bao_cao_thu_viec.docx'
if os.path.exists(fn):os.remove(fn)
with zipfile.ZipFile(fn,'w',zipfile.ZIP_DEFLATED) as z:
    z.write(OUT+'/[Content_Types].xml','[Content_Types].xml')
    for r,_,fs in os.walk(OUT):
        for f in fs:
            p=os.path.join(r,f);a=os.path.relpath(p,OUT)
            if a!='[Content_Types].xml':z.write(p,a)
print('built',len(toc))
