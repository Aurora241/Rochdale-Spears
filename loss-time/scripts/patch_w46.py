"""Thêm bảng ngày 9/11 (W46) vào Loss_time_05-10_den_08-11.xlsx -> Loss_time_05-10_den_09-11.xlsx. Chạy trong outputs.
Dùng: python ../scripts/patch_w46.py dd-mm  (ví dụ 09-11), ngày thường có tăng ca, không nhân viên mới."""
import openpyxl,copy,re,sys,warnings,datetime as dt
warnings.simplefilter('ignore')
nm=sys.argv[1];OLD=sys.argv[2];NEW=sys.argv[3];PREV=sys.argv[4]   # ví dụ 09-11, file cũ, file mới, sheet ngày liền trước
day=dt.datetime(2026,int(nm[3:]),int(nm[:2]))
D=lambda d,m:dt.datetime(2026,m,d)
SUN={D(25,10),D(1,11),D(8,11)}
wb=openpyxl.load_workbook(OLD)
def kday(g,area,d,suns=SUN):
    n=0;x=g
    while x<d:
        x+=dt.timedelta(days=1)
        if x.weekday()!=6 or (area=='Assembly' and x in suns): n+=1
    return n
def fix(v,row):
    if not(isinstance(v,str) and v.startswith('=')): return v
    parts=re.split(r"('[^']*')",v)
    return ''.join(p if p.startswith("'") else re.sub(r'(?<![A-Za-z0-9$])(\$?[A-Z]{1,2})(\d+)\b',lambda m:m.group(1)+str(row),p) for p in parts)
def clone(ws,src,dst,tpl):
    for c in range(1,ws.max_column+1): ws.cell(dst,c)._style=copy.copy(tpl.cell(src,c)._style)
nv=wb['Nhân viên mới - A2']
groups=[]
for r in range(4,nv.max_row+1):
    g=nv.cell(r,1).value
    if g and re.match(r'\d\d/\d\d$',str(g)):
        dd,mm=map(int,g.split('/'));groups.append((r,D(dd,mm),nv.cell(r,2).value,nv.cell(r,3).value,g))
AREA=lambda a:'Assembly' if a=='Assembly' else a
# 1) sheet ngày
t=wb['Loss time 06-11'];ws=wb.copy_worksheet(t);ws.title=f'Loss time {nm}'
wb.move_sheet(ws,offset=wb.sheetnames.index(f'Loss time {PREV}')+1-wb.sheetnames.index(ws.title))
ws['A1']=re.sub(r'\d\d/\d\d/2026',day.strftime('%d/%m/%Y'),t['A1'].value)
tr=next(r for r in range(3,t.max_row+1) if t.cell(r,16).value=='Tổng')
a2tpl=next(r for r in range(3,tr) if t.cell(r,8).value=='A2')
ref=[t.cell(a2tpl,c).value for c in range(1,ws.max_column+1)]
for r in range(3,23): ws.cell(r,2).value=day
for r in range(23,ws.max_row+1):
    for c in range(1,ws.max_column+1): ws.cell(r,c).value=None
rows=[]
for (_,g,ar,hc,gtxt) in groups:
    k=kday(g,AREA(ar),day)
    if 1<=k<=12: rows.append((g,ar,hc,k,gtxt))
order={'Assembly':0,'Machining':1};rows.sort(key=lambda x:(x[0],order.get(x[1],2)))
r=23
for (g,ar,hc,k,gtxt) in rows:
    clone(ws,a2tpl,r,t)
    for c,v in enumerate(ref,1): ws.cell(r,c).value=fix(v,r)
    ws.cell(r,2).value=day;ws.cell(r,7).value=ar;ws.cell(r,16).value=hc;ws.cell(r,12).value=f'A2 ngày {k}/12 - nhóm {gtxt}'
    ws.cell(r,11).value=re.sub(r'\d\d/\d\d$',gtxt,ref[10]);r+=1
last=r-1
clone(ws,tr,r,t)
for c in range(1,ws.max_column+1): ws.cell(r,c).value=fix(t.cell(tr,c).value,r)
ws.cell(r,17).value=f'=SUM(Q3:Q{last})'
for c in range(1,ws.max_column+1): ws.cell(r+2,c)._style=copy.copy(t.cell(tr+2,c)._style);ws.cell(r+2,c).value=t.cell(tr+2,c).value
print('A2 rows',[(x[2],x[1],x[3],x[4]) for x in rows],'HC',sum(x[2] for x in rows))
# 2) Tổng hợp: khối mới sau khối ngày trước
th=wb['Tổng hợp'];prev=wb[f'Loss time {PREV}']
src=next(r for r in range(1,th.max_row+1) if th.cell(r,1).value==f'Ngày {dt.datetime(2026,int(PREV[3:]),int(PREV[:2])).strftime("%d/%m/%Y")}')
note_row=next(r for r in range(src,th.max_row+1) if str(th.cell(r,1).value).startswith('Chưa gồm'))
note=th.cell(note_row,1).value;nstyle=copy.copy(th.cell(note_row,1)._style)
top=src+9;shift=top-src
th.cell(note_row,1).value=None
for i in range(7):
    clone(th,src+i,top+i,th)
    for c in range(1,th.max_column+1):
        v=th.cell(src+i,c).value
        if isinstance(v,str) and v.startswith('='):
            parts=re.split(r"('[^']*')",v);out=[]
            for p in parts:
                if p.startswith("'"):out.append(p.replace(f"'Loss time {PREV}'",f"'Loss time {nm}'"));continue
                p=re.sub(r'\$([QHG])\$3:\$\1\$\d+',lambda m:f'${m.group(1)}$3:${m.group(1)}${last}',p)
                p=re.sub(r'(?<![A-Za-z0-9$])(\$?[A-Z]{1,2}\$?)(\d+)\b',lambda m:m.group(1)+str(int(m.group(2))+shift) if int(m.group(2))>=src else m.group(0),p)
                out.append(p)
            v=''.join(out)
        th.cell(top+i,c).value=v
th.cell(top,1).value=f'Ngày {day.strftime("%d/%m/%Y")}'
th.cell(top+9,1).value=note;th.cell(top+9,1)._style=nstyle
# 3) Nhân viên mới - A2: cột ngày mới
c0=nv.max_column
nv.insert_cols(c0,1)
nv.cell(3,c0)._style=copy.copy(nv.cell(3,c0-1)._style);nv.cell(3,c0).value=f'Mã ngày {nm.replace("-","/")}'
for r in range(4,nv.max_row+1):
    g=nv.cell(r,1).value
    if not(g and re.match(r'\d\d/\d\d$',str(g))): continue
    dd,mm=map(int,g.split('/'));gd=D(dd,mm);ar=nv.cell(r,2).value
    c=nv.cell(r,c0);c._style=copy.copy(nv.cell(r,c0-1)._style)
    k=kday(gd,AREA(ar),day);kp=kday(gd,AREA(ar),day-dt.timedelta(days=1))
    c.value='P22' if gd==day else ('Đã xong' if (kp>=12) else ('-' if k==0 else f'A2 ngày {k}'))
wb.save(NEW);print('saved',NEW,'dòng cuối',last)
