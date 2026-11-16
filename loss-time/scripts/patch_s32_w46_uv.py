"""Thêm UV báo cáo ngày W46 (9-14/11) vào bảng S32 và sắp lại các dòng W46 theo ngày/giờ. Chạy trong outputs.
python patch_s32_w46_uv.py <file cũ> <file mới>"""
import openpyxl,copy,re,sys,datetime as dt,warnings
warnings.simplefilter('ignore')
OLD,NEW=sys.argv[1:3]
D=lambda d:dt.datetime(2026,11,d);T=lambda h,m=0:dt.time(h,m)
uv={9:(10,8),10:(10,8),11:(10,8),12:(11,9),13:(11,11),14:(11,8)}   # (HC làm việc, HC OT) theo báo cáo UV, tổng 712 giờ
def fix(v,row):
    if not(isinstance(v,str) and v.startswith('=')): return v
    parts=re.split(r"('[^']*')",v)
    return ''.join(p if p.startswith("'") else re.sub(r'(?<![A-Za-z0-9$])(\$?[A-Z]{1,2})(\d+)\b',lambda m:m.group(1)+str(row),p) for p in parts)
wb=openpyxl.load_workbook(OLD)
ws=next(w for w in wb if w.title.startswith('S32 5-10'));SO=ws.title;SN='S32 5-10 đến 14-11';ws.title=SN
last=next(r for r in range(3,ws.max_row+1) if ws.cell(r,16).value=='Tổng')-1
first=next(r for r in range(3,last+1) if ws.cell(r,2).value>=D(9))
blk=[]
for r in range(first,last+1):
    blk.append({c:ws.cell(r,c).value for c in range(1,19)}|{'fill':[copy.copy(ws.cell(r,c).fill) for c in (10,13,14)]})
tmpl=[copy.copy(ws.cell(first,c)._style) for c in range(1,19)]
for d,(w,o) in uv.items():
    for (fr,to,hc,gc) in ((T(7,30),T(15,30),w,'HC làm việc (ca thường)'),(T(15,30),T(19,30),o,'HC tăng ca (OT 4 giờ)')):
        blk.append({1:None,2:D(d),3:None,4:None,5:None,6:None,7:'Raw Cutting',8:'S32',9:'Other Support of Production',10:'FINISHING',11:'Không có bomtime chuyền UV',12:gc,13:fr,14:to,15:None,16:hc,17:None,18:'File UV (báo cáo ngày)','fill':None})
src_order={'File UV (báo cáo ngày)':0,'File Beam saw':1,'File UV (phát sinh)':2}
blk.sort(key=lambda b:(b[2],b[13]))
tail=[[ws.cell(r,c).value for c in range(1,ws.max_column+1)] for r in range(last+1,ws.max_row+1)]
tsty=[[copy.copy(ws.cell(r,c)._style) for c in range(1,ws.max_column+1)] for r in range(last+1,ws.max_row+1)]
nl=first+len(blk)-1
for i,b in enumerate(blk):
    r=first+i
    for c in range(1,19): ws.cell(r,c)._style=copy.copy(tmpl[c-1]);ws.cell(r,c).value=b[c]
    ws.cell(r,1).value=f'="W"&WEEKNUM(B{r},21)';ws.cell(r,15).value=f'=(N{r}-M{r})*24';ws.cell(r,17).value=f'=P{r}*O{r}'
    from openpyxl.styles import PatternFill
    if b['fill']:
        for c,f in zip((10,13,14),b['fill']): ws.cell(r,c).fill=f
    else:
        for c in (10,13,14): ws.cell(r,c).fill=PatternFill(fill_type=None)
for k,(row,st) in enumerate(zip(tail,tsty)):
    r=nl+1+k
    for c in range(1,len(row)+1): ws.cell(r,c)._style=copy.copy(st[c-1]);ws.cell(r,c).value=row[c-1]
ws.cell(nl+1,17).value=f'=SUM(Q3:Q{nl})'
def upd(v):
    if isinstance(v,str) and v.startswith('='):
        v=v.replace("'S32 5-10 đến 13-11'",f"'{SN}'");v=re.sub(r'(\$[A-Z]+\$3:\$[A-Z]+\$)%d'%last,lambda m:m.group(1)+str(nl),v)
    return v
for w in wb:
    for row in w.iter_rows():
        for c in row: c.value=upd(c.value)
# Nhập liệu S32: thêm 12 dòng UV
nlw=wb['Nhập liệu S32'];f0=next(r for r in range(2,nlw.max_row+1) if not nlw.cell(r,1).value)
notev=nlw.cell(f0+1,1).value;ns=[copy.copy(nlw.cell(f0-1,c)._style) for c in range(1,13)];nn=[copy.copy(nlw.cell(f0+1,c)._style) for c in range(1,13)]
for c in range(1,13): nlw.cell(f0+1,c).value=None
n=0
for d,(w,o) in uv.items():
    for (fr,to,hc,gc) in ((T(7,30),T(15,30),w,'HC làm việc (ca thường)'),(T(15,30),T(19,30),o,'HC tăng ca (OT 4 giờ)')):
        r=f0+n;n+=1
        for c,v in enumerate([D(d),'S32','Không có bomtime chuyền UV','Raw Cutting','FINISHING',fr,to,hc,None,None,gc,'File UV (báo cáo ngày)'],1): nlw.cell(r,c)._style=copy.copy(ns[c-1]);nlw.cell(r,c).value=v
nr=f0+n+1
for c in range(1,13): nlw.cell(nr,c)._style=copy.copy(nn[c-1])
nlw.cell(nr,1).value=notev
# Theo ngày: thêm 9/11 và 14/11; ghi lại toàn bộ cột ngày theo thứ tự và công thức theo dòng
th=wb['Theo ngày'];ld=next(r for r in range(4,th.max_row+1) if th.cell(r,1).value=='Tổng')-1
dates=sorted([th.cell(r,1).value for r in range(4,ld+1)]+[D(9),D(14)])
fref={c:th.cell(4,c).value for c in range(1,16)}
sty=[copy.copy(th.cell(ld,c)._style) for c in range(1,16)]
th.insert_rows(ld+1,2)
WN=['Thứ Hai','Thứ Ba','Thứ Tư','Thứ Năm','Thứ Sáu','Thứ Bảy','Chủ nhật']
for i,dd in enumerate(dates):
    r=4+i
    for c in range(1,16):
        th.cell(r,c)._style=copy.copy(sty[c-1])
        th.cell(r,c).value=fix(fref[c],r) if c>2 else None
    th.cell(r,1).value=dd;th.cell(r,2).value=WN[dd.weekday()]
ld=4+len(dates)-1
for r in range(ld+1,th.max_row+1):
    for c in range(1,16):
        v=th.cell(r,c).value
        if isinstance(v,str) and v.startswith('='):
            v=re.sub(r'SUM\(([A-Z]+)4:([A-Z]+)\d+\)',lambda m:f'SUM({m.group(1)}4:{m.group(2)}{ld})',v)
            if th.cell(r,1).value and str(th.cell(r,1).value).startswith('W'): v=re.sub(r',A\d+\)$',f',A{r})',v)
            th.cell(r,c).value=v
dc=wb['Đối chiếu nguồn']
for r in range(4,16):
    if dc.cell(r,1).value=='File UV (báo cáo ngày)': dc.cell(r,2).value=float(dc.cell(r,2).value)+712
gc=wb['Ghi chú review'];g=gc.max_row;gc.cell(g+1,1)._style=copy.copy(gc.cell(g,1)._style)
gc.cell(g+1,1).value='18. UV báo cáo ngày W46 (9-14/11, tổng 712 giờ công; 15/11 nghỉ): mỗi ngày 2 dòng (HC làm việc 7:30-15:30 và HC OT 15:30-19:30), thứ Bảy 14/11 cũng 2 dòng theo báo cáo. Dòng UV phát sinh trong file mới giống bản (2), không có dòng mới (12/11 và 13/11).'
wb.save(NEW);print('đã thêm',len(uv)*2,'dòng UV; dòng cuối',nl)
