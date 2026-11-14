"""Thêm S32 tuần W46 (từ 9/11) vào bảng S32: Beam saw bản (5) và UV phát sinh bản (2). Chạy trong outputs.
python patch_s32_w46.py <file cũ> <file mới>"""
import openpyxl,copy,re,sys,datetime as dt,warnings
warnings.simplefilter('ignore')
OLD,NEW=sys.argv[1:3]
D=lambda d:dt.datetime(2026,11,d);T=lambda h,m=0:dt.time(h,m)
CAM=openpyxl.styles.PatternFill('solid',fgColor='FCE4D6')
# (ngày, mã hàng, qty, khu vực, bộ phận, lí do, ghi chú, from, to, hc, nguồn, fill)
BS='Beam saw - phát sinh (chi tiết {})'
rows=[
(D(10),'MH635F798',20,'Raw Cutting','TECHNICAL FINISHING',BS.format('P01-2'),'Charge to: Technical (bổ sung)',T(15,0),T(15,10),3,'File Beam saw',None),
(D(10),'MH635F798',20,'Raw Cutting','TECHNICAL FINISHING',BS.format('P01-3'),'Charge to: Technical (bổ sung)',T(15,10),T(15,16),3,'File Beam saw',None),
(D(10),'RS16-1087-1641',3,'Raw Cutting','TECHNICAL FINISHING',BS.format('A1-S05'),'Charge to: Technical (bổ sung)',T(16,0),T(16,6),3,'File Beam saw',None),
(D(11),'W3772F484',60,'Raw Cutting','TECHNICAL FINISHING',BS.format('A1-S01-4'),'Charge to: HÀNG BẢNG MÀU',T(8,0),T(8,10),3,'File Beam saw',None),
(D(12),'W3877F2015',100,'Raw Cutting','TECHNICAL FINISHING',BS.format('P01'),'Charge to: HÀNG BẢNG MÀU',T(13,0),T(13,9),3,'File Beam saw',None),
(D(12),'W2804F1250',30,'Raw Cutting','TECHNICAL FINISHING',BS.format('P1'),'Charge to: HÀNG BẢNG MÀU',T(13,9),T(13,15),3,'File Beam saw',None),
(D(12),'W4075F1527',50,'Raw Cutting','TECHNICAL FINISHING',BS.format('P1'),'Charge to: HÀNG BẢNG MÀU',T(13,15),T(13,23),3,'File Beam saw',None),
(D(12),'RH12-1009-0040-R000+RH14-1130-0040-R000+RH18-1008-0040-R000',55,'Raw Cutting','QA/QC/TESTING','Phát sinh khu vực UV - sửa veneer sọc (3 mã hàng)','Charge to: QC A HIẾU; mã hàng xem sheet Chi tiết. 3 mã cùng ca (giờ chỉ ghi ở dòng đầu)',T(14,0),T(15,0),4,'File UV (phát sinh)',None),
(D(13),'AT12-1140-1056-R000',14,'Raw Cutting','QA/QC/TESTING','Phát sinh khu vực UV - sửa veneer sọc','Charge to: QC A HIẾU',T(8,0),T(11,0),4,'File UV (phát sinh)',None),
(D(13),'RH11-8245-2510-R000',16,'Raw Cutting','(cần xác nhận)','Phát sinh khu vực UV - chạy sai màu (3 dòng, 1 mã)','Charge to: IE (chưa có trong dropdown bộ phận, cần xác nhận)',T(13,0),T(13,45),4,'File UV (phát sinh)','cam'),
]
det=[(D(12),'File UV (phát sinh)','RH12-1009-0040-R000','Drawer',15),(D(12),'File UV (phát sinh)','RH14-1130-0040-R000','Drawer',26),(D(12),'File UV (phát sinh)','RH18-1008-0040-R000','Drawer',14),
(D(13),'File UV (phát sinh)','AT12-1140-1056-R000','Drawer',14),(D(13),'File UV (phát sinh)','RH11-8245-2510-R000','Platform',4),(D(13),'File UV (phát sinh)','RH11-8245-2510-R000','Platform',8),(D(13),'File UV (phát sinh)','RH11-8245-2510-R000','Platform',4)]
def fix(v,row):
    if not(isinstance(v,str) and v.startswith('=')): return v
    parts=re.split(r"('[^']*')",v)
    return ''.join(p if p.startswith("'") else re.sub(r'(?<![A-Za-z0-9$])(\$?[A-Z]{1,2})(\d+)\b',lambda m:m.group(1)+str(row),p) for p in parts)
wb=openpyxl.load_workbook(OLD)
ws=next(w for w in wb if w.title.startswith('S32 5-10'));SO=ws.title;SN='S32 5-10 đến 13-11';ws.title=SN
last=next(r for r in range(3,ws.max_row+1) if ws.cell(r,16).value=='Tổng')-1
tail=[[ws.cell(r,c).value for c in range(1,ws.max_column+1)] for r in range(last+1,ws.max_row+1)]
tsty=[[copy.copy(ws.cell(r,c)._style) for c in range(1,ws.max_column+1)] for r in range(last+1,ws.max_row+1)]
ref=last
N=len(rows);nl=last+N
for i,(d,mh,q,ar,dep,li,gc,fr,to,hc,src,fl) in enumerate(rows):
    r=last+1+i
    for c in range(1,19): ws.cell(r,c)._style=copy.copy(ws.cell(ref,c)._style)
    vals={1:f'="W"&WEEKNUM(B{r},21)',2:d,3:mh,4:None,5:q,6:None,7:ar,8:'S32',9:'Other Support of Production',10:dep,11:li,12:gc,13:fr,14:to,15:f'=(N{r}-M{r})*24',16:hc,17:f'=P{r}*O{r}',18:src}
    for c,v in vals.items(): ws.cell(r,c).value=v
    ws.cell(r,2).number_format=ws.cell(ref,2).number_format
    if fl: ws.cell(r,13).fill=CAM;ws.cell(r,14).fill=CAM;ws.cell(r,10).fill=CAM
    else:
        for c in (13,14,10): ws.cell(r,c).fill=copy.copy(ws.cell(ref,c).fill) if False else openpyxl.styles.PatternFill(fill_type=None)
for k,(row,st) in enumerate(zip(tail,tsty)):
    r=nl+1+k
    for c in range(1,len(row)+1): ws.cell(r,c)._style=copy.copy(st[c-1]);ws.cell(r,c).value=row[c-1]
ws.cell(nl+1,17).value=f'=SUM(Q3:Q{nl})'
def upd(v):
    if isinstance(v,str) and v.startswith('='):
        v=v.replace(f"'{SO}'",f"'{SN}'");v=re.sub(r'(\$[A-Z]+\$3:\$[A-Z]+\$)%d'%last,lambda m:m.group(1)+str(nl),v)
    return v
for w in wb:
    for row in w.iter_rows():
        for c in row: c.value=upd(c.value)
# Nhập liệu S32
nlw=wb['Nhập liệu S32'];first=next(r for r in range(2,nlw.max_row+1) if not nlw.cell(r,1).value)
notev=nlw.cell(first+1,1).value;ns=[copy.copy(nlw.cell(first-1,c)._style) for c in range(1,13)];nn=[copy.copy(nlw.cell(first+1,c)._style) for c in range(1,13)]
for c in range(1,13): nlw.cell(first+1,c).value=None
for i,(d,mh,q,ar,dep,li,gc,fr,to,hc,src,fl) in enumerate(rows):
    r=first+i
    for c,v in enumerate([d,'S32',li,ar,dep,fr,to,hc,mh,q,gc,src],1): nlw.cell(r,c)._style=copy.copy(ns[c-1]);nlw.cell(r,c).value=v
    if fl: nlw.cell(r,5).fill=CAM
nr=first+N+1
for c in range(1,13): nlw.cell(nr,c)._style=copy.copy(nn[c-1])
nlw.cell(nr,1).value=notev
# Theo ngày: thêm 10-13/11 và tuần W46
th=wb['Theo ngày'];lastday=next(r for r in range(4,th.max_row+1) if th.cell(r,1).value=='Tổng')-1
wn={10:'Thứ Ba',11:'Thứ Tư',12:'Thứ Năm',13:'Thứ Sáu'}
th.insert_rows(lastday+1,4)
for i,dd in enumerate((10,11,12,13)):
    r=lastday+1+i
    for c in range(1,16): th.cell(r,c)._style=copy.copy(th.cell(lastday,c)._style);th.cell(r,c).value=fix(th.cell(lastday,c).value,r)
    th.cell(r,1).value=D(dd);th.cell(r,2).value=wn[dd]
for r in range(lastday+5,th.max_row+1):
    for c in range(1,16):
        v=th.cell(r,c).value
        if isinstance(v,str) and v.startswith('='):
            v=re.sub(r'SUM\(([A-Z]+)4:([A-Z]+)%d\)'%lastday,lambda m:f'SUM({m.group(1)}4:{m.group(2)}{lastday+4})',v);v=re.sub(r',A\d+\)$',f',A{r})',v);th.cell(r,c).value=v
w45=next(r for r in range(lastday+5,th.max_row+1) if th.cell(r,1).value=='W45')
for c in range(1,16): th.cell(w45+1,c)._style=copy.copy(th.cell(w45,c)._style);th.cell(w45+1,c).value=th.cell(w45,c).value
th.cell(w45+1,1).value='W46';th.cell(w45+1,2).value=re.sub(r',A\d+\)$',f',A{w45+1})',th.cell(w45,2).value)
# Đối chiếu
dc=wb['Đối chiếu nguồn']
for r in range(4,16):
    if dc.cell(r,1).value=='File Beam saw': dc.cell(r,2).value=float(dc.cell(r,2).value)+2.75
    if dc.cell(r,1).value=='File UV (phát sinh)': dc.cell(r,2).value=float(dc.cell(r,2).value)+19
ct=wb['Chi tiết mã hàng'];lr=ct.max_row
for i,x in enumerate(det):
    for c,v in enumerate(x,1): ct.cell(lr+1+i,c)._style=copy.copy(ct.cell(lr,c)._style);ct.cell(lr+1+i,c).value=v
gc=wb['Ghi chú review'];g=gc.max_row;gc.cell(g+1,1)._style=copy.copy(gc.cell(g,1)._style)
gc.cell(g+1,1).value=('17. Tuần W46 (từ 9/11): nguồn nhận được là Beam saw bản (5) và UV phát sinh bản (2). Beam saw: 10/11 (3 dòng, Charge to Technical, ghi chú "bổ sung" trong nguồn), 11/11 (1 dòng) và 12/11 (3 dòng) hàng bảng màu, bộ phận TECHNICAL FINISHING (Technical ghi giống hàng bảng màu, cần xác nhận). UV phát sinh: 12/11 sửa veneer sọc (3 mã cùng ca, 4 người 14:00-15:00 = 4 giờ công), 13/11 sửa veneer sọc AT12-1140-1056-R000 (4 người 8:00-11:00 = 12 giờ công) và chạy sai màu RH11-8245-2510-R000 (4 người 13:00-13:45 = 3 giờ công), bộ phận "IE" chưa có trong dropdown nên ghi (cần xác nhận), tô cam. Chưa có dòng nào W46 trong Veneer và Ép nóng (file chỉ đến 7/11), báo cáo UV ngày chỉ đến 8/11; chưa nhận hình Assembly, hàng sửa MAC, hàng làm lại RC WOOD.')
wb.save(NEW);print('đã thêm',N,'dòng; dòng cuối',nl)
