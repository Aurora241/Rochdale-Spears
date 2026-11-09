"""Cập nhật bảng S32 đến 8/11/2026 (tuần W45): thợ sơn, cào cước, bắn cát, Veneer, Ép nóng.
Chạy trong thư mục outputs. Nguồn: 3 hình Assembly tuần 45, Veneer 26.11.07 (bản đầy đủ), FILE PHÒNG SỬA KV ÉP NÓNG."""
import openpyxl,copy,re,warnings,datetime as dt
from openpyxl.styles import PatternFill
warnings.simplefilter('ignore')
OLD='Loss_time_S32_05-10_den_04-11.xlsx';NEW='Loss_time_S32_05-10_den_08-11.xlsx'
wb=openpyxl.load_workbook(OLD)
SO='S32 5-10 đến 04-11';SN='S32 5-10 đến 08-11'
D=lambda d:dt.datetime(2026,11,d);T=lambda h,m=0:dt.time(h,m)
YEL=PatternFill('solid',fgColor='FFF2CC');CAM=PatternFill('solid',fgColor='FCE4D6')
rows=[]  # (ngày, mã hàng, qty, khu vực, bộ phận, lí do, ghi chú, from, to, hc, nguồn, fill)
SON='Không có bomtime sơn lòng trong'
for d,fr,to,tx,fl in ((2,T(7,30),T(19,30),'',None),(3,T(7,30),T(19,30),'',None),(4,T(7,30),T(19,30),'',None),(5,T(7,30),T(19,30),'',None),(6,T(7,30),T(19,30),'',None),(7,T(7,30),T(15,30),' (thứ Bảy 8 giờ theo báo cáo tuần 45)',None),(8,T(7,30),T(15,30),' (Chủ nhật)',CAM)):
    rows.append((D(d),None,None,'Assembly','FINISHING',SON,'Thợ sơn lòng trong tủ - hỗ trợ Fin'+tx,fr,to,12,'Hình: thợ sơn',fl))
rows.append((D(5),'RH7771+RH1773+RH1788 (+3 mã)',917,'Assembly','SUBCON','Bắn cát hỗ trợ Supcon Vượng Phát Lộc / Thành Lộc (chân)','Line bắn cát (7 dòng, 6 mã hàng; RH2251 hai dòng 222+126). Nguồn ghi 7:30-20:00 nhưng 12 giờ/người (tổng 36), bảng ghi 7:30-19:30 để đúng 36 giờ',T(7,30),T(19,30),3,'Hình: bắn cát',YEL))
rows.append((D(6),'RH337F1',None,'Assembly','MACHINING','Cào cước hỗ trợ MAC','Line 8 - hỗ trợ MAC. Nguồn ghi 8:00-16:00 nhưng tổng 7,5 giờ, bảng ghi 8:00-11:45 để đúng 7,5 giờ công',T(8,0),T(11,45),2,'Hình: cào cước',YEL))
rows.append((D(6),'code 119+code 007+code 290',90,'Raw Cutting','TECHNICAL FINISHING','Làm mẫu (veneer)','3 mã (code 119, 007, 290), 1 ca. Nguồn: hàng cắt ghép tốn nhiều time',T(8,0),T(20,0),2,'File Veneer',None))
rows.append((D(7),'W4074F1527EX3-R0000',30,'Raw Cutting','TECHNICAL FINISHING','Ép hàng sửa hàng bảng màu','Hàng bảng màu (nguồn ghi BÀNG MÀU), 1 mã có giờ. 2 mã W2373F1144-EX-R0000 (150) và W2127F1039-IN-R0000 (100) cùng ngày chưa có giờ/số người nên chưa nhập',T(7,0),T(9,30),4,'File Ép nóng',None))
rows.append((D(7),'code 035+code 007',60,'Raw Cutting','TECHNICAL FINISHING','Làm mẫu (veneer)','2 mã (code 035, code 007), 1 ca',T(9,0),T(11,30),2,'File Veneer',None))
rows.append((D(7),'code 119',50,'Raw Cutting','TECHNICAL FINISHING','Làm mẫu (veneer)','code 119',T(13,0),T(13,45),2,'File Veneer',None))
rows.append((D(8),'RH892',None,'Assembly','MACHINING','Cào cước hỗ trợ MAC','Line 8 - hỗ trợ MAC (Chủ nhật). Nguồn ghi 8:00-13:00 nhưng tổng 8 giờ, bảng ghi 8:00-12:00 để đúng 8 giờ công',T(8,0),T(12,0),2,'Hình: cào cước',CAM))
rows.append((D(8),'RH079F3',None,'Assembly','MACHINING','Cào cước hỗ trợ MAC','Line 8 - hỗ trợ MAC (Chủ nhật)',T(13,0),T(16,30),2,'Hình: cào cước',CAM))
rows.sort(key=lambda x:(x[0],{'Hình: thợ sơn':0,'Hình: bắn cát':1,'Hình: cào cước':2,'File Ép nóng':3,'File Veneer':4}[x[10]]))
N=len(rows)
ws=wb[SO];ws.title=SN
def fix(v,row):
    if not(isinstance(v,str) and v.startswith('=')): return v
    parts=re.split(r"('[^']*')",v)
    return ''.join(p if p.startswith("'") else re.sub(r'(?<![A-Za-z0-9$])(\$?[A-Z]{1,2})(\d+)\b',lambda m:m.group(1)+str(row),p) for p in parts)
last=126;tot=127
tail=[[ws.cell(r,c).value for c in range(1,ws.max_column+1)] for r in range(tot,ws.max_row+1)]
tsty=[[copy.copy(ws.cell(r,c)._style) for c in range(1,ws.max_column+1)] for r in range(tot,ws.max_row+1)]
ref=124
for i,(d,mh,qty,ar,dept,li,gc,fr,to,hc,src,fl) in enumerate(rows):
    r=last+1+i
    for c in range(1,19): ws.cell(r,c)._style=copy.copy(ws.cell(ref,c)._style)
    vals={1:f'="W"&WEEKNUM(B{r},21)',2:d,3:mh,4:None,5:qty,6:None,7:ar,8:'S32',9:'Other Support of Production',10:dept,11:li,12:gc,13:fr,14:to,15:f'=(N{r}-M{r})*24',16:hc,17:f'=P{r}*O{r}',18:src}
    vals[9]=ws.cell(ref,9).value
    for c,v in vals.items(): ws.cell(r,c).value=v
    ws.cell(r,2).number_format=ws.cell(ref,2).number_format
    if fl:
        ws.cell(r,13).fill=fl;ws.cell(r,14).fill=fl
newlast=last+N
for k,(row,st) in enumerate(zip(tail,tsty)):
    r=newlast+1+k
    for c in range(1,len(row)+1):
        ws.cell(r,c)._style=copy.copy(st[c-1]);ws.cell(r,c).value=row[c-1]
for c in range(1,ws.max_column+1):
    for r in range(last+1+N+len(tail),last+1+N+len(tail)+3): pass
ws.cell(newlast+1,17).value=f'=SUM(Q3:Q{newlast})'
ws['A1']=re.sub(r'04/11','08/11',ws['A1'].value) if ws['A1'].value else ws['A1'].value
# cập nhật công thức mọi sheet
def upd(v):
    if isinstance(v,str) and v.startswith('='):
        v=v.replace(f"'{SO}'",f"'{SN}'");v=re.sub(r'(\$[A-Z]+\$3:\$[A-Z]+\$)126',lambda m:m.group(1)+str(newlast),v)
    return v
for w in wb:
    for row in w.iter_rows():
        for c in row: c.value=upd(c.value)
# Nhập liệu S32
nl=wb['Nhập liệu S32']
first=next(r for r in range(2,nl.max_row+1) if not nl.cell(r,1).value)
note=[nl.cell(first+1,c).value for c in range(1,3)]
nsty=[copy.copy(nl.cell(first-1,c)._style) for c in range(1,13)];nstyn=[copy.copy(nl.cell(first+1,c)._style) for c in range(1,13)]
notev=nl.cell(first+1,1).value
for c in range(1,13): nl.cell(first+1,c).value=None
for i,(d,mh,qty,ar,dept,li,gc,fr,to,hc,src,fl) in enumerate(rows):
    r=first+i
    for c,v in enumerate([d,'S32',li,ar,dept,fr,to,hc,mh,qty,gc,src],1):
        nl.cell(r,c)._style=copy.copy(nsty[c-1]);nl.cell(r,c).value=v
    if fl:
        nl.cell(r,6).fill=fl;nl.cell(r,7).fill=fl
nr=first+N+1
for c in range(1,13): nl.cell(nr,c)._style=copy.copy(nstyn[c-1])
nl.cell(nr,1).value=notev
# Theo ngày: bảng có 4/10-4/11 (thiếu 3/11): thêm 3/11, 5/11, 6/11, 7/11, 8/11; dịch bảng tuần xuống
th=wb['Theo ngày']
lastday=32   # dòng 4/11 cũ -> đổi thành 3/11, thêm 5 dòng phía sau
th.insert_rows(lastday+1,5)
wn={3:'Thứ Ba',4:'Thứ Tư',5:'Thứ Năm',6:'Thứ Sáu',7:'Thứ Bảy',8:'Chủ nhật'}
for i,d in enumerate((3,4,5,6,7,8)):
    r=lastday+i
    if i>0:
        for c in range(1,16): th.cell(r,c)._style=copy.copy(th.cell(lastday,c)._style);th.cell(r,c).value=fix(th.cell(lastday,c).value,r)
    th.cell(r,1).value=D(d);th.cell(r,2).value=wn[d]
for r in range(lastday+6,th.max_row+1):
    for c in range(1,16):
        v=th.cell(r,c).value
        if isinstance(v,str) and v.startswith('='):
            v=re.sub(r'SUM\(([A-Z]+)4:([A-Z]+)32\)',lambda m:f'SUM({m.group(1)}4:{m.group(2)}{lastday+5})',v)
            v=re.sub(r',A\d+\)$',f',A{r})',v)
            th.cell(r,c).value=v
# Đối chiếu nguồn
dc=wb['Đối chiếu nguồn']
add={'File Ép nóng':10,'File Veneer':30.5,'Hình: thợ sơn':912,'Hình: cào cước':22.5,'Hình: bắn cát':36}
for r in range(4,15):
    k=dc.cell(r,1).value
    if k in add: dc.cell(r,2).value=round(float(dc.cell(r,2).value)+add[k],4)
# Chi tiết mã hàng
ct=wb['Chi tiết mã hàng'];lr=ct.max_row
det=[(D(5),'Hình: bắn cát','RH7771','CHÂN',296),(D(5),'Hình: bắn cát','RH1773','CHÂN',60),(D(5),'Hình: bắn cát','RH1788','CHÂN',128),(D(5),'Hình: bắn cát','RH2251','CHÂN',222),(D(5),'Hình: bắn cát','RH2676','CHÂN',60),(D(5),'Hình: bắn cát','RH2588','CHÂN',25),(D(5),'Hình: bắn cát','RH2251','CHÂN',126),
 (D(6),'File Veneer','code 119','320x620',30),(D(6),'File Veneer','code 007','320x620',30),(D(6),'File Veneer','code 290','320x620',30),
 (D(7),'File Ép nóng','W4074F1527EX3-R0000','A01-S01',30),(D(7),'File Veneer','code 035','320x620',30),(D(7),'File Veneer','code 007','320x620',30),(D(7),'File Veneer','code 119','320x620',50)]
for i,x in enumerate(det):
    r=lr+1+i
    for c,v in enumerate(x,1):
        ct.cell(r,c)._style=copy.copy(ct.cell(lr,c)._style);ct.cell(r,c).value=v
# Ghi chú review
gc=wb['Ghi chú review'];g=gc.max_row
gc.cell(g+1,1)._style=copy.copy(gc.cell(g,1)._style)
gc.cell(g+1,1).value=('15. Cập nhật tuần W45 (2/11-8/11): thợ sơn 12 người (T2-T6 7:30-19:30, thứ Bảy 7/11 và Chủ nhật 8/11 mỗi ngày 7:30-15:30 = 912 giờ; Chủ nhật tô cam). Cào cước: 6/11 RH337F1, 8/11 RH892 và RH079F3 (2 người mỗi dòng); báo cáo ghi giờ không khớp tổng (8:00-16:00 nhưng tổng 7,5; 8:00-13:00 nhưng tổng 8; 13:00-16:30 tổng 7) nên bảng dùng giờ TỔNG của nguồn (22,5 giờ công, tô vàng/cam). Bắn cát 5/11: 3 người x 12 giờ = 36 giờ (nguồn ghi 7:30-20:00, bảng ghi 7:30-19:30), Supcon Vượng Phát Lộc / Thành Lộc. Veneer: 6/11 code 119+007+290 (2 người 8:00-20:00, 24 giờ), 7/11 code 035+007 (2 người 9:00-11:30, 5 giờ) và code 119 (2 người 13:00-13:45, 1,5 giờ), bộ phận TECHNICAL FINISHING. Ép nóng 7/11: W4074F1527EX3 (4 người 7:00-9:30, 10 giờ); 2 mã cùng ngày không có giờ nên chưa nhập. Veneer 6/11 và 7/11 không có dòng nào cho 30/10-5/11 trong file. File "time tắt đừng máy khu vực panel Mr.Tâm" chỉ có dữ liệu tháng 4-9 nên không dùng.')
wb.save(NEW);print('saved',N,'dòng mới; dòng cuối',newlast)
