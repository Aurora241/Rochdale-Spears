"""Thêm 2 báo cáo Assembly W46 vào bảng S32: bắn cát Supcon 13/11 (36 giờ công) và đóng máy + chà nhám hỗ trợ P1 12/11 (48 giờ công).
python patch_s32_w46_assy.py <file cũ> <file mới>  (chạy trong outputs)"""
import openpyxl,copy,re,sys,datetime as dt,warnings
from openpyxl.styles import PatternFill
warnings.simplefilter('ignore')
OLD,NEW=sys.argv[1:3]
D=lambda d:dt.datetime(2026,11,d);T=lambda h,m=0:dt.time(h,m)
YEL=PatternFill('solid',fgColor='FFF2CC')
wb=openpyxl.load_workbook(OLD);ws=wb['S32 5-10 đến 14-11']
last=next(r for r in range(3,ws.max_row+1) if ws.cell(r,16).value=='Tổng')-1
first=next(r for r in range(3,last+1) if ws.cell(r,2).value>=D(9))
blk=[{c:ws.cell(r,c).value for c in range(1,19)}|{'fill':[copy.copy(ws.cell(r,c).fill) for c in (10,13,14)]} for r in range(first,last+1)]
tmpl=[copy.copy(ws.cell(first,c)._style) for c in range(1,19)]
BC='RH3065+RH403+RH937+RH1772+RH812+RH1971+RH1175+RH3350+RH3721'
new=[
 {1:None,2:D(12),3:'MN20-1005',4:None,5:20,6:None,7:'Assembly',8:'S32',9:'Other Support of Production',10:'OTHERS (INVESTIGATE)',11:'Đóng máy và chà nhám hỗ trợ P1 (ghế)',12:'CN - Hải: 2 người x 12 giờ = 24. Nguồn ghi 7:30-20:00 nhưng 12 giờ/người, bảng ghi 7:30-19:30',13:T(7,30),14:T(19,30),15:None,16:2,17:None,18:'Hình: đóng máy và chà nhám P1'},
 {1:None,2:D(12),3:'MN20-1005',4:None,5:None,6:None,7:'Assembly',8:'S32',9:'Other Support of Production',10:'OTHERS (INVESTIGATE)',11:'Đóng máy và chà nhám hỗ trợ P1 (ghế)',12:'ĐM - Tuấn: 3 người x 8 giờ = 24. Nguồn ghi 7:30-16:00 nhưng 8 giờ/người, bảng ghi 7:30-15:30. Số lượng 20 dùng chung với dòng CN, chỉ tính một lần',13:T(7,30),14:T(15,30),15:None,16:3,17:None,18:'Hình: đóng máy và chà nhám P1'},
 {1:None,2:D(13),3:BC,4:None,5:819,6:None,7:'Assembly',8:'S32',9:'Other Support of Production',10:'SUBCON',11:'Bắn cát hỗ trợ Supcon Thành Đạt / Đại Xuyên (chân)',12:'Line SANDLASTING, TE Đế (9 mã hàng, 819 chân). Nguồn ghi 7:00-20:00 nhưng 12 giờ/người (tổng 36), bảng ghi 7:30-19:30',13:T(7,30),14:T(19,30),15:None,16:3,17:None,18:'Hình: bắn cát'},
]
for n in new:n['fill']=[PatternFill(fill_type=None),YEL,YEL]
blk+=new;blk.sort(key=lambda b:(b[2],b[13]))
tail=[[ws.cell(r,c).value for c in range(1,ws.max_column+1)] for r in range(last+1,ws.max_row+1)]
tsty=[[copy.copy(ws.cell(r,c)._style) for c in range(1,ws.max_column+1)] for r in range(last+1,ws.max_row+1)]
nl=first+len(blk)-1
for i,b in enumerate(blk):
    r=first+i
    for c in range(1,19):ws.cell(r,c)._style=copy.copy(tmpl[c-1]);ws.cell(r,c).value=b[c]
    ws.cell(r,1).value=f'="W"&WEEKNUM(B{r},21)';ws.cell(r,15).value=f'=(N{r}-M{r})*24';ws.cell(r,17).value=f'=P{r}*O{r}'
    for c,f in zip((10,13,14),b['fill']):ws.cell(r,c).fill=f
for k,(row,st) in enumerate(zip(tail,tsty)):
    r=nl+1+k
    for c in range(1,len(row)+1):ws.cell(r,c)._style=copy.copy(st[c-1]);ws.cell(r,c).value=row[c-1]
ws.cell(nl+1,17).value=f'=SUM(Q3:Q{nl})'
def upd(v):
    if isinstance(v,str) and v.startswith('='):v=re.sub(r'(\$[A-Z]+\$3:\$[A-Z]+\$)%d'%last,lambda m:m.group(1)+str(nl),v)
    return v
for w in wb:
    for row in w.iter_rows():
        for c in row:c.value=upd(c.value)
# Đối chiếu nguồn: bắn cát +36, thêm dòng nguồn mới (đóng máy + chà nhám P1 = 48)
dc=wb['Đối chiếu nguồn']
for r in range(4,16):
    if dc.cell(r,1).value=='Hình: bắn cát':dc.cell(r,2).value=float(dc.cell(r,2).value)+36
rt=next(r for r in range(4,25) if dc.cell(r,1).value=='Tổng')
dc.insert_rows(rt)
for c in range(1,6):dc.cell(rt,c)._style=copy.copy(dc.cell(rt-1,c)._style)
dc.cell(rt,1).value='Hình: đóng máy và chà nhám P1';dc.cell(rt,2).value=48
dc.cell(rt,3).value=f"=SUMIFS('S32 5-10 đến 14-11'!$Q$3:$Q${nl},'S32 5-10 đến 14-11'!$R$3:$R${nl},A{rt})"
dc.cell(rt,4).value=f'=C{rt}-B{rt}';dc.cell(rt,5).value=f'=IF(ABS(D{rt})<0.01,"Khớp","Lệch")'
tr=rt+1
for c,L in ((2,'B'),(3,'C'),(4,'D')):dc.cell(tr,c).value=f'=SUM({L}4:{L}{rt})'
dc.cell(tr,5).value=f'=IF(ABS(D{tr})<0.01,"Khớp","Lệch")'
# Chi tiết mã hàng
cd=wb['Chi tiết mã hàng'];r0=cd.max_row+1
rows=[(D(13),'Hình: bắn cát',c,'CHÂN',q) for c,q in (('RH3065',40),('RH403',160),('RH937',20),('RH1772',88),('RH812',96),('RH1971',80),('RH1175',162),('RH3350',87),('RH3721',86))]+[(D(12),'Hình: đóng máy và chà nhám P1','MN20-1005','Ghế',20)]
for k,row in enumerate(rows):
    for c,v in enumerate(row,1):cd.cell(r0+k,c)._style=copy.copy(cd.cell(r0-1,c)._style);cd.cell(r0+k,c).value=v
# Nhập liệu S32 (thêm 3 dòng)
nlw=wb['Nhập liệu S32'];f0=next(r for r in range(2,nlw.max_row+1) if not nlw.cell(r,1).value)
notev=nlw.cell(f0+1,1).value;ns=[copy.copy(nlw.cell(f0-1,c)._style) for c in range(1,13)];nn=[copy.copy(nlw.cell(f0+1,c)._style) for c in range(1,13)]
for c in range(1,13):nlw.cell(f0+1,c).value=None
for k,n in enumerate(new):
    r=f0+k
    for c,v in enumerate([n[2],'S32',n[11],'Assembly',n[10],n[13],n[14],n[16],n[3],n[5],n[12],n[18]],1):nlw.cell(r,c)._style=copy.copy(ns[c-1]);nlw.cell(r,c).value=v
    for c in (6,7):nlw.cell(r,c).fill=YEL
nr=f0+len(new)+1
for c in range(1,13):nlw.cell(nr,c)._style=copy.copy(nn[c-1])
nlw.cell(nr,1).value=notev
# Ghi chú
gc=wb['Ghi chú review'];g=gc.max_row;gc.cell(g+1,1)._style=copy.copy(gc.cell(g,1)._style)
gc.cell(g+1,1).value='19. Báo cáo Assembly W46 (17/11): bắn cát Supcon 13/11 (3 người x 12 giờ = 36, 9 mã hàng 819 chân, bộ phận SUBCON) và đóng máy + chà nhám hỗ trợ P1 12/11 (ghế MN20-1005, 20 cái; CN-Hải 2 người x 12 giờ, ĐM-Tuấn 3 người x 8 giờ, tổng 48; bộ phận OTHERS (INVESTIGATE) theo tiền lệ hỗ trợ P1). Giờ bắt đầu/kết thúc chọn để ra đúng số giờ/người của báo cáo (tô vàng): nguồn ghi 7:00-20:00 và 7:30-20:00 nhưng 12 giờ, 7:30-16:00 nhưng 8 giờ. Chờ xác nhận bộ phận chịu trách nhiệm của dòng hỗ trợ P1.'
wb.save(NEW);print('đã thêm 3 dòng; dòng cuối',nl)
