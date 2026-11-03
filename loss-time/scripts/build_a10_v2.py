import openpyxl,datetime as dt
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
import byarea as B
T=lambda h,m=0:dt.time(h,m);D=lambda d:dt.datetime(2026,10,d)
# (ngày, khu vực báo cáo, mã máy, tên máy, hiện tượng, từ, đến, giờ báo cáo, time run, lịch sử)
H0='Họ CNCR: 2 người (30 lần), 3 người (14 lần)'
rep=[
(19,'Machining','CNCR02','Máy CNC','Lỗi nút bấm',T(7,43),T(8,30),0.78,24,'Chưa từng có. '+H0),
(19,'Machining','EDB01','Máy dán cạnh thẳng tự động','Lỗi hộp số',T(18,45),T(19,45),1.00,12,'Chưa từng có. Họ EDB: 2 người (3 lần), 3 người (1 lần)'),
(20,'RC','','Máy bơm sơn chuyền UV','Máy bơm sơn không hoạt động',T(9,1),T(9,58),0.95,12,'Chưa từng có. Chuyền UV khác: 2 người (nhám UV), 4 người (UVDR01, BECY06 ngày 26/6)'),
(20,'Machining','EDB05','Máy dán cạnh thẳng tự động','Lỗi nồi keo',T(18,0),T(21,0),3.00,12,'1 lần: 25/6, 3 người'),
(20,'Machining','CNCR09','Máy chép hình thẳng','Lỗi trục không quay',T(17,50),T(19,50),2.00,12,'Chưa từng có. Máy chép hình tròn 60-0041 từng ghi 2 người (khác loại)'),
(21,'RC','VEC04','Máy cắt veneer','Máy tự cắt',T(18,35),T(20,0),1.42,12,'4 lần (30/12, 20/1, 23/1, 30/7): đều 2 người'),
(22,'Machining','CNCR20','Máy CNC','Gãy giá đỡ dây điện',T(16,49),T(17,40),0.85,12,'3 dòng: 1/6 hai lần 2 người; 4/4 cả ngày 4 người (dòng bất thường)'),
(24,'Machining','CNCR11','Hogmag','Lỗi hút chân không',T(7,40),T(7,53),0.22,12,'1 lần: 9/4, 2 người'),
(25,'RC','VC14','Máy hút chân không','Lỗi hút chân không',T(8,25),T(10,9),1.73,12,'Chưa từng có. Máy hút chân không CNCR14 từng ghi 3 người'),
(26,'RC','EDB01','Máy dán cạnh thẳng tự động','Lỗi không đốt nhiệt độ',T(9,50),T(10,50),1.00,12,'Chưa từng có. Họ EDB: 2 người (3 lần), 3 người (1 lần). Báo cáo ghi khu vực RC, ngày 19/10 cùng máy ghi Machining'),
(26,'PANEL','HPTT02','Máy cảo cong','Lỗi không đốt nhiệt độ',T(15,15),T(16,15),1.00,12,'Chưa từng có (họ HPTT chưa có). Khu vực PANEL = RC - PANEL'),
(28,'RC','VES04','Máy may veneer','Có tiếng kêu lạ khi hoạt động',T(9,55),T(11,7),1.20,12,'3 lần (17/5, 20/5, 24/7): 2 người x2, 3 người x1. Họ VES: 2 người (17 lần)'),
(28,'Machining','CNCR20','Máy CNC 5 trục','Lỗi chương trình',T(12,27),T(14,0),1.55,12,'3 dòng: 1/6 hai lần 2 người; 4/4 cả ngày 4 người (dòng bất thường)'),
(29,'Machining','EDB05','Máy dán cạnh thẳng tự động','Lỗi nồi keo',T(8,35),T(9,50),1.25,12,'1 lần: 25/6, 3 người. Dòng 20/10 cùng máy cũng đang chờ số người'),
(29,'RC','VES07','Máy may veneer','Không hoạt động',T(15,36),T(16,4),0.47,12,'3 lần (5/3, 6/3, 7/5): đều 2 người'),
(29,'Machining','VC','Máy hút chân không','',T(14,54),T(16,3),1.15,12,'Chưa từng có. Báo cáo không ghi hiện tượng'),
(29,'Machining','CUFS02','Máy cắt','Kêu tiếng lớn khi hoạt động',T(17,50),T(19,15),1.42,12,'1 lần: 20/1, 2 người. Họ CUFS: 2 người (6 lần)'),
(29,'Machining','','Máy CNC Nesting','Lỗi nút dừng khẩn cấp',T(18,12),T(18,57),0.75,12,'Chưa từng có. Họ CNCR: 2 người (30 lần), 3 người (14 lần)'),
(30,'RC','VES07','Máy may veneer','Máy bị xì hơi',T(19,3),T(19,14),0.18,12,'3 lần (5/3, 6/3, 7/5): đều 2 người'),
]
# WBS20 từ ảnh trước (không có trong ảnh mới): hỏng 28/10 15:15 đến 30/10 15:15, tính trong khung 7:30-20:00
wbs=[(28,T(15,15),T(20,0),4.75),(29,T(7,30),T(20,0),12.5),(30,T(7,30),T(15,15),7.75)]
HW='Chưa từng có WBS20. Họ WBS (máy nhám thùng): 2 người (12 lần)'
ar_map={'RC':'Raw Cutting','PANEL':'Raw Cutting','Machining':'Machining'}
rows=[]
for d,ar,mc,mn,fail,fr,to,h,n,hs in rep:
    notes='-'.join(x for x in (mc,mn,fail) if x)
    rows.append((D(d),'A10',notes,ar_map[ar],'MAINTENANCE',fr,to,None))
for d,fr,to,h in wbs: rows.append((D(d),'A10','WBS20-Máy nhám thùng-Bị hư','Machining','MAINTENANCE',fr,to,None))
B.ORDER.insert(0,'A10')
wb=openpyxl.Workbook();wb.remove(wb.active)
F=B.F;bd=B.bd
for a in ('Machining','Raw Cutting'):
    ws=B.sheet(wb,a,[x for x in rows if x[3]==a])
    ws['A2']='WORKER QTY ĐỂ TRỐNG: chỉ điền SỐ CÔNG NHÂN THẬT đứng máy khi có xác nhận của tổ trưởng khu vực (ghi tên người xác nhận ở cột K). Không dùng cột "Time run" (12/24) của báo cáo bảo trì, đó là giờ máy chạy mỗi ngày. '+ws['A2'].value
    ws['A2'].font=Font(name=F,italic=True,color='C00000')
    ws.cell(4,11,'Người xác nhận số công nhân').fill=B.gry;ws.cell(4,11).font=Font(name=F,bold=True);ws.cell(4,11).alignment=B.ctr;ws.cell(4,11).border=bd;ws.column_dimensions['K'].width=26
    for r_ in range(5,ws.max_row+1):
        if isinstance(ws.cell(r_,1).value,dt.datetime):
            ws.cell(r_,8).fill=PatternFill('solid',fgColor='FFFF00');ws.cell(r_,11).border=bd
ws=wb.create_sheet('Đối chiếu bảo trì')
ws['A1']='BÁO CÁO BẢO TRÌ 19/10-30/10 VÀ LỊCH SỬ A10 CÙNG MÁY';ws['A1'].font=Font(name=F,bold=True,size=13)
HH=['Ngày','Khu vực (báo cáo)','Mã máy','Tên máy','Hiện tượng','Từ','Đến','Giờ dừng báo cáo','Giờ dừng tính lại','Chênh lệch','Time run (giờ máy chạy/ngày, KHÔNG phải số người)','Lịch sử A10 cùng máy (CHỈ THAM KHẢO)']
for i,h in enumerate(HH,1):
    c=ws.cell(3,i,h);c.fill=B.red;c.font=B.hf;c.alignment=B.ctr;c.border=bd
ws.row_dimensions[3].height=48
r=4
for d,ar,mc,mn,fail,fr,to,h,n,hs in rep:
    for c,v in enumerate([D(d),ar,mc,mn,fail,fr,to,h,f'=(G{r}-F{r})*24',f'=ROUND(I{r}-H{r},2)',n,hs],1):
        x=ws.cell(r,c,v);x.font=Font(name=F,color='7F7F7F' if c==12 else '000000');x.border=bd;x.alignment=Alignment(horizontal='left' if c in(4,5,12) else 'center',vertical='center',wrap_text=c in(4,5,12))
    ws.cell(r,1).number_format='dd/mm/yyyy';ws.cell(r,6).number_format=ws.cell(r,7).number_format='hh:mm';ws.cell(r,9).number_format='0.00'
    r+=1
for d,fr,to,h in wbs:
    for c,v in enumerate([D(d),'Machining','WBS20','Máy nhám thùng','Bị hư (hỏng 28/10 15:15, sửa xong 30/10 15:15; ảnh trước)',fr,to,h,f'=(G{r}-F{r})*24',f'=ROUND(I{r}-H{r},2)',None,HW],1):
        x=ws.cell(r,c,v);x.font=Font(name=F,color='7F7F7F' if c==12 else '000000');x.border=bd;x.alignment=Alignment(horizontal='left' if c in(4,5,12) else 'center',vertical='center',wrap_text=c in(4,5,12))
        x.fill=PatternFill('solid',fgColor='FFF2CC')
    ws.cell(r,1).number_format='dd/mm/yyyy';ws.cell(r,6).number_format=ws.cell(r,7).number_format='hh:mm';ws.cell(r,9).number_format='0.00'
    r+=1
ws.cell(r+1,8,'Tổng').font=Font(name=F,bold=True);ws.cell(r+1,9,f'=SUM(I4:I{r-1})').font=Font(name=F,bold=True);ws.cell(r+1,9).number_format='0.00'
ws.cell(r+3,1,'Vàng = WBS20 (không có trong ảnh báo cáo mới 19-30/10, giữ theo ảnh trước; giờ dừng tính trong khung chạy máy 7:30-20:00 = 25 giờ). 25/10 là Chủ nhật có sự cố VC14 ở RC. EDB01 ngày 26/10 ghi khu vực RC, các lần trước ghi Machining.').font=Font(name=F,italic=True)
for col,w in zip('ABCDEFGHIJKL',[12,12,10,28,34,8,8,12,12,10,22,56]): ws.column_dimensions[col].width=w
wb.save('Nhap_lieu_A10_19-10_den_30-10_NHAP.xlsx');print('rows',len(rows),'bao cao',len(rep))
