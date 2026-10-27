import openpyxl,datetime as dt
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
src=openpyxl.load_workbook('Nhap_lieu_05-10_den_26-10.xlsx',data_only=True)
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
red=PatternFill('solid',fgColor='C00000');hf=Font(name=F,bold=True,color='FFFFFF');ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
yel=PatternFill('solid',fgColor='FFF2CC');org=PatternFill('solid',fgColor='FCE4D6');grn=PatternFill('solid',fgColor='E2EFDA');gry=PatternFill('solid',fgColor='D9D9D9')
D=lambda d,m=10:dt.datetime(2026,m,d)
def rows(sh):
    return [list(r[:7]) for r in src[sh].iter_rows(min_row=2,values_only=True) if r[0]]
adds=[]   # (date,code,notes,process,dept,from,to,hc,reason)
for r in rows('Nhập 19-10'):
    if r[0]=='P22': adds.append((D(19),*r,'Thiếu P22 nhân viên mới ngày 19/10'))
for r in rows('Nhập 20-10'):
    if r[0]=='A2' and r[1].endswith('19/10'): adds.append((D(20),*r,'Thiếu A2 ngày 1 của nhóm 19/10'))
for r in rows('Nhập 21-10'):
    if r[0]=='P22' and r[2] in('Machining','Raw Cutting'): adds.append((D(21),*r,'Thiếu P22 nhân viên mới ngày 21/10'))
for r in rows('Nhập 23-10'):
    if r[0]=='P22': adds.append((D(23),*r,'Thiếu P22 nhân viên mới ngày 23/10 (Raw Cutting 1 người)'))
for r in rows('Nhập 24-10'): adds.append((D(24),*r,'Thứ Bảy 24/10 chưa nhập (không tăng ca)'))
for r in rows('Nhập 25-10'): adds.append((D(25),*r,'Chủ nhật 25/10 chỉ Assembly làm, không tăng ca'))
for r in rows('Nhập 26-10'): adds.append((D(26),*r,'Thứ Hai 26/10 (ngày thường, có tăng ca)'))
wb=openpyxl.Workbook();ws=wb.active;ws.title='1. Nhập thêm'
ws['A1']='DÒNG CẦN NHẬP THÊM (không gồm S32). Thứ tự cột nhập đúng như form: LossTime Code | Notes | Process | Dept In Charge | From | To | Worker Qty';ws['A1'].font=Font(name=F,bold=True,size=13)
H=['Ngày (Loss Date)','LossTime Code','Notes','Process','Dept In Charge','From','To','Worker Qty','(Tham khảo) Giờ công','(Tham khảo) Lý do']
for i,h in enumerate(H,1):
    c=ws.cell(3,i,h);c.fill=red if i<=8 else gry;c.font=hf if i<=8 else Font(name=F,bold=True);c.alignment=ctr;c.border=bd
ws.row_dimensions[3].height=34
for k,(d,code,notes,pr,dp,fr,to,hc,why) in enumerate(adds):
    r=4+k
    for c,v in enumerate([d,code,notes,pr,dp,fr,to,hc,f'=(G{r}-F{r})*24*H{r}',why],1):
        x=ws.cell(r,c,v);x.font=Font(name=F,color='000000' if c<=8 else '7F7F7F');x.border=bd;x.alignment=Alignment(horizontal='left' if c in(3,10) else 'center',vertical='center',wrap_text=c in(3,10))
    ws.cell(r,1).number_format='dd/mm/yyyy';ws.cell(r,6).number_format=ws.cell(r,7).number_format='hh:mm';ws.cell(r,9).number_format='0.0'
    fill={19:yel,20:yel,21:yel,23:yel}.get(d.day)
    if fill:
        for c in range(1,9): ws.cell(r,c).fill=fill
e=3+len(adds)
ws.cell(e+1,8,'Tổng').font=Font(name=F,bold=True);ws.cell(e+1,9,f'=SUM(I4:I{e})').font=Font(name=F,bold=True);ws.cell(e+1,9).number_format='#,##0.0'
ws.cell(e+3,1,'Vàng = dòng bổ sung cho ngày đã nhập một phần (19/10, 20/10, 21/10, 23/10). Trắng = cả ngày chưa nhập: Thứ Bảy 24/10, Chủ nhật 25/10 (chỉ Assembly), Thứ Hai 26/10 (không tăng ca).').font=Font(name=F,italic=True)
for col,w in zip('ABCDEFGHIJ',[14,12,46,14,16,8,8,10,12,40]): ws.column_dimensions[col].width=w
ws.freeze_panes='A4';ws.auto_filter.ref=f'A3:J{e}'
# 2 sửa
s2=wb.create_sheet('2. Sửa dòng đã nhập')
s2['A1']='DÒNG ĐÃ NHẬP NHƯNG SAI (mở dòng đó trên hệ thống rồi sửa đúng ô ghi bên dưới)';s2['A1'].font=Font(name=F,bold=True,size=13)
H2=['Ngày','Mã','Nhận ra dòng bằng (Notes / mô tả)','Ô cần sửa','Giá trị đang có','Giá trị đúng','Lý do','Giờ công thay đổi']
for i,h in enumerate(H2,1):
    c=s2.cell(3,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
fix=[
(D(20),'P20','Quét nhà toàn xưởng khu vực ASSY (Assembly, 7:30-20:00, 2 người)','Dept In Charge','MACHINING','ASSEMBLY','Sai bộ phận, đồng nghiệp đã chỉ ra',0),
(D(21),'P22','Đào tạo công nhân mới (Assembly, 8:00-16:00)','Worker Qty','1','2','Hệ thống ghi 1 người, danh sách có 2 người Assembly',8),
(D(21),'P20','Vệ sinh khu vực máy móc MAC FINAL (7:00-7:30, 7 người)','Process','Raw Cutting','Machining','Đây là dòng của Machining',0),
(D(21),'P20','Vệ sinh khu vực máy móc PRE MAC (7:00-7:30, 12 người)','Process','Raw Cutting','Machining','Đây là dòng của Machining',0),
(D(23),'P26','Nhận hàng gỗ đầu vào từ RC (Raw Cutting, 7:30-20:00, 1 người)','Dept In Charge','RAW CUTTING','MACHINING','Theo mẫu cố định từ 14/3 (Process Raw Cutting, bộ phận Machining)',0),
(D(5),'P22','Đào tạo công nhân mới (Assembly, 8:00-16:30, 6 người)','To','16:30','16:00','Tùy chọn, ngày cũ: chuẩn là 8:00-16:00',-3),
]
for k,row in enumerate(fix):
    r=4+k
    for c,v in enumerate(row,1):
        x=s2.cell(r,c,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='left' if c in(3,7) else 'center',vertical='center',wrap_text=c in(3,7))
    s2.cell(r,1).number_format='dd/mm/yyyy'
    if row[0].day==5:
        for c in range(1,9): s2.cell(r,c).fill=gry
for col,w in zip('ABCDEFGH',[12,8,52,16,16,16,44,12]): s2.column_dimensions[col].width=w
s2.cell(4+len(fix)+1,1,'Dòng xám (5/10) là tùy chọn, không ảnh hưởng tuần 43.').font=Font(name=F,italic=True)
# 3 xóa
s3=wb.create_sheet('3. Xóa dòng trùng')
s3['A1']='DÒNG TRÙNG CẦN XÓA';s3['A1'].font=Font(name=F,bold=True,size=13)
for i,h in enumerate(['Ngày','Mã','Nhận ra dòng bằng','Số dòng cần xóa','Giờ công bớt đi','Lý do'],1):
    c=s3.cell(3,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
for c,v in enumerate([D(21),'A2','Notes "HC mới- 50% hiệu suất- 2 tuần-09/10", Assembly, 7:30-12:00, 6 người (có 2 dòng giống nhau)',1,27,'Nhập hai lần. Giữ 1 dòng, xóa 1 dòng'],1):
    x=s3.cell(4,c,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='left' if c in(3,6) else 'center',vertical='center',wrap_text=True)
s3.cell(4,1).number_format='dd/mm/yyyy'
for col,w in zip('ABCDEF',[12,8,60,14,14,40]): s3.column_dimensions[col].width=w
# 4 S32 khu vuc khac
s4=wb.create_sheet('4. S32 khu vực khác')
s4['A1']='S32 CỦA RAW CUTTING / MACHINING (không gồm Assembly: cào cước, bắn cát, thợ sơn bạn sẽ nhập khi cần)';s4['A1'].font=Font(name=F,bold=True,size=13)
H4=['Việc cần làm','Ngày','LossTime Code','Notes','Process','Dept In Charge','From','To','Worker Qty','Giờ công','Ghi chú']
for i,h in enumerate(H4,1):
    c=s4.cell(3,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
T=lambda h,m=0:dt.time(h,m)
uv='Không có bomtime chuyền UV'
s4rows=[
('NHẬP THÊM',D(19),'S32',uv,'Raw Cutting','FINISHING',T(7,30),T(15,30),12,'','UV ca thường 19/10 (96 giờ)'),
('NHẬP THÊM',D(19),'S32',uv,'Raw Cutting','FINISHING',T(15,30),T(19,30),10,'','UV tăng ca 19/10 (40 giờ)'),
('NHẬP THÊM',D(20),'S32',uv,'Raw Cutting','FINISHING',T(7,30),T(15,30),12,'','UV ca thường 20/10 (96 giờ)'),
('NHẬP THÊM',D(20),'S32',uv,'Raw Cutting','FINISHING',T(15,30),T(19,30),10,'','UV tăng ca 20/10 (40 giờ)'),
('NHẬP THÊM',D(21),'S32',uv,'Raw Cutting','FINISHING',T(7,30),T(15,30),12,'','UV ca thường 21/10 (96 giờ), ca tăng ca đã có'),
('XÓA',D(24),'S32',uv,'Raw Cutting','FINISHING',T(7,30),T(15,30),12,'','Xóa dòng 12 người (96 giờ). Giữ dòng 11 người (88 giờ), khớp báo cáo UV'),
('SỬA NGÀY',D(18),'S32','Beam saw - phát sinh (chi tiết A05-S01-1, mã AR1113F1330)','Raw Cutting','MACHINING',T(18,0),T(18,2),3,'','Loss Date 18/10 đổi thành 19/10 (nguồn ghi ngày 19/10)'),
('SỬA NGÀY',D(18),'S32','Beam saw - phát sinh (chi tiết A05-S01-2, mã AR1113F1330)','Raw Cutting','MACHINING',T(18,2),T(18,4),3,'','Loss Date 18/10 đổi thành 19/10'),
('SỬA NGÀY',D(18),'S32','Beam saw - phát sinh (chi tiết A1-S02, mã SH20-1029-1471-41191)','Raw Cutting','MACHINING',T(18,4),T(18,7),3,'','Loss Date 18/10 đổi thành 19/10'),
('ĐIỀN BỘ PHẬN',D(20),'S32','Beam saw - phát sinh (CZ567F08), 18:10-18:15, 3 người','Raw Cutting','(trống)','','',3,'','Đang để trống Dept. Đề xuất OTHERS (INVESTIGATE), cùng cách ghi với dòng Unassigned khác'),
('ĐIỀN BỘ PHẬN',D(20),'S32','Beam saw - phát sinh (CZ567F09), 18:15-18:18, 3 người','Raw Cutting','(trống)','','',3,'','Đang để trống Dept. Đề xuất OTHERS (INVESTIGATE)'),
('ĐIỀN BỘ PHẬN',D(20),'S32','Veneer walnut cắt ghép (bs lại), RH21-2184-492, 10:00-11:15, 2 người','Raw Cutting','(trống)','','',2,'','Đang để trống Dept. Nguồn ghi "bs lại", cần hỏi bộ phận đúng'),
]
for k,row in enumerate(s4rows):
    r=4+k
    for c,v in enumerate(row[:9],1):
        x=s4.cell(r,c,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='left' if c in(4,) else 'center',vertical='center',wrap_text=c in(4,))
    s4.cell(r,10,f'=IF(OR(G{r}="",H{r}=""),"",(H{r}-G{r})*24*I{r})').border=bd
    s4.cell(r,11,row[10]).border=bd
    for c in (10,11): s4.cell(r,c).font=Font(name=F);s4.cell(r,c).alignment=Alignment(horizontal='left' if c==11 else 'center',vertical='center',wrap_text=True)
    s4.cell(r,2).number_format='dd/mm/yyyy';s4.cell(r,7).number_format=s4.cell(r,8).number_format='hh:mm';s4.cell(r,10).number_format='0.0'
    f={'NHẬP THÊM':grn,'XÓA':org,'SỬA NGÀY':yel,'ĐIỀN BỘ PHẬN':yel}[row[0]]
    s4.cell(r,1).fill=f
for col,w in zip('ABCDEFGHIJK',[16,12,10,50,12,18,8,8,10,10,60]): s4.column_dimensions[col].width=w
# 5 tong ket
s5=wb.create_sheet('5. Tổng kết')
s5['A1']='GIỜ CÔNG P20+P26+P22+A2 THEO NGÀY: HỆ THỐNG HIỆN TẠI VÀ SAU KHI SỬA';s5['A1'].font=Font(name=F,bold=True,size=13)
for i,h in enumerate(['Ngày','Hệ thống hiện tại','Thêm (sheet 1)','Sửa (sheet 2)','Xóa (sheet 3)','Sau khi sửa','Bảng chuẩn của mình','Khớp'],1):
    c=s5.cell(3,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
cur={19:524,20:545,21:509.5,22:509,23:470,24:0,25:0,26:0}
lt=openpyxl.load_workbook('Loss_time_05-10_den_26-10.xlsx',data_only=True)
def tot(d):
    ws=lt[f'Loss time {d:02d}-10'];return round(sum(r[16] or 0 for r in ws.iter_rows(min_row=3,values_only=True) if r[7] in('P20','P26','P22','A2')),2)
std={d:tot(d) for d in cur}
sfix={21:8}; sdel={21:27}
n=len(adds)
for k,d in enumerate(sorted(cur)):
    r=4+k
    s5.cell(r,1,D(d)).number_format='dd/mm/yyyy';s5.cell(r,2,cur[d])
    s5.cell(r,3,f"=SUMIFS('1. Nhập thêm'!$I$4:$I${3+n},'1. Nhập thêm'!$A$4:$A${3+n},A{r})")
    s5.cell(r,4,sfix.get(d,0));s5.cell(r,5,-sdel.get(d,0))
    s5.cell(r,6,f'=B{r}+C{r}+D{r}+E{r}');s5.cell(r,7,std[d]);s5.cell(r,8,f'=IF(ABS(F{r}-G{r})<0.01,"Khớp","Lệch")')
    for c in range(1,9):
        x=s5.cell(r,c);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='center')
        if c in(2,3,4,5,6,7): x.number_format='#,##0.0;-#,##0.0;-'
s5.cell(14,1,'Cột "Hệ thống hiện tại" lấy từ file xuất 26/10 (chỉ dòng do bạn tạo, mã P20, P26, P22, A2). "Bảng chuẩn" là bảng loss time mình lập. 25/10 (Chủ nhật) chưa có.').font=Font(name=F,italic=True)
for col,w in zip('ABCDEFGH',[14,18,16,14,14,14,20,10]): s5.column_dimensions[col].width=w
wb.move_sheet('5. Tổng kết',offset=-4)
wb.save('Danh_sach_sua_them_19-10_den_26-10.xlsx');print('adds',len(adds))
