import pickle,datetime as dt,openpyxl,collections
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.utils import get_column_letter as L
R=pickle.load(open('s32all.pkl','rb'))
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
red=PatternFill('solid',fgColor='C00000');hf=Font(name=F,bold=True,color='FFFFFF');ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
yel=PatternFill('solid',fgColor='FFF2CC');org=PatternFill('solid',fgColor='FCE4D6')
SRC={'Ép nóng':'File Ép nóng','UV hằng ngày':'File UV (báo cáo ngày)','Beam saw':'File Beam saw','Veneer':'File Veneer','UV phát sinh':'File UV (phát sinh)','Thợ sơn (hình)':'Hình: thợ sơn','Cào cước (hình)':'Hình: cào cước','Bắn cát (hình)':'Hình: bắn cát'}
def pn(x):
    its=x['items']
    if its:
        codes=[str(i[0]) for i in its]
        return '+'.join(codes[:3])+(f' (+{len(codes)-3} mã)' if len(codes)>3 else '')
    return x['item']
def qty(x):
    its=x['items']
    if its:
        q=[i[2] for i in its if i[2] is not None]
        return sum(q) if q else None
    return x['qty']
wb=openpyxl.Workbook();ws=wb.active;ws.title='S32 5-10 đến 20-10'
ws['A1']='LOSS TIME MÃ S32 (OTHER SUPPORT OF PRODUCTION) - 05/10 ĐẾN 20/10/2026 - LÀM RIÊNG, KHÔNG LẪN MÃ KHÁC';ws['A1'].font=Font(name=F,bold=True,size=15)
H=['Week','Date \n(Ngày)',"Item's\nMã hàng",'Code Máy Hư',"Q'ty",'Số NCR/ME','Khu vực phát sinh/bộ phận hỗ trợ','Code','Name code','Chart / bộ phận chịu trách nhiệm','Lí Do','Ghi chú','From','To','LossTime (Hr)','HC','Loss time labor (Hrs)','Nguồn']
for i,h in enumerate(H,1):
    c=ws.cell(2,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
ws.row_dimensions[2].height=37
for k,x in enumerate(R):
    r=3+k
    vals=[f'="W"&WEEKNUM(B{r},21)',x['date'],pn(x),None,qty(x),None,x['area'],'S32','Other Support of Production',x['dept'],x['lido'],x['gc'],x['fr'],x['to'],f'=(N{r}-M{r})*24',x['hc'],f'=P{r}*O{r}',SRC[x['src']]]
    for c,v in enumerate(vals,1):
        a=ws.cell(r,c,v);a.font=Font(name=F);a.border=bd;a.alignment=Alignment(horizontal='center',vertical='center',wrap_text=c in(3,9,11,12))
        if x['flag'] in('P','A'): a.fill=yel
        elif x['flag'] in('U','E'): a.fill=org
    ws.cell(r,2).number_format='d-mmm';ws.cell(r,13).number_format=ws.cell(r,14).number_format='h:mm;@'
    ws.cell(r,15).number_format='0.00';ws.cell(r,17).number_format='0.00'
e=2+len(R)
ws.cell(e+1,16,'Tổng').font=Font(name=F,bold=True);ws.cell(e+1,17,f'=SUM(Q3:Q{e})').font=Font(name=F,bold=True);ws.cell(e+1,17).number_format='#,##0.00'
ws.cell(e+3,1,'Vàng = giờ cần xác nhận (dự kiến hoặc giả định). Cam = dòng cần quyết định (Chủ nhật 11/10 UV; Beam saw 16/10 chưa phân công bộ phận).').font=Font(name=F,italic=True)
for col,w in zip('ABCDEFGHIJKLMNOPQR',[7,9,28,10,8,10,16,8,26,16,42,36,9,9,10,8,15,22]): ws.column_dimensions[col].width=w
ws.freeze_panes='A3';ws.auto_filter.ref=f'A2:R{e}'
M=f"'S32 5-10 đến 20-10'"
# ---- theo ngày
dd=sorted({x['date'] for x in R});srcs=['Ép nóng','UV hằng ngày','Beam saw','Veneer','UV phát sinh','Thợ sơn (hình)','Cào cước (hình)','Bắn cát (hình)']
sd=wb.create_sheet('Theo ngày')
sd['A1']='GIỜ CÔNG S32 THEO NGÀY VÀ THEO NGUỒN';sd['A1'].font=Font(name=F,bold=True,size=14)
hh=['Ngày','Thứ','Tuần']+[SRC[s] for s in srcs]+['Tổng ngày']
for i,h in enumerate(hh,1):
    c=sd.cell(3,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
sd.row_dimensions[3].height=40
wdn=['Thứ Hai','Thứ Ba','Thứ Tư','Thứ Năm','Thứ Sáu','Thứ Bảy','Chủ nhật']
for k,d in enumerate(dd):
    r=4+k
    sd.cell(r,1,d).number_format='dd/mm/yyyy';sd.cell(r,2,wdn[d.weekday()]);sd.cell(r,3,f'="W"&WEEKNUM(A{r},21)')
    for j,s in enumerate(srcs):
        col=4+j
        sd.cell(r,col,f'=SUMIFS({M}!$Q$3:$Q${e},{M}!$B$3:$B${e},$A{r},{M}!$R$3:$R${e},{L(col)}$3)')
    sd.cell(r,4+len(srcs),f'=SUM(D{r}:{L(3+len(srcs))}{r})')
rl=4+len(dd)
sd.cell(rl,1,'Tổng')
for j in range(len(srcs)+1):
    col=4+j;sd.cell(rl,col,f'=SUM({L(col)}4:{L(col)}{rl-1})')
for rr in range(4,rl+1):
    for c in range(1,5+len(srcs)):
        x=sd.cell(rr,c);x.font=Font(name=F,bold=(rr==rl or c==4+len(srcs)));x.border=bd;x.alignment=Alignment(horizontal='center',vertical='center')
        if c>=4: x.number_format='#,##0.0;-#,##0.0;-'
# weekly
wr=rl+3;sd.cell(wr-1,1,'TỔNG THEO TUẦN').font=Font(name=F,bold=True)
for i,h in enumerate(['Tuần','Giờ công'],1):
    c=sd.cell(wr,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
for k,w in enumerate(['W41','W42','W43']):
    r=wr+1+k;sd.cell(r,1,w);sd.cell(r,2,f'=SUMIFS({M}!$Q$3:$Q${e},{M}!$A$3:$A${e},A{r})')
    for c in (1,2):
        x=sd.cell(r,c);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='center')
    sd.cell(r,2).number_format='#,##0.0'
sd.column_dimensions['A'].width=13;sd.column_dimensions['B'].width=11;sd.column_dimensions['C'].width=8
for c in range(4,5+len(srcs)): sd.column_dimensions[L(c)].width=14
# ---- nhập liệu
ip=wb.create_sheet('Nhập liệu S32')
hi=['Ngày (Loss Date)','LossTime Code','Notes','Process','Dept In Charge','From','To','Worker Qty','(Tham khảo) Part Number','(Tham khảo) Total Order','(Tham khảo) Ghi chú','(Tham khảo) Nguồn']
for i,h in enumerate(hi,1):
    c=ip.cell(1,i,h);c.fill=red if 2<=i<=8 else (PatternFill('solid',fgColor='7F7F7F') if i>8 else red);c.font=hf;c.alignment=ctr;c.border=bd
ip.row_dimensions[1].height=34
for k,x in enumerate(R):
    r=2+k
    for c,v in enumerate([x['date'],'S32',x['lido'],x['area'],x['dept'],x['fr'],x['to'],x['hc'],pn(x),qty(x),x['gc'],SRC[x['src']]],1):
        a=ip.cell(r,c,v);a.font=Font(name=F,color='000000' if c<=8 else '7F7F7F');a.border=bd;a.alignment=Alignment(horizontal='left' if c in(3,11) else 'center',vertical='center',wrap_text=c in(3,9,11))
        if x['flag'] in('P','A') and c<=8: a.fill=yel
        elif x['flag'] in('U','E') and c<=8: a.fill=org
    ip.cell(r,1).number_format='dd/mm/yyyy';ip.cell(r,6).number_format=ip.cell(r,7).number_format='hh:mm'
for col,w in zip('ABCDEFGHIJKL',[14,10,44,12,14,8,8,10,30,12,36,22]): ip.column_dimensions[col].width=w
ip.freeze_panes='A2'
ip.cell(len(R)+3,1,'Đổi Loss Date theo cột A khi nhập. Cột 2 đến 8 là các ô cần nhập (From/To bạn tự nhập). Các cột xám chỉ để tham khảo.').font=Font(name=F,italic=True)
# ---- đối chiếu nguồn
import openpyxl as ox
Z='s32zip/Lost time S32/'
def sumcol(f,sheet,pred,col):
    ws2=ox.load_workbook(Z+f,read_only=True,data_only=True)[sheet];s=0
    for r in ws2.iter_rows(values_only=True):
        if pred(r) and r[col] is not None:
            try: s+=float(r[col])
            except: pass
    return s
D0=dt.datetime(2026,10,5)
ref={
 'File Beam saw':sumcol('LOST TIME BEMSAW 20.10.xlsx','data ',lambda r:isinstance(r[6],dt.datetime) and r[6]>=D0 and r[10]>=0,10),
 'File UV (báo cáo ngày)':sumcol('lost time khu vực UV (1).xlsx','LOST TIME',lambda r:isinstance(r[1],dt.datetime) and r[1]>=D0,5),
 'File UV (phát sinh)':sumcol('lost time khu vực UV (1).xlsx','Sheet1',lambda r:isinstance(r[6],dt.datetime) and r[6]>=D0,10),
 'File Ép nóng':sumcol('FILE ÉP HÀNG SỬA ( KV ÉP NÓNG ) (1).xlsx','datta',lambda r:isinstance(r[6],dt.datetime) and r[6]>=D0,10),
 'File Veneer':sumcol('26.10.20 Lost time veneer Hàng BS.xlsx','tổng',lambda r:r[11] is not None and isinstance(r[11],dt.datetime) and r[11]>=D0,10),
 'Hình: thợ sơn':816,'Hình: cào cước':9,'Hình: bắn cát':32}
dr=wb.create_sheet('Đối chiếu nguồn')
dr['A1']='ĐỐI CHIẾU GIỜ CÔNG VỚI SỐ GIỜ TRONG FILE/HÌNH NGUỒN';dr['A1'].font=Font(name=F,bold=True,size=14)
for i,h in enumerate(['Nguồn','Số giờ trong file/hình nguồn','Giờ công trong bảng S32','Chênh lệch','Kết quả'],1):
    c=dr.cell(3,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
dr.row_dimensions[3].height=36
for k,(s,v) in enumerate(ref.items()):
    r=4+k;dr.cell(r,1,s);a=dr.cell(r,2,round(v,4));a.font=Font(name=F,color='0000FF')
    dr.cell(r,3,f'=SUMIFS({M}!$Q$3:$Q${e},{M}!$R$3:$R${e},A{r})');dr.cell(r,4,f'=C{r}-B{r}');dr.cell(r,5,f'=IF(ABS(D{r})<0.01,"Khớp","Lệch")')
r=4+len(ref);dr.cell(r,1,'Tổng')
for c,col in ((2,'B'),(3,'C'),(4,'D')): dr.cell(r,c,f'=SUM({col}4:{col}{r-1})')
dr.cell(r,5,f'=IF(ABS(D{r})<0.01,"Khớp","Lệch")')
for rr in range(4,r+1):
    for c in range(1,6):
        x=dr.cell(rr,c);x.border=bd;x.alignment=Alignment(horizontal='left' if c==1 else 'center',vertical='center')
        if c!=2 or rr==r: x.font=Font(name=F,bold=(rr==r))
        if c in(2,3,4): x.number_format='#,##0.00'
dr.cell(r+2,1,'Chữ xanh = số lấy từ file/hình nguồn (cột "SỐ GIỜ" hoặc tổng tuần trong báo cáo). Chỉ lấy các dòng từ 05/10/2026 trở đi.').font=Font(name=F,italic=True)
for col,w in zip('ABCDE',[28,26,24,14,12]): dr.column_dimensions[col].width=w
# ---- chi tiết mã hàng
cd=wb.create_sheet('Chi tiết mã hàng')
for i,h in enumerate(['Ngày','Nguồn','Mã hàng / code','Tên chi tiết','Số lượng'],1):
    c=cd.cell(1,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
r=2
for x in R:
    for it in x['items']:
        for c,v in enumerate([x['date'],SRC[x['src']],it[0],it[1],it[2]],1):
            a=cd.cell(r,c,v);a.font=Font(name=F);a.border=bd;a.alignment=Alignment(horizontal='center')
        cd.cell(r,1).number_format='dd/mm/yyyy';r+=1
for col,w in zip('ABCDE',[12,24,26,16,10]): cd.column_dimensions[col].width=w
# ---- ghi chú review
rv=wb.create_sheet('Ghi chú review')
Lns=[('GHI CHÚ REVIEW - BẢNG S32 (LÀM RIÊNG)',True),
('1. Bảng này CHỈ chứa mã S32, tách hoàn toàn khỏi bảng P20/P26/P22/A2 (không lẫn mã khác). Khi bạn cập nhật, tôi giữ nguyên nguyên tắc này.',False),
('2. Nguồn: 5 file trong gói zip (lấy bản mới nhất mỗi loại: Beam saw bản 20/10, UV bản (1), Ép nóng bản (1), Veneer bản 20/10) và 3 hình (thợ sơn tuần 42, cào cước 16/10, bắn cát 15/10). Chỉ lấy dòng từ 05/10 đến 20/10. UV (báo cáo ngày đến 11/10, phát sinh đến 13/10) và Ép nóng (đến 6/10) chưa có số liệu mới.',False),
('3. Tất cả giờ công trong bảng khớp với số giờ trong nguồn (sheet "Đối chiếu nguồn"). Tổng giờ công xem sheet "Đối chiếu nguồn". Riêng dòng Beam saw 20/10 (18:15-18:13) nguồn ghi -0,1 giờ nên bảng ghi 0 giờ và không tính vào đối chiếu.',False),
('4. UV báo cáo ngày: báo cáo ghi giờ = 8 giờ x HC làm việc + 4 giờ x HC tăng ca. Tôi tách mỗi ngày thành 2 dòng (ca thường 7:30-15:30; tăng ca 15:30-19:30) để ra đúng số giờ báo cáo. Tiền lệ trong file ghi 1 dòng/ngày (7:30-19:30), cần xác nhận cách ghi.',False),
('5. UV Chủ nhật 11/10: có 9 người làm 8 giờ (72 giờ). Cần xác nhận có nhập Chủ nhật không (tô cam).',False),
('6. UV: file báo cáo ngày chỉ đến 11/10; chưa có số liệu từ 12/10. UV phát sinh 12/10 có 23 mã hàng SLAT nhưng chỉ 2 dòng có giờ (HC 1, 9:30-20:00 và 7:30-15:00); danh sách mã ở sheet "Chi tiết mã hàng". Bộ phận ghi "QC" (báo cáo ghi "QC A HIẾU"), cần xác nhận.',False),
('7. Beam saw: 14 dòng từ 08/10 đến 20/10. Dòng CZ567F09 ngày 20/10 ghi 18:15-18:13 (giờ kết thúc nhỏ hơn giờ bắt đầu), cần xác nhận giờ đúng (tô cam). 3 dòng ngày 19/10 (18:00-18:07) trên hệ thống đang nhập ngày 18/10, nguồn ghi ngày 19/10, cần kiểm tra lại. Khu vực ghi RC - PANEL (Raw Cutting, người dùng xác nhận 20/10). Bộ phận lấy từ cột Charge to: Machining CNC và Machining ghi "Machining"; "QA ..." ghi "QC"; 4 dòng ngày 16/10 và 20/10 là "Unassigned" (chưa phân công), cần xác định bộ phận (tô cam).',False),
('8. Ép nóng: 3 nhóm (5/10 hỗ trợ P1 4 người 16 giờ; 5/10 ép lót 5 người 30 giờ; 6/10 ép chính 5 người 55 giờ). Mỗi nhóm có nhiều mã hàng cùng một ca nên ghi một dòng, mã hàng nối bằng "+" (xem sheet Chi tiết). Đừng nhân giờ cho từng mã. Giờ bắt đầu 6:30 hiếm, nên xác nhận.',False),
('9. Veneer (làm mẫu): 10/10 (2 người 9:00-10:30, 3 giờ, TF), 15/10 (1 người 8:00-8:20, ME) và 20/10 (code 191: 1 người 9:00-9:50, TF; walnut cắt ghép RH21-2184-492: 2 người 10:00-11:15, nguồn ghi Charge to "bs lại" nên bộ phận cần xác nhận, tô cam; code 014: 1 người 10:00-10:40, ME). Khu vực ghi RC - PANEL (Raw Cutting, người dùng xác nhận 20/10).',False),
('10. Thợ sơn tuần 42: 12 người, 7:30-19:30 từ 12/10 đến 16/10 (720 giờ), thứ Bảy 17/10 7:30-15:30 (96 giờ, có thể là giờ dự kiến). Cào cước 16/10: dòng 15:00-20:00 có thể là giờ dự kiến. Bắn cát 15/10: ghi 7:00-15:00 để đúng 8 giờ/người (báo cáo ghi 7:00-16:00).',False),
('11. Tên bộ phận và khu vực đã đổi theo dropdown hệ thống (lấy từ file xuất 20/10): FINISHING, TECHNICAL FINISHING, OTHERS (INVESTIGATE) cho hỗ trợ P1, QA/QC/TESTING, SUBCON, MACHINING, ME; khu vực Raw Cutting và Assembly. Cách ghi này theo người nhập trước (Hùng, Tiến).',False)]
rv.column_dimensions['A'].width=150
for i,(tx,b) in enumerate(Lns,1):
    c=rv.cell(i,1,tx);c.font=Font(name=F,bold=b,size=12 if b else 11);c.alignment=Alignment(wrap_text=True,vertical='top')
wb.save('Loss_time_S32_05-10_den_20-10.xlsx');print('ok',len(R))
