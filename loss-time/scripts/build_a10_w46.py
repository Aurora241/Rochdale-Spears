"""Bảng nhập liệu A10 từ báo cáo bảo trì Stop line + Repair WW 9/11-14/11/2026 (tuần W46).
Chạy trong thư mục outputs. Worker Qty chỉ điền khi lịch sử A10 và file HC Mapping cùng cho một số người."""
import openpyxl,json,re,datetime as dt,sys
from openpyxl.styles import Font,PatternFill,Alignment
sys.path.insert(0,'../scripts')
import build_nhap_lieu_theo_bo_phan as B
T=lambda h,m=0:dt.time(h,m)
# (ngày, khu vực báo cáo, mã máy, tên máy, hiện tượng, từ, đến, giờ báo cáo, ghi chú nguồn)
# (STT báo cáo, ngày, khu vực, mã máy, tên máy, vấn đề, từ, đến, ghi chú chia ngày) - mỗi dòng là một dòng nhập (máy hỏng qua nhiều ngày đã chia theo ngày, chỉ tính trong khung chạy máy)
D=lambda d:dt.datetime(2026,11,d)
rep=[
(180,9,'Machining','CNNR01','Máy CNC 5 trục','Máy bị lỗi chương trình',T(9,18),T(9,52),''),
(181,9,'Machining','DECO5','Máy cắt hai đầu băng tải','Máy không ra vo băng tải được',T(11,6),T(13,41),''),
(182,9,'Machining','ADET02','Máy đánh mộng dương 2 đầu','Máy bị hư',T(8,50),T(20,0),'Hỏng 9/11 8:50, sửa xong 11/11 16:29: tính khung 7:30-20:00 (11,17 + 12,5 + 8,98 = 32,65 giờ)'),
(182,10,'Machining','ADET02','Máy đánh mộng dương 2 đầu','Máy bị hư',T(7,30),T(20,0),'Hỏng 9/11 8:50, sửa xong 11/11 16:29: tính khung 7:30-20:00'),
(182,11,'Machining','ADET02','Máy đánh mộng dương 2 đầu','Máy bị hư',T(7,30),T(16,29),'Hỏng 9/11 8:50, sửa xong 11/11 16:29: tính khung 7:30-20:00'),
(183,11,'Machining','WBS09','Máy nhám thùng','Máy bị lỗi chương trình',T(7,41),T(9,13),''),
(184,11,'Machining','CNCR14','Máy CNC 5 trục','Ống hút bụi bị hư',T(18,0),T(18,45),''),
(185,12,'Machining','MOT01','Máy đục đơn','Máy bị hư',T(18,41),T(19,23),''),
(186,12,'Machining','CORCO1','Máy phay ngàm','Máy bị hư',T(19,37),T(20,0),'Hỏng 12/11 19:37, sửa xong 13/11 10:25: tính khung 7:30-20:00 (0,38 + 2,92 = 3,3 giờ)'),
(186,13,'Machining','CORCO1','Máy phay ngàm','Máy bị hư',T(7,30),T(10,25),'Hỏng 12/11 19:37, sửa xong 13/11 10:25: tính khung 7:30-20:00'),
(187,13,'RC','STS01','Máy chà nhám băng/ Hand Stroke Sanding','Máy bị hư bạc đạn',T(13,16),T(16,34),''),
(188,13,'Machining','CNCC01','Máy CNC nhiều mũi','Máy bị lỗi trục X',T(13,50),T(16,34),''),
(189,13,'Machining','CNCR01','Máy Hogmang 1','Máy bị hư bạc đạn trục X',T(22,52),T(23,59),'Hỏng 13/11 22:52, sửa xong 14/11 0:45 (Time run 24 giờ/ngày): chia theo ngày, ngoài ca ngày, cần xác nhận'),
(189,14,'Machining','CNCR01','Máy Hogmang 1','Máy bị hư bạc đạn trục X',T(0,0),T(0,45),'Hỏng 13/11 22:52, sửa xong 14/11 0:45 (Time run 24 giờ/ngày): chia theo ngày, ngoài ca ngày, cần xác nhận'),
(190,14,'Machining','CNCR03','Máy CNC shoda','Máy bị lỗi bàn đạp hơi',T(9,36),T(9,46),''),
]
AR={'RC':'Raw Cutting','Machining':'Machining'}
wbh=openpyxl.load_workbook('../outputs/A10_so_nguoi_theo_may.xlsx',data_only=True)
hist={r[0]:r for r in wbh['3. Số người theo mã máy'].iter_rows(min_row=4,values_only=True) if r[0]}
mp=json.load(open('../reference/hc_mapping_by_machine.json'))
w45={}
for r in list(openpyxl.load_workbook('/root/.claude/uploads/64423226-a456-5e96-8225-6a5ffe123efb/9ca35d80-tmp265B.tmp.xlsx',data_only=True).active.iter_rows(values_only=True))[14:]:
    if r[2]=='A10' and r[14]:
        m=re.match(r'^([A-Z]{2,}[0-9]{1,3})-',r[14])
        if m: w45.setdefault(m.group(1),[]).append(r[17])
def norm(c): return re.sub(r'(?<=[A-Z])O(?=\d)','0',c)
FAM={'CNCR':'Họ CNCR lịch sử: 2 người (68%), 3 người (32%)','VEC':'Họ VEC (cắt veneer): lịch sử VEC04 2 người (4 lần)'}
rows=[];NOTE={};ref=[]
for stt,d,ar,mc,mn,fl,fr,to,cn in rep:
    h=((dt.datetime.combine(dt.date.today(),to)-dt.datetime.combine(dt.date.today(),fr)).seconds/3600)
    notes='-'.join(x for x in (mc,mn,fl) if x)
    nc=norm(mc);typo=(nc!=mc)
    hs=hist.get(nc);mapv=mp.get(nc)
    hm=hs[4] if hs else None;mv=mapv['actual_ops'] if mapv else None
    fill=hm if (hm is not None and mv is not None and int(hm)==int(mv) and not typo and 'cần xác nhận' not in cn) else None
    if fill: src=f'Lịch sử A10 ({hs[3]} lần: {hs[6]}) và file HC Mapping cùng {int(fill)} người (chưa có tổ trưởng xác nhận)'
    elif mc and (hm or mv): src=f'Lịch sử: {hs[6] if hs else "chưa từng có"}; Mapping thực tế: {mv if mv is not None else "không có"} - lệch/chỉ một nguồn nên để trống'
    elif mc: src='Chưa có nguồn'
    else: src='Báo cáo không ghi mã máy. '+('Họ CNCR lịch sử: 2 người (68%), 3 người (32%)' if 'CNC' in mn else 'Chưa có nguồn')
    if typo: src='Mã máy ghi "'+mc+'" nghi là '+nc+' nên không tự điền; '+src
    if mc in w45 or nc in w45: src+=f'. Tuần W45 bạn nhập: {sorted(set(w45.get(mc,w45.get(nc))))} người'
    if cn: src+=' | '+cn
    rows.append((D(d),'A10',notes,AR[ar],'MAINTENANCE',fr,to,int(fill) if fill else None));NOTE[(D(d),notes,fr)]=src
    ref.append((D(d),ar,mc,mn,fl,fr,to,h,hm,mv))
B.ORDER.insert(0,'A10')
wb=openpyxl.Workbook();wb.remove(wb.active)
F=B.F;bd=B.bd;nfill=sum(1 for r in rows if r[7])
for a in ('Machining','Raw Cutting'):
    ws=B.sheet(wb,a,[x for x in rows if x[3]==a])
    ws['A2']=f'WORKER QTY: {nfill} dòng đã điền vì lịch sử A10 và file HC Mapping cùng cho một số người (chưa có tổ trưởng xác nhận, ghi tên người xác nhận ở cột L). Các dòng để trống (ô vàng) chưa đủ nguồn: cần tổ trưởng cho số công nhân thật đứng máy. Cột "Time run" 12 trong báo cáo là giờ máy chạy mỗi ngày, KHÔNG phải số người.'
    ws['A2'].font=Font(name=F,italic=True,color='C00000')
    for cc,tx in ((11,'Nguồn / số tham khảo'),(12,'Người xác nhận số công nhân')):
        ws.cell(4,cc,tx).fill=B.gry;ws.cell(4,cc).font=Font(name=F,bold=True);ws.cell(4,cc).alignment=B.ctr;ws.cell(4,cc).border=bd
    ws.column_dimensions['K'].width=58;ws.column_dimensions['L'].width=26
    for r_ in range(5,ws.max_row+1):
        if isinstance(ws.cell(r_,1).value,dt.datetime):
            hv=ws.cell(r_,8).value
            ws.cell(r_,8).fill=PatternFill('solid',fgColor='E2EFDA' if hv else 'FFFF00')
            ws.cell(r_,11,NOTE.get((ws.cell(r_,1).value,ws.cell(r_,3).value,ws.cell(r_,6).value),'')).font=Font(name=F,color='7F7F7F');ws.cell(r_,11).border=bd;ws.cell(r_,11).alignment=Alignment(wrap_text=True,vertical='center');ws.cell(r_,12).border=bd
ws=wb.create_sheet('Đối chiếu bảo trì')
ws['A1']='BÁO CÁO BẢO TRÌ 9/11-14/11 (STOP LINE + REPAIR WW) VÀ NGUỒN SỐ NGƯỜI';ws['A1'].font=Font(name=F,bold=True,size=13)
HH=['Ngày','Khu vực (báo cáo)','Mã máy','Tên máy','Hiện tượng','Từ','Đến','Giờ dừng báo cáo','Giờ dừng tính lại','Chênh lệch','Lịch sử A10: số người hay gặp','Mapping: thực tế vận hành']
for i,h in enumerate(HH,1):
    c=ws.cell(3,i,h);c.fill=B.red;c.font=B.hf;c.alignment=B.ctr;c.border=bd
ws.row_dimensions[3].height=40
for k,(d,ar,mc,mn,fl,fr,to,h,hm,mv) in enumerate(ref):
    r=4+k
    for c,v in enumerate([d,ar,mc,mn,fl,fr,to,h,f'=(G{r}-F{r})*24',f'=ROUND(I{r}-H{r},2)',hm,mv],1):
        x=ws.cell(r,c,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='left' if c in(4,5) else 'center',vertical='center',wrap_text=c in(4,5))
    ws.cell(r,1).number_format='dd/mm/yyyy';ws.cell(r,6).number_format=ws.cell(r,7).number_format='hh:mm';ws.cell(r,9).number_format='0.00'
r=4+len(ref)
ws.cell(r+1,8,'Tổng').font=Font(name=F,bold=True);ws.cell(r+1,9,f'=SUM(I4:I{r-1})').font=Font(name=F,bold=True);ws.cell(r+1,9).number_format='0.00'
ws.cell(r+3,1,'Ghi chú: STT 182 (ADET02) và STT 186 (CORCO1) hỏng qua nhiều ngày, đã chia mỗi ngày một dòng và chỉ tính trong khung chạy máy 7:30-20:00 (tổng 32,65 giờ và 3,3 giờ khớp báo cáo). STT 189 (CNCR01, Time run 24) hỏng 13/11 22:52 đến 14/11 0:45 ngoài ca ngày: chia 2 dòng (22:52-23:59 và 0:00-0:45, tổng 1,88 giờ), cần xác nhận cách nhập. Mã máy CNNR01, DECO5, CORCO1 có thể ghi nhầm. STS01 ngày 13/11 ghi khu vực RC (ngày 4/11 báo cáo ghi Machining).').font=Font(name=F,italic=True)
for col,w in zip('ABCDEFGHIJKL',[12,12,10,30,34,8,8,12,12,10,18,18]): ws.column_dimensions[col].width=w
wb.save('Nhap_lieu_A10_theo_bo_phan_09-11_den_14-11.xlsx');print('dòng',len(rows),'đã điền',nfill,'tổng giờ máy',round(sum(x[7] for x in ref),2))
for x in rows:print(x[0].strftime('%d/%m'),x[2][:34],x[3],x[7])
