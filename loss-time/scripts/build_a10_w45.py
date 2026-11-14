"""Bảng nhập liệu A10 từ báo cáo bảo trì Stop line + Repair WW 2/11-8/11/2026 (tuần W45).
Chạy trong thư mục outputs. Worker Qty chỉ điền khi lịch sử A10 và file HC Mapping cùng cho một số người."""
import openpyxl,json,datetime as dt,sys
from openpyxl.styles import Font,PatternFill,Alignment
sys.path.insert(0,'../scripts')
import build_nhap_lieu_theo_bo_phan as B
T=lambda h,m=0:dt.time(h,m);D=lambda d:dt.datetime(2026,11,d)
# (ngày, khu vực báo cáo, mã máy, tên máy, hiện tượng, từ, đến, giờ báo cáo, ghi chú nguồn)
rep=[
(2,'Machining','EDB05','Máy dán cạnh thẳng tự động','Máy bị lỗi cỡ lên xuống',T(8,20),T(9,0),0.67),
(2,'Machining','','Máy CNC 5 trục','Máy bị lỗi chương trình',T(12,48),T(13,3),0.25),
(2,'RC','VET02','Máy thẩm cạnh Veneer','Dừng máy vệ sinh hút bụi',T(14,14),T(14,38),0.40),
(2,'RC','VEC03','Máy cắt veneer','Dừng máy vệ sinh hút bụi',T(14,14),T(14,38),0.40),
(2,'RC','','Máy chuyền UV','Đèn thủy ngân bị lỗi',T(17,45),T(18,25),0.67),
(3,'RC','VEC04','Máy cắt veneer','Máy bị lỗi tự động cắt',T(18,10),T(18,50),0.67),
(4,'RC','STS01','Máy chà nhám băng/ Hand Stroke Sanding','Con lăn của máy không di chuyển được',T(13,20),T(13,38),0.30),
(4,'RC','VES09','Máy may veneer','Máy bị hư dây cu-roa',T(16,30),T(18,1),1.52),
(4,'Machining','','Máy CNC Nesting (Panel)','Máy bị lỗi nút dừng khẩn cấp',T(19,7),T(19,37),0.50),
(5,'Machining','DEC06','Máy cắt 2 đầu băng tải / Double End Cutting','Máy bị lỗi không lên nguồn',T(8,15),T(13,55),5.67),
(5,'Machining','CUFS19','Máy cắt','Máy bị hư dây cu-roa',T(13,6),T(13,50),0.73),
(5,'RC','VES08','Máy may veneer','Máy không hoạt động',T(18,9),T(19,16),1.12),
(6,'Machining','CNCR11','Máy CNC','Máy bị bể ống hơi',T(6,40),T(7,18),0.63),
(6,'RC','VES02','Máy may veneer','Máy bị lỗi con lăn',T(7,38),T(8,10),0.53),
(6,'RC','VES04','Máy may veneer','Máy bị lỗi băng truyền động',T(9,0),T(10,23),1.38),
(6,'Machining','','Máy CNC 5 trục','Máy bị lệch trục C',T(10,28),T(11,9),0.68),
(6,'RC','BEMS04','Máy cắt MDF','Máy bị đứt dây Cu-roa',T(10,25),T(13,55),3.50),
(6,'Machining','CNCR11','Máy CNC','Máy bị lỗi',T(17,8),T(17,59),0.85),
(6,'RC','UVDR01','Chuyền UV','Đèn UV bị lỗi',T(17,47),T(18,4),0.28),
(7,'Machining','CNCR01','Máy Hogmang 1','Máy bị bung ốc',T(9,35),T(10,50),1.25),
(8,'RC','DEC06','Máy cắt 2 đầu băng tải / Double End Cutting','Máy bị lỗi không lên nguồn',T(15,45),T(16,15),0.50),
]
AR={'RC':'Raw Cutting','Machining':'Machining'}
wbh=openpyxl.load_workbook('../outputs/A10_so_nguoi_theo_may.xlsx',data_only=True)
hist={r[0]:r for r in wbh['3. Số người theo mã máy'].iter_rows(min_row=4,values_only=True) if r[0]}
mp=json.load(open('../reference/hc_mapping_by_machine.json'))
FAM={'CNCR':'Họ CNCR lịch sử: 2 người (68%), 3 người (32%)','VEC':'Họ VEC (cắt veneer): lịch sử VEC04 2 người (4 lần)'}
rows=[];NOTE={};ref=[]
for d,ar,mc,mn,fl,fr,to,h in rep:
    notes='-'.join(x for x in (mc,mn,fl) if x)
    hs=hist.get(mc);mapv=mp.get(mc)
    hm=hs[4] if hs else None;mv=mapv['actual_ops'] if mapv else None
    fill=hm if (hm is not None and mv is not None and int(hm)==int(mv)) else None
    if fill: src=f'Lịch sử A10 ({hs[3]} lần: {hs[6]}) và file HC Mapping cùng {int(fill)} người (chưa có tổ trưởng xác nhận)'
    elif mc and (hm or mv): src=f'Lịch sử: {hs[6] if hs else "chưa từng có"}; Mapping thực tế: {mv if mv is not None else "không có"} - lệch/chỉ một nguồn nên để trống'
    elif mc: src='Chưa có nguồn'
    else: src='Báo cáo không ghi mã máy. '+('Họ CNCR lịch sử: 2 người (68%), 3 người (32%)' if 'CNC' in mn else 'Chưa có nguồn')
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
ws['A1']='BÁO CÁO BẢO TRÌ 2/11-8/11 (STOP LINE + REPAIR WW) VÀ NGUỒN SỐ NGƯỜI';ws['A1'].font=Font(name=F,bold=True,size=13)
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
ws.cell(r+3,1,'Ghi chú: báo cáo ghi mã máy "DECO6" (chữ O), đã đọc là DEC06 (máy cắt 2 đầu băng tải, kế hoạch bảo trì W45 có DEC06). DEC06 ngày 5/11 ghi khu vực Machining, ngày 8/11 ghi RC: giữ theo báo cáo, cần xác nhận khu vực. Dòng STT 162 (3/11, RC) không có tên máy, giờ hay hiện tượng nên không nhập. 8/11 là Chủ nhật.').font=Font(name=F,italic=True)
for col,w in zip('ABCDEFGHIJKL',[12,12,10,30,34,8,8,12,12,10,18,18]): ws.column_dimensions[col].width=w
wb.save('Nhap_lieu_A10_theo_bo_phan_02-11_den_08-11.xlsx');print('dòng',len(rows),'đã điền',nfill,'tổng giờ máy',round(sum(x[7] for x in rep),2))
for x in rows:print(x[0].strftime('%d/%m'),x[2][:34],x[3],x[7])
