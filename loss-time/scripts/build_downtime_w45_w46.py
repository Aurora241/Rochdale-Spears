"""File Machining_Assy_Downtime_W46.xlsx theo mẫu 'Machining & Assy Downtime' (STOP LINE + REPAIR WW) cho tuần W45 + W46 (2-15/11/2026).
Chạy trong thư mục outputs. Dữ liệu: báo cáo bảo trì Stop line + Repair WW (STT 156-190)."""
import openpyxl,datetime as dt,collections
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.chart import BarChart,LineChart,PieChart,Reference,Series
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
ORG=PatternFill('solid',fgColor='F8CBAD');BLU=PatternFill('solid',fgColor='9DC3E6');GRN=PatternFill('solid',fgColor='E2EFDA');GRY=PatternFill('solid',fgColor='D9D9D9')
HF=Font(name=F,bold=True);NF=Font(name=F)
CEN=Alignment(horizontal='center',vertical='center',wrap_text=True);LFT=Alignment(horizontal='left',vertical='center',wrap_text=True)
T=lambda h,m=0:dt.time(h,m);D=lambda d:dt.datetime(2026,11,d)
PIC='Mr.Phong-\nMr.Khiêm - Maint'
# (STT, ngày, bộ phận, mã máy, tên máy, vấn đề, từ, đến, time run)
rep=[
(156,2,'Machining','EDB05','Máy dán cạnh thẳng tự động','Máy bị lỗi cỡ lên xuống',T(8,20),2,T(9,0),None,12),
(157,2,'Machining','','Máy CNC 5 trục','Máy bị lỗi chương trình',T(12,48),2,T(13,3),None,12),
(158,2,'RC','VET02','Máy thẩm cạnh Veneer','Dừng máy vệ sinh hút bụi',T(14,14),2,T(14,38),None,12),
(159,2,'RC','VEC03','Máy cắt veneer','Dừng máy vệ sinh hút bụi',T(14,14),2,T(14,38),None,12),
(160,2,'RC','','Máy chuyền UV','Đèn thủy ngân bị lỗi',T(17,45),2,T(18,25),None,12),
(161,3,'RC','VEC04','Máy cắt veneer','Máy bị lỗi tự động cắt',T(18,10),3,T(18,50),None,12),
(163,4,'RC','STS01','Máy chà nhám băng/ Hand Stroke Sanding','Con lăn của máy không di chuyển được',T(13,20),4,T(13,38),None,12),
(164,4,'RC','VES09','Máy may veneer','Máy bị hư dây cu-roa',T(16,30),4,T(18,1),None,12),
(165,4,'Machining','','Máy CNC Nesting (Panel)','Máy bị lỗi nút dừng khẩn cấp',T(19,7),4,T(19,37),None,12),
(166,5,'Machining','DEC06','Máy cắt 2 đầu băng tải / Double End Cutting','Máy bị lỗi không lên nguồn',T(8,15),5,T(13,55),None,12),
(167,5,'Machining','CUFS19','Máy cắt','Máy bị hư dây cu-roa',T(13,6),5,T(13,50),None,12),
(168,5,'RC','VES08','Máy may veneer','Máy không hoạt động',T(18,9),5,T(19,16),None,12),
(169,6,'Machining','CNCR11','Máy CNC','Máy bị bể ống hơi',T(6,40),6,T(7,18),None,12),
(171,6,'RC','VES02','Máy may veneer','Máy bị lỗi con lăn',T(7,38),6,T(8,10),None,12),
(172,6,'RC','VES04','Máy may veneer','Máy bị lỗi băng truyền động',T(9,0),6,T(10,23),None,12),
(173,6,'Machining','','Máy CNC 5 trục','Máy bị lệch trục C',T(10,28),6,T(11,9),None,12),
(174,6,'RC','BEMS04','Máy cắt MDF','Máy bị đứt dây Cu-roa',T(10,25),6,T(13,55),None,12),
(175,6,'Machining','CNCR11','Máy CNC','Máy bị lỗi',T(17,8),6,T(17,59),None,12),
(176,6,'RC','UVDR01','Chuyền UV','Đèn UV bị lỗi',T(17,47),6,T(18,4),None,12),
(177,7,'Machining','CNCR01','Máy Hogmang 1','Máy bị bung ốc',T(9,35),7,T(10,50),None,12),
(179,8,'RC','DEC06','Máy cắt 2 đầu băng tải / Double End Cutting','Máy bị lỗi không lên nguồn',T(15,45),8,T(16,15),None,12),
(180,9,'Machining','CNNR01','Máy CNC 5 trục','Máy bị lỗi chương trình',T(9,18),9,T(9,52),None,12),
(181,9,'Machining','DECO5','Máy cắt hai đầu băng tải','Máy không ra vo băng tải được',T(11,6),9,T(13,41),None,12),
(182,9,'Machining','ADET02','Máy đánh mộng dương 2 đầu','Máy bị hư',T(8,50),11,T(16,29),32.65,12),
(183,11,'Machining','WBS09','Máy nhám thùng','Máy bị lỗi chương trình',T(7,41),11,T(9,13),None,12),
(184,11,'Machining','CNCR14','Máy CNC 5 trục','Ống hút bụi bị hư',T(18,0),11,T(18,45),None,12),
(185,12,'Machining','MOT01','Máy đục đơn','Máy bị hư',T(18,41),12,T(19,23),None,12),
(186,12,'Machining','CORCO1','Máy phay ngàm','Máy bị hư',T(19,37),13,T(10,25),3.3,12),
(187,13,'RC','STS01','Máy chà nhám băng/ Hand Stroke Sanding','Máy bị hư bạc đạn',T(13,16),13,T(16,34),None,12),
(188,13,'Machining','CNCC01','Máy CNC nhiều mũi','Máy bị lỗi trục X',T(13,50),13,T(16,34),None,12),
(189,13,'Machining','CNCR01','Máy Hogmang 1','Máy bị hư bạc đạn trục X',T(22,52),14,T(0,45),None,24),
(190,14,'Machining','CNCR03','Máy CNC shoda','Máy bị lỗi bàn đạp hơi',T(9,36),14,T(9,46),None,12),
]
wb=openpyxl.Workbook();wb.remove(wb.active)
HDR=['STT','Submit Date\n(Ngày nộp)','Department\n(Bộ phận)','Machine Code\n(Mã Máy)','Machine\n(Tên Máy)','Problem\n(Vấn đề)','Status\n(tình trạng máy lúc báo hư => không phải trình trạng máy hiện tại)','P.I.C\n(Người chịu trách nhiệm)']
def build(title,dept_filter,note=None):
    ws=wb.create_sheet(title)
    ws['A1']='STOP LINE + REPAIR WW REPORT - W45 + W46 (02/11/2026 - 15/11/2026)';ws['A1'].font=Font(name=F,bold=True,size=13)
    for i,h in enumerate(HDR,1):
        c=ws.cell(2,i,h);c.font=HF;c.fill=ORG;c.alignment=CEN;c.border=bd
        ws.merge_cells(start_row=2,start_column=i,end_row=3,end_column=i)
    for c0,h in ((9,'TIME BREAKDOWN'),(11,'TIME REPAIRED')):
        ws.cell(2,c0,h).font=HF;ws.cell(2,c0).fill=ORG;ws.cell(2,c0).alignment=CEN;ws.merge_cells(start_row=2,start_column=c0,end_row=2,end_column=c0+1)
    for i,h in enumerate(('Time','Date','Time','Date'),9):
        c=ws.cell(3,i,h);c.font=HF;c.fill=ORG;c.alignment=CEN;c.border=bd
    ws.cell(2,13,'Total').font=HF;ws.cell(2,13).fill=ORG;ws.cell(2,13).alignment=CEN
    ws.cell(3,13,'Time stop line (h)').font=HF;ws.cell(3,13).fill=ORG;ws.cell(3,13).alignment=CEN;ws.cell(3,13).border=bd
    for cc,wk in ((14,'W45'),(15,'W46')):
        ws.cell(2,cc,wk).font=HF;ws.cell(2,cc).fill=ORG;ws.cell(2,cc).alignment=CEN;ws.cell(3,cc,wk).font=HF;ws.cell(3,cc).fill=ORG
    ws.cell(2,16,'Time run').font=HF;ws.cell(2,16).fill=ORG;ws.cell(2,16).alignment=CEN;ws.cell(3,16,'(giờ máy chạy/ngày)').font=HF;ws.cell(3,16).fill=ORG;ws.cell(3,16).alignment=CEN
    ws.cell(2,17,'Note').font=HF;ws.cell(2,17).fill=ORG;ws.cell(3,17,'(Tình trạng máy hiện tại)').font=HF;ws.cell(3,17).fill=ORG
    r=4;first=r
    for stt,d,dep,mc,mn,pb,fr,d2,to,tot,run in rep:
        if dep not in dept_filter:continue
        vals=[stt,D(d),dep,mc,mn,pb,'Machine stop',PIC,fr,D(d),to,D(d2),(tot if tot is not None else (f'=(L{r}+K{r}-J{r}-I{r})*24' if d2!=d else f'=(K{r}-I{r})*24')),(f'=M{r}' if d<=8 else 0),(f'=M{r}' if d>8 else 0),run,'Done']
        for i,v in enumerate(vals,1):
            c=ws.cell(r,i,v);c.font=NF;c.border=bd;c.alignment=LFT if i in(5,6,8) else CEN
            if i in(1,)and False:pass
            c.fill=BLU
        ws.cell(r,2).number_format=ws.cell(r,10).number_format=ws.cell(r,12).number_format='d-mmm-yy';ws.cell(r,9).number_format=ws.cell(r,11).number_format='h:mm';ws.cell(r,13).number_format=ws.cell(r,14).number_format='0.00'
        ws.row_dimensions[r].height=32;r+=1
    last=r-1
    if last<first:
        ws.cell(r,1,note or 'Không có sự cố ghi nhận trong W45, W46').font=Font(name=F,italic=True);ws.cell(r,13,0);ws.cell(r,14,0);ws.cell(r,15,0);last=r;r+=1
    ws.cell(r+1,12,'Tổng').font=HF
    ws.cell(r+1,13,f'=SUM(M{first}:M{last})').font=HF;ws.cell(r+1,14,f'=SUM(N{first}:N{last})').font=HF;ws.cell(r+1,15,f'=SUM(O{first}:O{last})').font=HF;ws.cell(r+1,15).number_format='0.00'
    ws.cell(r+1,13).number_format=ws.cell(r+1,14).number_format='0.00'
    for col,w in zip('ABCDEFGHIJKLMNOPQ',[7,12,12,13,26,28,16,20,9,11,9,11,12,8,8,12,12]):ws.column_dimensions[col].width=w
    ws.freeze_panes='A4'
    return ws,first,last,r+1
rc,rf,rl,rt=build('RC',('RC',))
mac,mf,ml,mt=build('MAC',('Machining',))
asy,af,al,at=build('ASSY',(),'Không có sự cố băng chuyền/máy Assembly ghi nhận trong W45, W46')
xe,xf,xl,xt=build('XE NÂNG',(),'Không có sự cố xe nâng ghi nhận trong W45, W46')
# TOTAL (bố cục như mẫu: Month, Week W46-W48, Total, TOTAL YEAR 2026)
tt=wb.create_sheet('TOTAL');tt['A2']='TOTAL WHITE WOOD DOWNTIME';tt['A2'].font=Font(name=F,bold=True,size=13)
tt['A3']='Month';tt['B3']='NOVEMBER';tt['G3']='TOTAL YEAR 2026'
for i,w_ in enumerate(('W45','W46','W47','W48','Total')):tt.cell(4,2+i,w_)
tt['A4']='Week';tt['G4']=None
for c in ['A3','B3','G3','A4']+[f'{x}4' for x in 'BCDEF']:
    tt[c].font=HF;tt[c].fill=ORG;tt[c].alignment=CEN;tt[c].border=bd
for i,(nm,sh,rt_) in enumerate((('RC','RC',rt),('MACHING','MAC',mt),('ASSY','ASSY',at),('XE NÂNG','XE NÂNG',xt))):
    tt.cell(5+i,1,nm).font=HF;tt.cell(5+i,1).border=bd
    for cc,colL in ((2,'N'),(3,'O')):
        c=tt.cell(5+i,cc,f"='{sh}'!{colL}{rt_}");c.number_format='0.00';c.border=bd;c.font=NF
    tt.cell(5+i,6,f'=SUM(B{5+i}:E{5+i})').number_format='0.00';tt.cell(5+i,6).border=bd;tt.cell(5+i,6).font=HF
    tt.cell(5+i,7,None).border=bd
tt.cell(9,1,'TOTAL').font=HF;tt.cell(9,1).border=bd
for j in range(5):
    col=chr(66+j);c=tt.cell(9,2+j,f'=SUM({col}5:{col}8)');c.font=HF;c.number_format='0.00';c.border=bd
tt.cell(9,7).border=bd
tt['A11']='TOTAL YEAR 2026 để trống: cần downtime các tuần W01-W44 (chưa có dữ liệu). Cột W47-W48 điền khi có báo cáo các tuần sau.';tt['A11'].font=Font(name=F,italic=True)
tt['A12']='Downtime = số giờ máy dừng (Time stop line), không nhân số công nhân. RC gồm RC WOOD và RC PANEL.';tt['A12'].font=Font(name=F,italic=True)
tt.column_dimensions['A'].width=18;tt.column_dimensions['G'].width=20
# Sheet1 (biểu đồ)
s=wb.create_sheet('Sheet1')
s['B1']='            TOTAL MACHINE DOWNTIME';s['B1'].font=Font(name=F,bold=True,size=13)
for c,v in (('B3','Bộ Phận'),('C3','W45'),('D3','W46')):s[c]=v;s[c].font=HF;s[c].fill=ORG;s[c].border=bd;s[c].alignment=CEN
for i,(nm,rw) in enumerate((('RC-WOOD/PANEL',5),('MAC',6),('ASSY',7))):
    s.cell(4+i,2,nm).border=bd;s.cell(4+i,2).font=NF
    for cc,tc in ((3,'B'),(4,'C')):
        c=s.cell(4+i,cc,f'=TOTAL!{tc}{rw}');c.number_format='0.00';c.border=bd
s['B7']='Total';s['B7'].font=HF;s['B7'].border=bd
for cc,L_ in ((3,'C'),(4,'D')):
    c=s.cell(7,cc,f'=SUM({L_}4:{L_}6)');c.font=HF;c.number_format='0.00';c.border=bd
s['B14']='            TOTAL WW Machine Downtime';s['B14'].font=Font(name=F,bold=True,size=13)
s['C16']='W45';s['D16']='W46'
for c in ('C16','D16'):s[c].font=HF;s[c].fill=ORG
rows_=(('White Wood ( RC + PANEL + MAC)','=C7','=D7'),('Đường 10 giờ (theo mẫu W23)',10,10),('Đường 18 giờ (theo mẫu W23)',18,18))
for i,(n,v,v2) in enumerate(rows_):
    s.cell(17+i,2,n).font=NF;s.cell(17+i,2).border=bd
    for cc,vv in ((3,v),(4,v2)):
        c=s.cell(17+i,cc,vv);c.border=bd;c.number_format='0.00'
s['B21']='Chênh lệch so với đường 18 giờ';s['C21']='=C17-C19';s['D21']='=D17-D19'
s['B22']='Chênh lệch so với đường 10 giờ';s['C22']='=C17-C18';s['D22']='=D17-D18'
for c in ('C21','D21','C22','D22'):s[c].number_format='0.00'
s['B25']='Top máy dừng nhiều nhất W45 + W46 (giờ)';s['B25'].font=Font(name=F,bold=True)
agg=collections.defaultdict(float)
for stt,d,dep,mc,mn,pb,fr,d2,to,tot,run in rep:
    h=tot if tot is not None else (dt.datetime.combine(dt.date(2026,11,d2),to)-dt.datetime.combine(dt.date(2026,11,d),fr)).total_seconds()/3600
    agg[(mc or mn,mn)]+=h
top=sorted(agg.items(),key=lambda x:-x[1])[:6]
s['B26']='Máy';s['C26']='Giờ dừng';s['D26']='Tên máy'
for c in ('B26','C26','D26'):s[c].font=HF;s[c].fill=ORG;s[c].border=bd
for i,((k,mn),v) in enumerate(top):
    s.cell(27+i,2,k).border=bd;c=s.cell(27+i,3,round(v,2));c.number_format='0.00';c.border=bd;s.cell(27+i,4,mn).border=bd
b=BarChart();b.type='col';b.title='Thời gian dừng máy RC WOOD/PANEL/MAC/ASSY W45, W46 (hrs)';b.height=8;b.width=14
b.add_data(Reference(s,min_col=3,max_col=4,min_row=3,max_row=6),titles_from_data=True);b.set_categories(Reference(s,min_col=2,min_row=4,max_row=6));s.add_chart(b,'F2')
b2=BarChart();b2.type='col';b2.title='Downtime WW so với đường 10 / 18 giờ (hrs)';b2.height=8;b2.width=14
b2.add_data(Reference(s,min_col=3,max_col=4,min_row=16,max_row=19),titles_from_data=True);b2.set_categories(Reference(s,min_col=2,min_row=17,max_row=19));s.add_chart(b2,'F20')
p=PieChart();p.title='Downtime từng bộ phận W46';p.height=8;p.width=10
p.add_data(Reference(s,min_col=4,min_row=3,max_row=6),titles_from_data=True);p.set_categories(Reference(s,min_col=2,min_row=4,max_row=6));s.add_chart(p,'P2')
s.column_dimensions['B'].width=34;s.column_dimensions['C'].width=12;s.column_dimensions['D'].width=30
wb.move_sheet('TOTAL',offset=0)
wb.save('Machining_Assy_Downtime_W45_W46.xlsx');print('file 2 ok',sum(1 for x in rep))
