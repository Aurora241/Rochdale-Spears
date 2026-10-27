import pickle,pandas as pd,numpy as np,datetime as dt,openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.chart import BarChart,LineChart,Reference
df=pickle.load(open('a10_df.pkl','rb'))
c=df[df.clean].copy()
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
red=PatternFill('solid',fgColor='C00000');hf=Font(name=F,bold=True,color='FFFFFF');ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
yel=PatternFill('solid',fgColor='FFF2CC');org=PatternFill('solid',fgColor='FCE4D6')
wb=openpyxl.Workbook();wb.remove(wb.active)
def table(ws,r0,headers,rows,widths=None,fmt=None):
    for i,h in enumerate(headers,1):
        x=ws.cell(r0,i,h);x.fill=red;x.font=hf;x.alignment=ctr;x.border=bd
    for k,row in enumerate(rows):
        for i,v in enumerate(row,1):
            x=ws.cell(r0+1+k,i,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='left' if isinstance(v,str) and len(str(v))>14 else 'center',vertical='center',wrap_text=True)
            if fmt and i in fmt: x.number_format=fmt[i]
    if widths:
        for i,w in enumerate(widths): ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=w
    return r0+len(rows)
def title(ws,txt,sub=None):
    ws['A1']=txt;ws['A1'].font=Font(name=F,bold=True,size=14)
    if sub: ws['A2']=sub;ws['A2'].font=Font(name=F,italic=True)
tot_lab=df.labor.sum();tot_c=c.labor.sum();tot_h=df.hours.sum();tot_hc=c.hours.sum()
# 1 tổng quan
ws=wb.create_sheet('1. Tổng quan');title(ws,'PHÂN TÍCH MÃ A10 - MACHINE/TOOLING FAILURE','Nguồn: file Total loss time (29/12/2025-3/10/2026, 186 dòng) + hệ thống (4/10-11/10, 5 dòng). Không gồm báo cáo bảo trì 19-25/10 (chưa có số công nhân).')
rows=[('Số sự cố ghi nhận',len(df),'186 dòng file + 5 dòng hệ thống'),
('Dòng bất thường (cả ngày 00:00-23:59 hoặc trùng)',int((~df.clean).sum()),'8 dòng máy dừng cả ngày (1 dòng trong đó cũng là dòng trùng) và 1 dòng trùng khác. Đã tách ra phân tích riêng'),
('Giờ công A10 theo số ghi',round(tot_lab,1),'Gồm cả dòng bất thường'),
('Giờ công A10 sau khi loại dòng bất thường',round(tot_c,1),f'Dòng bất thường chiếm {100*(1-tot_c/tot_lab):.0f}% số giờ công ghi'),
('Giờ máy dừng (sau khi loại bất thường)',round(tot_hc,1),f'Trung bình {tot_hc/43:.1f} giờ/tuần (43 tuần)'),
('Thời gian dừng mỗi sự cố (trung bình / trung vị)',f'{c.hours.mean():.1f} giờ / {c.hours.median():.1f} giờ','Nhiều sự cố ngắn, vài sự cố rất dài'),
('Số công nhân bình quân mỗi sự cố',round(c.labor.sum()/c.hours.sum(),2),'= giờ công / giờ máy dừng (Assembly băng chuyền 11-12 người kéo lên)'),
('Sự cố mỗi tuần',round(len(c)/43,1),'Trung bình 43 tuần của năm 2026'),
('Tỉ trọng A10 trong tổng loss time',f'{100*df[df.src.str.startswith("File")].labor.sum()/140012:.1f}%','Tính trên 140.012 giờ công của file đến 3/10'),
]
r=table(ws,4,['Chỉ số','Giá trị','Ghi chú'],rows,[48,22,80])
r+=2
find=['1. A10 nhỏ về tổng nhưng bị nhiễu: hai nhóm dòng "máy hỏng cả ngày" (CNCR14 tháng 9, CNCR20 ngày 4/4, CNCR01 ngày 4/4) chiếm khoảng 485 giờ công, 35% tổng A10 ghi nhận. Nếu giữ nguyên thì tháng 9 bị thổi phồng (394 giờ thay vì 53 giờ).',
'2. Nhóm máy CNC router (CNCR) là trọng tâm: 52 sự cố, 343 giờ công (sau khi loại bất thường), 37% giờ công A10. Ba máy cần chú ý là CNCR01, CNCR03, CNCR14.',
'3. CNCR01 hỏng nhiều và kéo dài: 17 sự cố, 216 giờ công, trong đó đợt tháng 1 liên tục 6 ngày (19-24/1) mỗi ngày 12-18 giờ, nguyên nhân thiên về điện - điều khiển và dao.',
'4. Băng chuyền Assembly (line 4, line 1) ít sự cố nhưng giờ công rất lớn vì 11-12 người đứng chờ mỗi lần (ví dụ 7/8: 7 giờ x 11 người = 77 giờ công).',
'5. 37% dòng chỉ ghi chung chung ("máy hư", "không hoạt động") nên không biết nguyên nhân gốc. Nên yêu cầu bảo trì ghi hiện tượng cụ thể.',
'6. Chỉ có 8 dòng không có mã máy, và vài mã ghi lệch (CNCR014 và CNCR14). Cần chuẩn hóa mã máy khi nhập.',
'7. Dữ liệu nhập A10 đang trễ: tuần 43 bảo trì báo 9 sự cố (11,95 giờ máy dừng) trong khi hệ thống chưa có dòng A10 nào sau 11/10. Tuần 43 cao hơn trung bình 8,9 giờ/tuần.']
ws.cell(r,1,'PHÁT HIỆN CHÍNH').font=Font(name=F,bold=True,size=12)
for k,tx in enumerate(find):
    x=ws.cell(r+1+k,1,tx);x.font=Font(name=F);x.alignment=Alignment(wrap_text=True,vertical='top');ws.merge_cells(start_row=r+1+k,start_column=1,end_row=r+1+k,end_column=3);ws.row_dimensions[r+1+k].height=34
# 2 pareto
ws=wb.create_sheet('2. Pareto máy');title(ws,'PARETO THEO MÁY (giờ công, đã loại dòng bất thường)','Mã máy chuẩn hóa: bỏ số 0 thừa (CNCR014 = CNCR14). Dòng cả ngày và trùng nằm ở sheet 9.')
g=df.groupby('code').agg(family=('family','first'),n=('labor','size'),h=('hours_clean','sum'),lab=('labor_clean','sum'),mx=('hours_clean','max'),n_full=('fullday','sum'),lab_all=('labor','sum')).sort_values('lab',ascending=False)
g['pct']=g.lab/g.lab.sum();g['cum']=g.pct.cumsum()
rows=[(i,row.family,int(row.n),round(row.h,1),round(row.lab,1),round(row.mx,1),int(row.n_full),round(row.lab_all,1),row.pct,row.cum) for i,row in g.iterrows()]
r=table(ws,4,['Mã máy','Họ máy','Số sự cố','Giờ máy dừng','Giờ công','Dừng lâu nhất (giờ)','Số dòng cả ngày','Giờ công kể cả dòng bất thường','Tỉ trọng','Cộng dồn'],rows,[18,10,10,12,12,14,12,20,10,10],{9:'0.0%',10:'0.0%'})
ch=BarChart();ch.type='col';ch.title='Top 10 máy theo giờ công (sau khi loại bất thường)';ch.height=9;ch.width=20
ch.add_data(Reference(ws,min_col=5,min_row=4,max_row=14),titles_from_data=True);ch.set_categories(Reference(ws,min_col=1,min_row=5,max_row=14));ws.add_chart(ch,'L4')
# 3 họ máy
ws=wb.create_sheet('3. Họ máy');title(ws,'THEO HỌ MÁY')
fam=df.groupby('family').agg(n=('labor','size'),h=('hours_clean','sum'),lab=('labor_clean','sum'),lab_all=('labor','sum'),machines=('code','nunique')).sort_values('lab',ascending=False)
names={'CNCR':'CNC router','VES':'Máy may veneer','VEC':'Máy cắt veneer','WBS':'Nhám thùng','HOT':'Ép nóng','HFG':'(HFG)','BEMS':'Cắt MDF/Beam saw','MMOT':'Máy mộng','4SP':'Bào 4 mặt','BC-LINE':'Băng chuyền Assembly','EDB':'Dán cạnh','DEC':'Cắt 2 đầu','CUFS':'(CUFS)','SLTS':'Cắt bàn','SET':'Đánh mộng','Không mã':'Không có mã máy'}
rows=[(i,names.get(i,''),int(r_.machines),int(r_.n),round(r_.h,1),round(r_.lab,1),round(r_.lab_all,1)) for i,r_ in fam.iterrows()]
table(ws,3,['Họ máy','Tên (suy ra từ mô tả)','Số máy','Số sự cố','Giờ máy dừng','Giờ công','Giờ công kể cả bất thường'],rows,[12,24,10,10,14,12,22])
# 4 nguyên nhân
ws=wb.create_sheet('4. Nguyên nhân');title(ws,'PHÂN LOẠI HIỆN TƯỢNG HỎNG (theo từ khóa trong cột Lí do)','Phân loại gần đúng bằng từ khóa. 37% dòng chỉ ghi chung chung nên chưa biết nguyên nhân.')
cg=df.groupby('cat').agg(n=('labor','size'),h=('hours_clean','sum'),lab=('labor_clean','sum')).sort_values('lab',ascending=False)
rows=[(i,int(r_.n),round(r_.n/len(df),3),round(r_.h,1),round(r_.lab,1)) for i,r_ in cg.iterrows()]
table(ws,4,['Nhóm hiện tượng','Số sự cố','Tỉ lệ số sự cố','Giờ máy dừng','Giờ công'],rows,[40,10,14,14,12],{3:'0.0%'})
# 5 tháng
ws=wb.create_sheet('5. Theo tháng');title(ws,'THEO THÁNG')
mg=df.groupby('month').agg(n=('labor','size'),lab=('labor','sum'),h=('hours_clean','sum'),lab_c=('labor_clean','sum'),n_full=('fullday','sum'))
rows=[(i,int(r_.n),round(r_.lab,1),round(r_.lab_c,1),round(r_.h,1),int(r_.n_full)) for i,r_ in mg.iterrows()]
r=table(ws,3,['Tháng','Số sự cố','Giờ công theo số ghi','Giờ công sau khi loại bất thường','Giờ máy dừng (loại bất thường)','Số dòng cả ngày'],rows,[12,10,20,24,22,14])
ch=BarChart();ch.type='col';ch.title='Giờ công A10 theo tháng: số ghi và sau khi loại bất thường';ch.height=9;ch.width=20
ch.add_data(Reference(ws,min_col=3,max_col=4,min_row=3,max_row=r),titles_from_data=True);ch.set_categories(Reference(ws,min_col=1,min_row=4,max_row=r));ws.add_chart(ch,'H3')
# 6 tuần
ws=wb.create_sheet('6. Theo tuần');title(ws,'THEO TUẦN ISO (đã loại dòng bất thường)')
wg=c.groupby('isoweek').agg(n=('labor','size'),h=('hours','sum'),lab=('labor','sum'))
rows=[(f'W{int(i):02d}',int(r_.n),round(r_.h,1),round(r_.lab,1)) for i,r_ in wg.iterrows()]
table(ws,3,['Tuần','Số sự cố','Giờ máy dừng','Giờ công'],rows,[10,10,14,12])
# 7 khu vực, thứ, giờ
ws=wb.create_sheet('7. Khu vực - thứ - giờ');title(ws,'THEO KHU VỰC, THỨ, GIỜ BẮT ĐẦU (đã loại bất thường)')
ag=c.groupby('area').agg(n=('labor','size'),h=('hours','sum'),lab=('labor','sum'))
r=table(ws,3,['Khu vực','Số sự cố','Giờ máy dừng','Giờ công'],[(i,int(r_.n),round(r_.h,1),round(r_.lab,1)) for i,r_ in ag.iterrows()],[26,12,14,12])
wn=['Thứ Hai','Thứ Ba','Thứ Tư','Thứ Năm','Thứ Sáu','Thứ Bảy','Chủ nhật']
wg2=c.groupby('weekday').agg(n=('labor','size'),lab=('labor','sum'))
r=table(ws,r+3,['Thứ','Số sự cố','Giờ công'],[(wn[int(i)],int(r_.n),round(r_.lab,1)) for i,r_ in wg2.iterrows()])
hg=c.groupby('start_hour').agg(n=('labor','size'),lab=('labor','sum'))
table(ws,r+3,['Giờ bắt đầu','Số sự cố','Giờ công'],[(f'{int(i):02d}h',int(r_.n),round(r_.lab,1)) for i,r_ in hg.iterrows()])
# 8 lặp lại
ws=wb.create_sheet('8. Máy hỏng lặp lại');title(ws,'MÁY HỎNG LẶP LẠI (từ 4 sự cố trở lên)')
rows=[]
for code,x in df.groupby('code'):
    if len(x)>=4 and code!='Không có mã máy':
        d=x.date.drop_duplicates().sort_values();gap=d.diff().dt.days.dropna()
        rows.append((code,len(x),x.date.min().strftime('%d/%m/%Y'),x.date.max().strftime('%d/%m/%Y'),round(gap.mean(),1) if len(gap) else '',round(x.hours_clean.mean(),2),round(x.labor_clean.sum(),1),x.cat.value_counts().index[0],int(x.fullday.sum())))
rows.sort(key=lambda z:-z[6])
table(ws,3,['Mã máy','Số sự cố','Sự cố đầu','Sự cố cuối','Khoảng cách trung bình giữa 2 lần hỏng (ngày)','Giờ dừng trung bình','Giờ công','Nhóm hiện tượng nhiều nhất','Số dòng cả ngày'],rows,[14,10,14,14,22,14,12,36,12])
# 9 bất thường
ws=wb.create_sheet('9. Dòng bất thường');title(ws,'DÒNG CẦN XEM LẠI','Cam = máy dừng cả ngày 00:00-23:59 (giờ công = HC x 24 giờ, phồng so với ca làm thật). Vàng = trùng hoặc dữ liệu thiếu.')
bad=df[(~df.clean)|(df.code=='Không có mã máy')|(~df.dept.str.upper().eq('MAINTENANCE'))].sort_values('date')
rows=[]
for _,r_ in bad.iterrows():
    why=[]
    if r_.fullday: why.append('Máy dừng cả ngày 00:00-23:59')
    if r_.dup: why.append('Trùng dòng')
    if r_.code=='Không có mã máy': why.append('Không có mã máy')
    if str(r_.dept).upper()!='MAINTENANCE': why.append(f'Bộ phận ghi "{r_.dept}"')
    rows.append((r_.date.strftime('%d/%m/%Y'),r_.area,r_.code,r_.text[:70],r_.fr.strftime('%H:%M') if hasattr(r_.fr,'strftime') else '',r_.to.strftime('%H:%M') if hasattr(r_.to,'strftime') else '',r_.hc,round(r_.labor,1),'; '.join(why)))
r=table(ws,4,['Ngày','Khu vực','Mã máy','Lí do','From','To','HC','Giờ công','Vấn đề'],rows,[12,14,16,58,8,8,6,10,40])
for k,rr in enumerate(rows):
    if 'cả ngày' in rr[8]:
        for cc in range(1,10): ws.cell(5+k,cc).fill=org
    else:
        for cc in range(1,10): ws.cell(5+k,cc).fill=yel
# 10 báo cáo bảo trì
ws=wb.create_sheet('10. Bảo trì 19-25.10');title(ws,'BÁO CÁO BẢO TRÌ 19/10-25/10 SO VỚI DỮ LIỆU A10 TRÊN HỆ THỐNG','Số công nhân: CHƯA có (chờ tổ trưởng xác nhận), nên chưa tính giờ công.')
rep=[('19/10','Machining','CNCR02','Máy CNC','Lỗi nút bấm','07:43','08:30',0.78),('19/10','Machining','EDB01','Máy dán cạnh thẳng tự động','Lỗi hộp số','18:45','19:45',1.00),('20/10','Machining','EDB05','Máy dán cạnh thẳng tự động','Lỗi nồi keo','18:00','21:00',3.00),('20/10','Machining','CNCR09','Máy chép hình thẳng','Lỗi trục không quay','17:50','19:50',2.00),('20/10','Raw Cutting','','Máy bơm sơn chuyền UV','Máy bơm sơn không hoạt động','09:01','09:58',0.95),('21/10','Raw Cutting','VEC04','Máy cắt veneer','Máy tự cắt','18:35','20:00',1.42),('22/10','Machining','CNCR20','Máy CNC','Gãy giá đỡ dây điện','16:49','17:40',0.85),('24/10','Machining','CNCR11','Hogmag','Lỗi hút chân không','07:40','07:53',0.22),('25/10','Raw Cutting','VC14','Máy hút chân không','Lỗi hút chân không','08:25','10:09',1.73)]
r=table(ws,4,['Ngày','Khu vực','Mã máy','Tên máy','Hiện tượng','Từ','Đến','Giờ máy dừng','Đã có trong hệ thống?'],[x+('Chưa có',) for x in rep],[10,14,10,30,34,8,8,14,22])
ws.cell(r+1,7,'Tổng').font=Font(name=F,bold=True);ws.cell(r+1,8,round(sum(x[7] for x in rep),2)).font=Font(name=F,bold=True)
ws.cell(r+3,1,'So với trung bình 2026: 4,2 sự cố/tuần và 8,9 giờ máy dừng/tuần. Tuần 43: 9 sự cố, 11,95 giờ, cao hơn khoảng 2 lần số sự cố. Dòng 25/10 là Chủ nhật. EDB05 dừng đến 21:00, quá giờ tăng ca.').font=Font(name=F,italic=True)
# 11 chi tiết
ws=wb.create_sheet('11. Dữ liệu chi tiết')
rows=[(r_.date.strftime('%d/%m/%Y'),r_.src,r_.area,r_.code,r_.family,r_['cat'],r_.text,r_.fr.strftime('%H:%M') if hasattr(r_.fr,'strftime') else '',r_.to.strftime('%H:%M') if hasattr(r_.to,'strftime') else '',round(r_.hours,2),r_.hc,round(r_.labor,2),'Cả ngày' if r_.fullday else ('Trùng' if r_.dup else '')) for _,r_ in df.sort_values('date').iterrows()]
table(ws,1,['Ngày','Nguồn','Khu vực','Mã máy (chuẩn hóa)','Họ máy','Nhóm hiện tượng','Lí do gốc','From','To','Giờ dừng','HC','Giờ công','Cờ'],rows,[12,20,12,14,10,34,60,8,8,10,6,10,10])
ws.freeze_panes='A2';ws.auto_filter.ref=f'A1:M{len(rows)+1}'
wb.save('Phan_tich_A10.xlsx');print('ok')
