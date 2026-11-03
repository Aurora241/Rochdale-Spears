import pickle,pandas as pd,datetime as dt,openpyxl,collections
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
df=pickle.load(open('a10_df.pkl','rb'))
c=df[df.clean&(df.family!='BC-LINE')].copy()      # bỏ dòng cả ngày/trùng và băng chuyền Assembly (HC 11-12)
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
red=PatternFill('solid',fgColor='C00000');hf=Font(name=F,bold=True,color='FFFFFF');ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
grn=PatternFill('solid',fgColor='E2EFDA');yel=PatternFill('solid',fgColor='FFF2CC');org=PatternFill('solid',fgColor='FCE4D6')
def stat(x):
    vc=x.hc.value_counts();mode=int(vc.index[0]);share=vc.iloc[0]/len(x)
    dist=', '.join(f'{int(k)} người x{v}' for k,v in sorted(vc.items()))
    return len(x),mode,share,dist,round(x.hc.mean(),2)
def conf(n,share,nf,sharef):
    if n>=3 and share>=0.8: return 'Cao'
    if n>=1 and share==1.0 and n<3: return 'Trung bình'
    if n==0 and nf>=5 and sharef>=0.8: return 'Trung bình (theo họ máy)'
    if n>=3 and nf>=5 and sharef>=0.8: return 'Trung bình (máy dao động, họ máy ổn định)'
    if n>=3: return 'Thấp (dao động)'
    return 'Thấp'
# bảng theo mã máy
rows=[]
for code,x in c.groupby('code'):
    if code=='Không có mã máy': continue
    n,mode,share,dist,mean=stat(x)
    rows.append((code,x.family.iloc[0],' / '.join(sorted(set(x.area))),n,mode,share,dist,mean,int(x.hc.min()),int(x.hc.max()),x.date.max().strftime('%d/%m/%Y'),conf(n,share,0,0)))
rows.sort(key=lambda r:(-r[3],r[0]))
fam={}
for f,x in c.groupby('family'):
    n,mode,share,dist,mean=stat(x);fam[f]=(n,mode,share,dist)
# 19 sự cố + WBS20
inc=[('19/10','Machining','CNCR02','Máy CNC','Lỗi nút bấm',0.78),('19/10','Machining','EDB01','Máy dán cạnh thẳng tự động','Lỗi hộp số',1.00),('20/10','Raw Cutting','','Máy bơm sơn chuyền UV','Máy bơm sơn không hoạt động',0.95),('20/10','Machining','EDB05','Máy dán cạnh thẳng tự động','Lỗi nồi keo',3.00),('20/10','Machining','CNCR09','Máy chép hình thẳng','Lỗi trục không quay',2.00),('21/10','Raw Cutting','VEC04','Máy cắt veneer','Máy tự cắt',1.42),('22/10','Machining','CNCR20','Máy CNC','Gãy giá đỡ dây điện',0.85),('24/10','Machining','CNCR11','Hogmag','Lỗi hút chân không',0.22),('25/10','Raw Cutting','VC14','Máy hút chân không','Lỗi hút chân không',1.73),('26/10','Raw Cutting','EDB01','Máy dán cạnh thẳng tự động','Lỗi không đốt nhiệt độ',1.00),('26/10','Raw Cutting','HPTT02','Máy cảo cong','Lỗi không đốt nhiệt độ',1.00),('28/10','Raw Cutting','VES04','Máy may veneer','Có tiếng kêu lạ khi hoạt động',1.20),('28/10','Machining','CNCR20','Máy CNC 5 trục','Lỗi chương trình',1.55),('29/10','Machining','EDB05','Máy dán cạnh thẳng tự động','Lỗi nồi keo',1.25),('29/10','Raw Cutting','VES07','Máy may veneer','Không hoạt động',0.47),('29/10','Machining','VC','Máy hút chân không','',1.15),('29/10','Machining','CUFS02','Máy cắt','Kêu tiếng lớn khi hoạt động',1.42),('29/10','Machining','','Máy CNC Nesting','Lỗi nút dừng khẩn cấp',0.75),('30/10','Raw Cutting','VES07','Máy may veneer','Máy bị xì hơi',0.18),('28-30/10','Machining','WBS20','Máy nhám thùng','Bị hư (3 ngày)',25.0)]
famof={'CNCR02':'CNCR','CNCR09':'CNCR','CNCR20':'CNCR','CNCR11':'CNCR','EDB01':'EDB','EDB05':'EDB','VEC04':'VEC','VES04':'VES','VES07':'VES','CUFS02':'CUFS','WBS20':'WBS','VC14':'VC','HPTT02':'HPTT','VC':'VC'}
out=[]
for d,ar,code,name,fail,h in inc:
    fm=famof.get(code,'CNCR' if 'Nesting' in name else '')
    x=c[c.code==code] if code else c.iloc[0:0]
    n,mode,share,dist,_=stat(x) if len(x) else (0,None,0,'chưa từng có',0)
    nf,modef,sharef,distf=fam.get(fm,(0,None,0,''))
    cf=conf(n,share,nf,sharef)
    if code=='' and 'UV' in name: cf='Thấp'; n,mode,dist=0,None,'chưa từng có (chuyền UV khác: 2 người, 4 người, 4 người)'; modef=None
    prop=mode if n>=1 else (modef if cf.startswith('Trung bình (theo họ') else None)
    if cf.startswith('Trung bình (máy dao động'): prop=modef
    out.append((d,ar,code or '(không mã)',name,fail,h,n,dist,share if n else None,fm,f'{nf} dòng: {distf}' if nf else '',prop,cf))
pickle.dump(out,open('a10_hc_out.pkl','wb'))
wb=openpyxl.Workbook();wb.remove(wb.active)
def table(ws,r0,headers,data,widths,fmt=None,fills=None):
    for i,h in enumerate(headers,1):
        x=ws.cell(r0,i,h);x.fill=red;x.font=hf;x.alignment=ctr;x.border=bd
    for k,row in enumerate(data):
        for i,v in enumerate(row,1):
            x=ws.cell(r0+1+k,i,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='left' if isinstance(v,str) and len(v)>16 else 'center',vertical='center',wrap_text=True)
            if fmt and i in fmt: x.number_format=fmt[i]
    for i,w in enumerate(widths): ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=w
ws=wb.create_sheet('1. Đề xuất cho 19 sự cố');ws['A1']='SỐ NGƯỜI ĐỀ XUẤT CHO 19 SỰ CỐ 19/10-30/10 VÀ WBS20 (DỰA TRÊN LỊCH SỬ A10, CHỈ ĐỂ XÁC NHẬN, KHÔNG PHẢI SỐ NHẬP)';ws['A1'].font=Font(name=F,bold=True,size=13)
ws['A2']='Cột "Độ tin cậy": Cao = máy có từ 3 sự cố trở lên và từ 80% cùng một số người; Trung bình = 1-2 sự cố cùng số người hoặc cả họ máy ổn định; Thấp = chưa có lịch sử hoặc dao động. Chỉ nhập sau khi tổ trưởng xác nhận.';ws['A2'].font=Font(name=F,italic=True,color='C00000')
H=['Ngày','Khu vực','Mã máy','Tên máy','Hiện tượng','Giờ máy dừng','Số sự cố cùng máy trong lịch sử','Số người đã ghi (cùng máy)','Tỉ lệ cùng số người','Họ máy','Họ máy: số người đã ghi','Số người đề xuất (để xác nhận)','Độ tin cậy']
table(ws,4,H,out,[10,12,12,26,28,10,12,34,10,8,36,12,22],{9:'0%',6:'0.00'})
for k,row in enumerate(out):
    f=grn if row[12]=='Cao' else (yel if row[12].startswith('Trung') else org)
    ws.cell(5+k,13).fill=f
cnt=collections.Counter(r[12] for r in out)
ws.cell(6+len(out),1,'Tóm tắt: '+'; '.join(f'{k}: {v} sự cố' for k,v in cnt.items())).font=Font(name=F,bold=True)
ws2=wb.create_sheet('2. Số người theo mã máy');ws2['A1']='SỐ NGƯỜI A10 THEO MÃ MÁY (lịch sử 29/12/2025 đến 11/10, đã bỏ dòng cả ngày, dòng trùng và băng chuyền Assembly)';ws2['A1'].font=Font(name=F,bold=True,size=13)
table(ws2,3,['Mã máy','Họ máy','Khu vực','Số sự cố','Số người hay gặp nhất','Tỉ lệ','Phân bố','Trung bình','Ít nhất','Nhiều nhất','Sự cố gần nhất','Độ tin cậy'],rows,[12,10,22,10,12,10,36,10,8,8,14,18],{6:'0%'})
for k,row in enumerate(rows):
    ws2.cell(4+k,12).fill=grn if row[11]=='Cao' else (yel if row[11].startswith('Trung') else org)
ws3=wb.create_sheet('3. Theo họ máy');ws3['A1']='SỐ NGƯỜI THEO HỌ MÁY (dùng khi máy chưa có lịch sử)';ws3['A1'].font=Font(name=F,bold=True,size=13)
fr=[(f,v[0],v[1],v[2],v[3]) for f,v in sorted(fam.items(),key=lambda kv:-kv[1][0])]
table(ws3,3,['Họ máy','Số sự cố','Số người hay gặp nhất','Tỉ lệ','Phân bố'],fr,[14,10,16,10,40],{4:'0%'})
wb.save('A10_so_nguoi_theo_may.xlsx')
print(cnt)
for r in out: print(r[0],r[2],'|',r[6],'|',r[7],'|',r[9],r[10][:40],'|',r[11],'|',r[12])
