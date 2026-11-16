"""Hai file báo cáo tuần W46 (2-8/11/2026) theo mẫu W23: Loss_time_White_Wood_W46.xlsx và Machining_Assy_Downtime_W46.xlsx.
Chạy trong thư mục outputs. Dữ liệu: bảng chuẩn P20/P26/P22/A2 + S32 (đã gồm cào cước ME 2/11) + A10 từ file xuất hệ thống (có Worker Qty)."""
import sys,os;sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from fix_charts import fix_charts
import openpyxl,datetime as dt,re,sys,collections,warnings
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.chart import BarChart,LineChart,PieChart,Reference,Series
from openpyxl.chart.label import DataLabelList
warnings.simplefilter('ignore')
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
RED=PatternFill('solid',fgColor='C00000');BLU=PatternFill('solid',fgColor='1F4E78');GRY=PatternFill('solid',fgColor='D9D9D9');YEL=PatternFill('solid',fgColor='FFF2CC')
HF=Font(name=F,bold=True,color='FFFFFF');BF=Font(name=F,bold=True);NF=Font(name=F)
CEN=Alignment(horizontal='center',vertical='center',wrap_text=True)
DAYS=[dt.datetime(2026,11,d) for d in range(2,16)]
from openpyxl.utils import get_column_letter as L
dc=lambda j:2+j if j<7 else 3+j   # cột của ngày thứ j (0-13)
T=lambda h,m=0:dt.time(h,m)
# ---------------- dữ liệu
rows=[]
for fn in ('Loss_time_White_Wood_W45.xlsx','Loss_time_White_Wood_W46.xlsx'):
    w=openpyxl.load_workbook(fn,data_only=True)['Total loss time']
    for r in w.iter_rows(min_row=3,values_only=True):
        if r[1] and hasattr(r[1],'strftime') and r[7]:
            rows.append((r[1],r[2],r[3],r[4],r[5],r[6],r[7],r[8],r[9],r[10],r[11],r[12],r[13],r[15]))
nA10=sum(1 for x in rows if x[6]=='A10')
NAMEX=None
order={'P20':0,'P26':1,'P22':2,'A2':3,'S32':4,'A10':5}
rows.sort(key=lambda x:(x[0],order[x[6]]))
RCN=['RC - WOOD','RC - PANEL','Raw Cutting','Raw Cutting Panel','RC']
# ---------------- FILE 1
wb=openpyxl.Workbook();ws=wb.active;ws.title='Total loss time'
ws['A1']='LOSS TIME WHITE WOOD - TUẦN W45 + W46 (02/11/2026 - 15/11/2026)';ws['A1'].font=Font(name=F,bold=True,size=13)
HD=['Week','Date \n(Ngày)',"Item's\nMã hàng",'Code Máy Hư',"Q'ty",'Số NCR/ME','Khu vực phát sinh/bộ phận hỗ trợ','Code','Name code','Chart / bộ phận chịu trách nhiệm','Lí Do','Ghi chú','From','To','LossTime (Hr)','HC','Loss time labor (Hrs)']
for i,h in enumerate(HD,1):
    c=ws.cell(2,i,h);c.font=HF;c.fill=RED;c.alignment=CEN;c.border=bd
ws.row_dimensions[2].height=44
for k,x in enumerate(rows):
    r=3+k
    vals=[f'="W"&+WEEKNUM(B{r},21)',x[0],x[1],x[2],x[3],x[4],x[5],x[6],x[7],x[8],x[9],x[10],x[11],x[12],f'=(N{r}-M{r})*24',x[13],f'=P{r}*O{r}']
    for i,v in enumerate(vals,1):
        c=ws.cell(r,i,v);c.font=NF;c.border=bd;c.alignment=Alignment(horizontal='left' if i in(3,11,12) else 'center',vertical='center',wrap_text=i in(11,12))
    ws.cell(r,2).number_format='dd/mm/yyyy';ws.cell(r,13).number_format=ws.cell(r,14).number_format='hh:mm';ws.cell(r,15).number_format='0.00';ws.cell(r,17).number_format='0.0'
LAST=2+len(rows)
ws.cell(LAST+1,16,'Tổng').font=BF;ws.cell(LAST+1,17,f'=SUM(Q3:Q{LAST})').font=BF;ws.cell(LAST+1,17).number_format='#,##0.0'
for col,w in zip('ABCDEFGHIJKLMNOPQ',[9,12,24,12,8,10,22,8,26,24,46,40,8,8,10,6,12]):ws.column_dimensions[col].width=w
ws.freeze_panes='A3'
RNG=lambda col:f"'Total loss time'!${col}$3:${col}${LAST}"
def sumifs_area(areas,code=None,day=None,extra=''):
    terms=[]
    for a in areas:
        f=f"SUMIFS({RNG('Q')},{RNG('G')},\"{a}\""
        if day:f+=f",{RNG('B')},{day}"
        if code:f+=f",{RNG('H')},\"{code}\""
        f+=extra+')';terms.append(f)
    return '='+'+'.join(terms)
AREAS=[('RC - WOOD',['RC - WOOD']),('RC - PANEL',['RC - PANEL','Raw Cutting','Raw Cutting Panel','RC']),('Machining',['Machining']),('Assembly',['Assembly'])]  # như mẫu W23: RC - WOOD (P20, P26) và RC - PANEL (còn lại)
# --- Time part
tp=wb.create_sheet('Time part')
tp['A1']='LOSS TIME THEO BỘ PHẬN - TUẦN W45 + W46';tp['A1'].font=Font(name=F,bold=True,size=13)
tp['A2']='WEEK';tp['B2']='WEEK 45';tp['J2']='WEEK 46';tp['A3']='Bộ Phận'
for c_ in ('A2','B2','J2'):tp[c_].font=BF
for c_ in (tp.cell(3,1),):c_.font=HF;c_.fill=RED;c_.alignment=CEN;c_.border=bd
for j,d in enumerate(DAYS):
    c=tp.cell(3,dc(j),d);c.number_format='dd/mm';c.font=HF;c.fill=RED;c.alignment=CEN;c.border=bd
for col,lab in ((9,'Total W45'),(17,'Total W46'),(18,'Total W45+W46')):
    c=tp.cell(3,col,lab);c.font=HF;c.fill=RED;c.alignment=CEN;c.border=bd
for i,(nm,ar) in enumerate(AREAS):
    r=4+i;tp.cell(r,1,nm).font=NF;tp.cell(r,1).border=bd
    for j in range(14):
        c=tp.cell(r,dc(j),sumifs_area(ar,day=f'{L(dc(j))}$3'));c.number_format='#,##0.0';c.border=bd;c.font=NF
    for col,f in ((9,f'=SUM(B{r}:H{r})'),(17,f'=SUM(J{r}:P{r})'),(18,f'=I{r}+Q{r}')):
        c=tp.cell(r,col,f);c.number_format='#,##0.0';c.border=bd;c.font=BF
tp.cell(8,1,'Total').font=BF;tp.cell(8,1).border=bd
for col in range(2,19):
    c=tp.cell(8,col,f'=SUM({L(col)}4:{L(col)}7)');c.font=BF;c.border=bd;c.number_format='#,##0.0'
tp.column_dimensions['A'].width=34
for j in range(2,19):tp.column_dimensions[L(j)].width=11
for (c1,c2,ttl,anc) in ((2,8,'P2 LOST TIME W45 (giờ công)','A11'),(10,16,'P2 LOST TIME W46 (giờ công)','K11')):
    ch=BarChart();ch.type='col';ch.grouping='stacked';ch.overlap=100;ch.title=ttl
    for i in range(4):ch.series.append(Series(Reference(tp,min_col=c1,max_col=c2,min_row=4+i),title=tp.cell(4+i,1).value))
    ch.set_categories(Reference(tp,min_col=c1,max_col=c2,min_row=3));ch.height=9;ch.width=16;ch.y_axis.title='giờ công';tp.add_chart(ch,anc)
# pareto
def val(area_list,pred):  # tính trước để sắp xếp
    tot=0
    for x in rows:
        if x[5] in area_list and pred(x):
            hh=((dt.datetime.combine(dt.date.today(),x[12])-dt.datetime.combine(dt.date.today(),x[11])).seconds/3600)*x[13];tot+=hh
    return tot
def cats(ar):
    c=[('Đào tạo NV mới (P22 + A2)',lambda x:x[6] in('P22','A2'),None),('Downtime máy (A10)',lambda x:x[6]=='A10',None)]
    S=lambda k:(lambda x:x[6]=='S32' and k in (x[9] or '').lower())
    if ar=='Assembly':
        c+=[('Bắn cát Supcon (S32)',S('bắn cát'),None),('Cào cước hỗ trợ MAC (S32)',S('cào cước'),None)]
    elif ar=='Machining':
        c+=[('Hỗ trợ khác (S32: cào cước ME)',lambda x:x[6]=='S32',None)]
    else:
        c+=[('Chuyền UV (S32)',S('bomtime chuyền uv'),None),('UV phát sinh (S32)',S('phát sinh khu vực uv'),None),('Beam saw phát sinh (S32)',S('beam saw'),None),('Veneer làm mẫu (S32)',S('veneer'),None),('Ép nóng hàng sửa (S32)',S('ép hàng'),None)]
    return c
crit={'Chuyền UV (S32)':'*bomtime chuyền UV*','UV phát sinh (S32)':'Phát sinh khu vực UV*','Sơn lòng trong (S32)':'*sơn lòng*','Bắn cát Supcon (S32)':'*bắn cát*','Cào cước hỗ trợ MAC (S32)':'*cào cước*','Veneer làm mẫu (S32)':'*veneer*','Ép nóng hàng sửa (S32)':'ép hàng*','Beam saw phát sinh (S32)':'*beam saw*'}
def pareto(top,title,areas,key):
    tp.cell(top,1,title).font=Font(name=F,bold=True,size=12)
    cl=[(n,p,val(areas,p)) for n,p,_ in cats(key)];cl=[c for c in cl if c[2]>0.001];cl.sort(key=lambda c:-c[2])
    tp.cell(top+1,1,'Hạng mục').font=HF;tp.cell(top+1,1).fill=RED
    for j,(n,p,v) in enumerate(cl):
        c=tp.cell(top+1,2+j,n);c.font=HF;c.fill=RED;c.alignment=CEN;c.border=bd
    tp.cell(top+1,2+len(cl),'Total').font=HF;tp.cell(top+1,2+len(cl)).fill=RED
    tp.row_dimensions[top+1].height=44
    labs=['Losstime','%Losstime','%Cumulative']
    for i,l in enumerate(labs):tp.cell(top+2+i,1,l).font=BF;tp.cell(top+2+i,1).border=bd
    for j,(n,p,v) in enumerate(cl):
        col=chr(66+j)
        if n.startswith('Đào tạo'):f=sumifs_area(areas,code='P22')+'+'+sumifs_area(areas,code='A2')[1:]
        elif n.startswith('Downtime'):f=sumifs_area(areas,code='A10')
        elif n in crit:f=sumifs_area(areas,code='S32',extra=f",{RNG('K')},\"{crit[n]}\"")
        else:f=sumifs_area(areas,code='S32')
        tp.cell(top+2,2+j,f).number_format='#,##0.0'
        tp.cell(top+3,2+j,f'={col}{top+2}/${chr(66+len(cl))}${top+2}*100').number_format='0.0'
        tp.cell(top+4,2+j,f'=SUM($B{top+2}:{col}{top+2})/${chr(66+len(cl))}${top+2}').number_format='0.0%'
        for k in (2,3,4):tp.cell(top+k,2+j).border=bd;tp.cell(top+k,2+j).font=NF
    tc=chr(66+len(cl));tp.cell(top+2,2+len(cl),f'=SUM(B{top+2}:{chr(65+len(cl))}{top+2})').number_format='#,##0.0'
    tp.cell(top+2,2+len(cl)).font=BF;tp.cell(top+3,2+len(cl),100).font=BF
    b=BarChart();b.type='col';b.title=title;b.height=8;b.width=16
    b.add_data(Reference(tp,min_col=1,max_col=1+len(cl),min_row=top+2),from_rows=True,titles_from_data=True);b.set_categories(Reference(tp,min_col=2,max_col=1+len(cl),min_row=top+1))
    l=LineChart();l.add_data(Reference(tp,min_col=1,max_col=1+len(cl),min_row=top+4),from_rows=True,titles_from_data=True);l.y_axis.axId=200;l.y_axis.number_format='0%';l.y_axis.crosses='max'
    b.y_axis.title='giờ công';b+=l;tp.add_chart(b,f'A{top+6}')
    return top+6+18
nxt=pareto(30,'RC W45+W46 - Pareto loss time',RCN,'Raw Cutting');nxt=pareto(nxt,'MAC W45+W46 - Pareto loss time',['Machining'],'Machining');nxt=pareto(nxt,'ASSY W45+W46 - Pareto loss time',['Assembly'],'Assembly')
tp.cell(nxt,1,'Pareto không gồm mã cố định P20 (vệ sinh) và P26 (Recorder/Supporter), theo cách trình bày của mẫu W23.').font=Font(name=F,italic=True)
# --- TỔNG (nhập file Onl)
tg=wb.create_sheet('TỔNG (nhập file Onl)')
CODES=['Dailly Cleaning','Machine/Tooling Failure','NCR by ECR (process change)','Power cut','Planned Losstime','Training','HR Training New Workers','Others Losstime','Other Support of Production','Recorder/ Supporter','Change model time']
CA={'Dailly Cleaning':'P20','Machine/Tooling Failure':'A10','Training':'A2','HR Training New Workers':'P22','Other Support of Production':'S32','Recorder/ Supporter':'P26'}
tg['B1']='WEEK 45';tg['J1']='WEEK 46'
for c_ in ('B1','J1'):tg[c_].font=BF
r=2;tot_rows=[]
for nm,ar in AREAS:
    tg.cell(r,1,nm).font=HF;tg.cell(r,1).fill=RED
    for j,d in enumerate(DAYS):
        c=tg.cell(r,dc(j),d);c.number_format='dd/mm';c.font=HF;c.fill=RED;c.alignment=CEN
    for col,lab in ((9,'Total W45'),(17,'Total W46'),(18,'Total W45+W46')):
        tg.cell(r,col,lab).font=HF;tg.cell(r,col).fill=RED
    hr=r
    for k,cn in enumerate(CODES):
        rr=r+1+k;tg.cell(rr,1,cn).font=NF
        for j in range(14):
            f=sumifs_area(ar,code=CA[cn],day=f'{L(dc(j))}${hr}') if cn in CA else 0
            c=tg.cell(rr,dc(j),f);c.number_format='#,##0.0';c.font=NF
        for col,f in ((9,f'=SUM(B{rr}:H{rr})'),(17,f'=SUM(J{rr}:P{rr})'),(18,f'=I{rr}+Q{rr}')):tg.cell(rr,col,f).number_format='#,##0.0'
    rt=r+1+len(CODES);tg.cell(rt,1,'Total').font=BF
    for col in range(2,19):
        c=tg.cell(rt,col,f'=SUM({L(col)}{r+1}:{L(col)}{rt-1})');c.font=BF;c.number_format='#,##0.0'
    tot_rows.append(rt);r=rt+3
tg.cell(r,1,'Tổng White Wood').font=BF
for col in range(2,19):
    c=tg.cell(r,col,'='+'+'.join(f'{L(col)}{x}' for x in tot_rows));c.font=BF;c.number_format='#,##0.0'
tg.column_dimensions['A'].width=34
for j in range(2,19):tg.column_dimensions[L(j)].width=11
# --- Top 3 Loss
tl=wb.create_sheet('Top 3 Loss')
tl['A1']='TOP 3 LOSS TIME WHITE WOOD - W45 và W46';tl['A1'].font=Font(name=F,bold=True,size=13)
h=['Week','Downtime of machine','NCR by ECR (process change)','Other Support of Production','Others Losstime','Power cut']
for i,x in enumerate(h,1):
    c=tl.cell(2,i,x);c.font=HF;c.fill=RED;c.alignment=CEN;c.border=bd
tl.row_dimensions[2].height=40
nm={2:'Machine/Tooling Failure',3:'NCR by ECR (process change)',4:'Other Support of Production',5:'Others Losstime',6:'Power cut'}
for k,wk in enumerate(('W45','W46')):
    tl.cell(3+k,1,wk).border=bd
    for i,n in nm.items():
        c=tl.cell(3+k,i,f"=SUMIFS({RNG('Q')},{RNG('I')},\"{n}\",{RNG('A')},\"{wk}\")");c.number_format='#,##0.0';c.border=bd
tl['A6']='Tổng loss time theo mã';tl['A6'].font=BF
h2=['Week']+CODES+['Total']
for i,x in enumerate(h2,1):
    c=tl.cell(7,i,x);c.font=HF;c.fill=RED;c.alignment=CEN;c.border=bd
tl.row_dimensions[7].height=44
for k,wk in enumerate(('W45','W46')):
    rr=8+k;tl.cell(rr,1,wk).border=bd
    for i,cn in enumerate(CODES):
        c=tl.cell(rr,2+i,f"=SUMIFS({RNG('Q')},{RNG('I')},\"{cn}\",{RNG('A')},\"{wk}\")");c.number_format='#,##0.0';c.border=bd
    c=tl.cell(rr,2+len(CODES),f'=SUM(B{rr}:{L(1+len(CODES))}{rr})');c.font=BF;c.number_format='#,##0.0';c.border=bd
tl.cell(10,1,'Total').font=BF
for col in range(2,3+len(CODES)):
    c=tl.cell(10,col,f'=SUM({L(col)}8:{L(col)}9)');c.font=BF;c.number_format='#,##0.0';c.border=bd
for i in range(1,len(h2)+1):tl.column_dimensions[L(i)].width=16
b=BarChart();b.type='col';b.title='Loss time của downtime, NCR và hỗ trợ khu vực khác (W45, W46)';b.height=8;b.width=18
for i in range(2,7):b.series.append(Series(Reference(tl,min_col=i,min_row=3,max_row=4),title=tl.cell(2,i).value))
b.set_categories(Reference(tl,min_col=1,min_row=3,max_row=4));tl.add_chart(b,'A13')
b2=BarChart();b2.type='col';b2.title='Loss time theo mã (giờ công), W45 và W46';b2.height=9;b2.width=24
for i in range(2,2+len(CODES)):b2.series.append(Series(Reference(tl,min_col=i,min_row=8,max_row=9),title=tl.cell(7,i).value))
b2.set_categories(Reference(tl,min_col=1,min_row=8,max_row=9));tl.add_chart(b2,'A31')
tl['A50']='Lưu ý: chưa có NCR (O15), Others Losstime, Power cut, Kaizen P24 trong W45, W46. S32 chưa đủ nguồn (xem sheet Ghi chú).';tl['A50'].font=Font(name=F,italic=True)
# --- Loss time categoty (sao chép từ mẫu)
tc=wb.create_sheet('Loss time categoty')
tpl=openpyxl.load_workbook('/root/.claude/uploads/64423226-a456-5e96-8225-6a5ffe123efb/285a5198-Copy_of_Loss_time_White_Wood_2026.xlsx',data_only=True)['Loss time categoty']
for row in tpl.iter_rows():
    for c in row:
        if c.value is not None:
            n=tc.cell(c.row,c.column,c.value);n.font=Font(name=F,bold=c.font.b);n.alignment=Alignment(wrap_text=True,vertical='top')
for k,v in tpl.column_dimensions.items():tc.column_dimensions[k].width=v.width or 12
gn=wb.create_sheet('Ghi chú W45-W46',0)
notes=['GHI CHÚ BÁO CÁO LOSS TIME WHITE WOOD - TUẦN W45 (02-08/11) VÀ W46 (09-15/11/2026)',
'1. Gộp từ hai file tuần W45 và W46; mỗi dòng chi tiết có cột Week. Chủ nhật 8/11 chỉ Assembly làm; Chủ nhật 15/11 nghỉ cả xưởng.',
'2. Giờ công mất = (To - From) x HC. Time part tách RC - WOOD và RC - PANEL như mẫu W23; Pareto gộp hai tuần, không gồm P20 và P26.',
'3. W45: S32 thiếu UV (từ 26/10), hàng sửa MAC (sau 28/10), Beam saw sau 4/11, 2 mã Ép nóng 7/11 chưa có giờ. W46: S32 chưa có hình Assembly, hàng sửa MAC, Veneer, Ép nóng.',
'4. A10 theo file xuất hệ thống (số công nhân do người nhập ghi); STS01 tính là Raw Cutting.',
'5. Thợ sơn lòng trong tủ không nhập. Thứ Bảy 14/11 có 10 HC tăng ca 15:30-20:00 sửa hàng (45 giờ công) chưa tính vì chưa rõ mã/bộ phận.',
'6. Không có NCR (O15), Others Losstime, Power cut, Kaizen (P24) trong dữ liệu hai tuần.']
for i,t_ in enumerate(notes,1):
    gn.cell(i,1,t_).font=Font(name=F,bold=(i==1),size=12 if i==1 else 11);gn.cell(i,1).alignment=Alignment(wrap_text=True,vertical='top')
gn.column_dimensions['A'].width=150
fix_charts(wb)
wb.save('Loss_time_White_Wood_W45_W46.xlsx')
print('file 1: dòng',len(rows),'A10',nA10)
