import openpyxl,datetime as dt,collections,warnings
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.utils import get_column_letter as L
warnings.simplefilter('ignore')
U='/root/.claude/uploads/64423226-a456-5e96-8225-6a5ffe123efb/'
rows=[r for r in openpyxl.load_workbook(U+'04ede00f-Loss_time_White_Wood_2026_-_New_1.xlsx',read_only=True)['Total loss time'].iter_rows(min_row=3,max_col=17,values_only=True) if r[1]]
assert max(r[1] for r in rows).date()==dt.date(2026,10,3)
# ---- fixed rows: latest weekday 2/10, must equal 28/9..2/10
def fx(d): return sorted([(r[6],r[7],r[9],r[10],r[11],r[12],r[13],r[15]) for r in rows if r[1].date()==d and r[7] in('P20','P26')],key=str)
base=fx(dt.date(2026,10,2)); assert len(base)==20
for d in [dt.date(2026,9,28),dt.date(2026,9,29),dt.date(2026,9,30),dt.date(2026,10,1)]: assert fx(d)==base,d
fixed=[r for r in rows if r[1].date()==dt.date(2026,10,2) and r[7] in('P20','P26')]
# ---- A2 cohorts as of 3/10
g=collections.defaultdict(dict)
for r in rows:
    if r[7]=='A2' and r[10] and '50%' in r[10]: g[(r[10],r[6])][r[1].date()]=g[(r[10],r[6])].get(r[1].date(),0)+(r[15] or 0)
co=[];stalled=[]
for (tag,ar),dd in g.items():
    ds=sorted(dd);n=len(ds)
    if n<12 and ds[-1]>=dt.date(2026,10,2): co.append((tag,ar,n,dd[ds[-1]],str(ds[-1])))
    elif n<12 and dt.date(2026,9,15)<=ds[-1]<dt.date(2026,10,2): stalled.append((tag.split('-')[-1],ar,n,str(ds[-1])))
co.sort(key=lambda x:(x[1],x[0]))
DAYS=[dt.datetime(2026,10,5),dt.datetime(2026,10,6),dt.datetime(2026,10,7),dt.datetime(2026,10,8),dt.datetime(2026,10,9),dt.datetime(2026,10,10),dt.datetime(2026,10,12),dt.datetime(2026,10,13),dt.datetime(2026,10,14),dt.datetime(2026,10,15),dt.datetime(2026,10,16),dt.datetime(2026,10,17),dt.datetime(2026,10,19),dt.datetime(2026,10,20),dt.datetime(2026,10,21),dt.datetime(2026,10,22),dt.datetime(2026,10,23),dt.datetime(2026,10,24),dt.datetime(2026,10,25),dt.datetime(2026,10,26),dt.datetime(2026,10,27),dt.datetime(2026,10,28),dt.datetime(2026,10,29)]
BASE=dt.datetime(2026,10,3)
def wd(a,b):
    n=0;x=a
    while x<b:
        x+=dt.timedelta(days=1)
        if x.weekday()!=6: n+=1
    return n
def kd(a,b):
    return -1 if b<a else wd(a,b)
SAT_NO_OT={dt.datetime(2026,10,10),dt.datetime(2026,10,17),dt.datetime(2026,10,24),dt.datetime(2026,10,25)}
WEEKDAY_NO_OT=set()   # Thứ Hai-Thứ Sáu luôn tăng ca (người dùng xác nhận 27/10); chỉ Thứ Bảy, Chủ nhật không tăng ca
SUN_DAY=dt.datetime(2026,10,25)
SUN_ASSY={SUN_DAY}   # Chủ nhật chỉ Assembly làm, không tăng ca; A2 của Assembly vẫn tính ngày đó
HIRE={DAYS[0]:[('Assembly',6)],DAYS[1]:[('Assembly',7),('Raw Cutting Panel',4),('Machining',3)],DAYS[2]:[('Assembly',6),('Raw Cutting Panel',2)],DAYS[3]:[('Raw Cutting Panel',1),('Machining',3)],DAYS[4]:[('Assembly',6),('Raw Cutting Panel',2)],DAYS[5]:[],DAYS[6]:[('Assembly',2),('Raw Cutting Panel',1)],DAYS[7]:[('Assembly',2),('Raw Cutting Panel',4),('Machining',1)],DAYS[8]:[('Assembly',1)],DAYS[9]:[('Assembly',4)],DAYS[10]:[('Assembly',1)],DAYS[11]:[],DAYS[12]:[('Assembly',2),('Machining',1)],DAYS[13]:[('Assembly',4),('Machining',2)],DAYS[14]:[('Raw Cutting Panel',1),('Machining',2),('Assembly',2)],DAYS[15]:[('Raw Cutting Panel',3),('Machining',1),('Assembly',2)],DAYS[16]:[('Raw Cutting Panel',1)],DAYS[17]:[],DAYS[18]:[],DAYS[19]:[('Assembly',2)],DAYS[20]:[('Assembly',4),('Raw Cutting Panel',1)],DAYS[21]:[],DAYS[22]:[]}
names={'P20':'Dailly Cleaning','P26':'Recorder/ Supporter','P22':'HR Training New Workers','A2':'Training'}
T=lambda h,m:dt.time(h,m)
def day_rows(D):
    sun=D in SUN_ASSY
    out=[(r[6],r[7],r[9],r[10],r[11],r[12],((T(15,30) if D in SAT_NO_OT else T(16,0)) if ((D in SAT_NO_OT or D in WEEKDAY_NO_OT) and r[13]==T(20,0)) else r[13]),r[15],'f') for r in fixed if (not sun or r[6]=='Assembly')]
    for ar,hc in HIRE[D]: out.append((ar,'P22','HR/HSE','Đào tạo công nhân mới',None,T(8,0),T(16,0),hc,'p'))
    for tag,ar,n,hc,last in co:
        if sun and ar!='Assembly': continue
        k=n+wd(BASE,D) if not sun else n+wd(BASE,D-dt.timedelta(days=1))+1
        if ar=='Assembly' and D>SUN_DAY: k+=1
        if k<=12: out.append((ar,'A2','HR/HSE',tag,f'A2 ngày {k}/12 - nhóm {tag.split("-")[-1]}',T(7,30),T(12,0),hc,'a'))
    for hd in DAYS:
        kk=wd(hd,D) if not sun else wd(hd,D-dt.timedelta(days=1))+1
        asm_shift=1 if (D>SUN_DAY and hd<SUN_DAY) else 0
        if True:
            for ar,hc in HIRE[hd]:
                if sun and ar!='Assembly': continue
                k2=kk+(asm_shift if ar=='Assembly' else 0)
                if not 1<=k2<=12: continue
                out.append((ar,'A2','HR/HSE',f'HC mới- 50% hiệu suất- 2 tuần-{hd:%d/%m}',f'A2 ngày {k2}/12 - nhóm {hd:%d/%m}',T(7,30),T(12,0),hc,'a'))
    return out
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
red=PatternFill('solid',fgColor='C00000');hf=Font(name=F,bold=True,color='FFFFFF');ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
H=['Week','Date \n(Ngày)',"Item's\nMã hàng",'Code Máy Hư',"Q'ty",'Số NCR/ME','Khu vực phát sinh/bộ phận hỗ trợ','Code','Name code','Chart / bộ phận chịu trách nhiệm','Lí Do','Ghi chú','From','To','LossTime (Hr)','HC','Loss time labor (Hrs)']
wb=openpyxl.Workbook();wb.remove(wb.active)
last={}
for D in DAYS:
    ws=wb.create_sheet(f'Loss time {D:%d-%m}')
    ws['A1']=f'LOSS TIME WHITE WOOD - {D:%d/%m/%Y}';ws['A1'].font=Font(name=F,bold=True,size=20)
    for i,h in enumerate(H,1):
        c=ws.cell(2,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
    data=day_rows(D)
    for k,(ar,code,dep,lido,gc,fr,to,hc,kind) in enumerate(data):
        r=3+k
        row=[f'="W"&WEEKNUM(B{r},21)',D,None,None,None,None,ar,code,names[code],dep,lido,gc,fr,to,f'=(N{r}-M{r})*24',hc,f'=P{r}*O{r}']
        for c,v in enumerate(row,1):
            x=ws.cell(r,c,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=c in(9,11,12))
        ws.cell(r,2).number_format='d-mmm';ws.cell(r,13).number_format=ws.cell(r,14).number_format='h:mm;@'
        ws.cell(r,15).number_format=ws.cell(r,17).number_format='0.0'
        fill={'p':'FFF2CC','a':'DDEBF7'}.get(kind)
        if fill:
            for c in range(1,18): ws.cell(r,c).fill=PatternFill('solid',fgColor=fill)
    e=2+len(data);last[D]=e
    for col,w in zip('ABCDEFGHIJKLMNOPQ',[7,9,10,10,8,10,22,8,26,14,40,28,9,9,10,8,15]): ws.column_dimensions[col].width=w
    ws.row_dimensions[2].height=37;ws.freeze_panes='A3';ws.auto_filter.ref=f'A2:Q{e}'
    ws.cell(e+1,16,'Tổng').font=Font(name=F,bold=True);ws.cell(e+1,17,f'=SUM(Q3:Q{e})').font=Font(name=F,bold=True);ws.cell(e+1,17).number_format='0.0'
    ws.cell(e+3,1,'Vàng = nhân viên mới ngày đầu (P22). Xanh = nhân viên đang ghi A2 (cột Ghi chú cho biết ngày thứ mấy trong 12 ngày).').font=Font(name=F,italic=True)
# ---- summary
sm=wb.create_sheet('Tổng hợp')
areas=['RC - WOOD','RC - PANEL','Machining','Assembly','RC','Raw Cutting Panel','Chưa rõ khu vực']
sm['A1']='TỔNG HỢP LOSS TIME LABOR (GIỜ CÔNG)';sm['A1'].font=Font(name=F,bold=True,size=14)
r0=3
for D in DAYS:
    sm.cell(r0,1,f'Ngày {D:%d/%m/%Y}').font=Font(name=F,bold=True)
    for i,h in enumerate(['Code','Tên']+areas+['Tổng'],1):
        c=sm.cell(r0+1,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
    sh=f"'Loss time {D:%d-%m}'";e=last[D]
    for k,code in enumerate(['P20','P26','P22','A2']):
        r=r0+2+k;sm.cell(r,1,code);sm.cell(r,2,names[code])
        for j in range(len(areas)):
            col=L(3+j)
            sm.cell(r,3+j,f'=SUMIFS({sh}!$Q$3:$Q${e},{sh}!$H$3:$H${e},$A{r},{sh}!$G$3:$G${e},{col}${r0+1})')
        sm.cell(r,3+len(areas),f'=SUM(C{r}:{L(2+len(areas))}{r})')
    r=r0+6;sm.cell(r,1,'Tổng')
    for j in range(len(areas)+1): col=L(3+j);sm.cell(r,3+j,f'=SUM({col}{r0+2}:{col}{r0+5})')
    for rr in range(r0+2,r0+7):
        for c in range(1,4+len(areas)):
            x=sm.cell(rr,c);x.font=Font(name=F,bold=(rr==r0+6 or c==3+len(areas)));x.border=bd
            if c>2:x.number_format='#,##0.0;-#,##0.0;-';x.alignment=Alignment(horizontal='center')
    r0+=9
sm.cell(r0-1,1,'Chưa gồm các mã lỗi máy móc và các mã khác theo yêu cầu.').font=Font(name=F,italic=True)
sm.column_dimensions['A'].width=8;sm.column_dimensions['B'].width=26
for j in range(len(areas)+1): sm.column_dimensions[L(3+j)].width=14
# ---- staff sheet
st=wb.create_sheet('Nhân viên mới - A2')
st['A1']='NHÂN VIÊN MỚI: P22 VÀ A2 THEO TỪNG NGÀY';st['A1'].font=Font(name=F,bold=True,size=14)
h2=['Nhóm (ngày vào làm)','Khu vực','HC']+[f'Mã ngày {D:%d/%m}' for D in DAYS]+['Ngày cuối A2 (không tính CN)']
for i,h in enumerate(h2,1):
    c=st.cell(3,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
st.row_dimensions[3].height=32
lines=[]
def kH(hd,ar,D):
    # số ngày A2 của nhóm vào làm hd, khu vực ar, tại ngày D (Chủ nhật 25/10 chỉ Assembly làm và được tính)
    if D in SUN_ASSY: return -1 if ar!='Assembly' else (wd(hd,D-dt.timedelta(days=1))+1 if hd<D else -1)
    k=kd(hd,D)
    if k>=0 and ar=='Assembly' and D>SUN_DAY and hd<SUN_DAY: k+=1
    return k
def kC(n,ar,D):
    if D in SUN_ASSY: return -1 if ar!='Assembly' else n+wd(BASE,D-dt.timedelta(days=1))+1
    k=n+wd(BASE,D)
    if ar=='Assembly' and D>SUN_DAY: k+=1
    return k
def endf(hd,ar):
    nd=12
    if ar=='Assembly' and hd<SUN_DAY:
        x=hd;c=0
        while c<12:
            x+=dt.timedelta(days=1)
            if x.weekday()!=6: c+=1
        if x>=SUN_DAY: nd=11
    return f'=WORKDAY.INTL(DATE(2026,{hd.month},{hd.day}),{nd},11)'
def code(k):
    return '-' if k<0 else 'P22' if k==0 else (f'A2 ngày {k}' if k<=12 else 'Đã xong')
for hd in reversed(DAYS):
    for ar,hc in HIRE[hd]:
        lines.append((f'{hd:%d/%m}',ar,hc,[code(kH(hd,ar,D)) for D in DAYS],endf(hd,ar)))
for tag,ar,n,hc,lastd in co:
    lines.append((tag.split('-')[-1],ar,hc,[code(kC(n,ar,D)) for D in DAYS],f'=WORKDAY.INTL(DATE(2026,10,5),{12-(n+1)-(1 if ar=="Assembly" else 0)},11)'))
nc=len(DAYS)
for k,(a_,b_,c_,cs,f_) in enumerate(lines):
    r=4+k
    for i,v in enumerate([a_,b_,c_]+cs+[f_],1):
        x=st.cell(r,i,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='center',vertical='center')
    st.cell(r,4+nc).number_format='d-mmm-yyyy'
    fill='FFF2CC' if 'P22' in cs else 'DDEBF7'
    for i in range(1,5+nc): st.cell(r,i).fill=PatternFill('solid',fgColor=fill)
n0=4+len(lines)+1
notes=['Quy ước: P22 ngày đầu (8:00-16:00); A2 các ngày sau (12 ngày làm việc, bỏ Chủ nhật), giờ A2 gần nhất trong file là 7:30-12:00.',
'Số ngày A2 đã ghi của các nhóm cũ được đếm theo số ngày thực tế có dòng A2 trong file Loss time (đến 03/10); HC lấy theo ngày ghi gần nhất.',
'Nhóm 18/9 Assembly: dòng A2 cuối là 02/10, không có 03/10. Nếu 03/10 bị bỏ sót thì số ngày A2 của nhóm này lớn hơn 1 ngày.',
'Nhóm 19/9 Assembly (13 ngày) và 19/9 Machining (12 ngày) đã đủ A2 trước 05/10 nên không ghi. Nếu quy định là 13 ngày thì Machining 19/9 (8 người) còn 1 ngày.',
'Người ghi "0" ở cột CT/TV vẫn tính vào HC (05/10: 1 người; 06/10: 3 người; 13/10: 3 người Raw Cutting Panel). Danh sách 07/10 không có cột CT/TV nên tính đủ 8 người; 08/10 có 4 người, đều NCU.',
'Ngày 09/10 có 8 nhân viên mới (Assembly 6, Raw Cutting Panel 2; lấy theo P22 đã nhập trên hệ thống), A2 bắt đầu từ 10/10. Ngày 10/10 chưa thấy có nhân viên mới (chờ xác nhận).',
'Chủ nhật 11/10 không làm nên không tính vào số ngày A2. Ngày 12/10 có 3 nhân viên mới (Assembly 2, Raw Cutting Panel 1), A2 bắt đầu từ 13/10. Ngày 13/10 có 7 nhân viên mới (Assembly 2, Raw Cutting Panel 4, Machining 1), ngày 14/10 có 1 (Assembly), ngày 15/10 có 4 (Assembly), ngày 16/10 có 1 nhân viên mới (Assembly, theo dòng P22 đã nhập trên hệ thống), ngày 17/10 không có nhân viên mới. Ngày 19/10 có 3 nhân viên mới (Assembly 2, Machining 1; người ghi "0" ở cột CT/TV vẫn tính vào HC), ngày 20/10 có 6 (Assembly 4, gồm 2 người dưới Mr.Dũng-Inlay; Machining 2). Thứ Bảy 17/10 không tăng ca (P26 và P20 Assembly 7:30-15:30). Ngày 21/10 có 5 nhân viên mới (Raw Cutting Panel 1, Machining 2, Assembly 2 gồm 1 người dưới Mr.Dũng-Inlay). Ngày 22/10 có 6 nhân viên mới (Raw Cutting Panel 3, Machining 1, Assembly 2; người ghi "0" ở cột CT/TV vẫn tính vào HC). Ngày 23/10 và 24/10 chưa có nhân viên mới (chờ báo). Thứ Bảy 24/10 mặc định không tăng ca (chờ xác nhận); Ngày 23/10 có 1 nhân viên mới Raw Cutting Panel (người ghi "0" ở cột CT/TV vẫn tính). Thứ Hai-Thứ Sáu luôn tăng ca (26/10 đến 29/10 theo mẫu ngày thường đến 20:00; chỉ Thứ Bảy, Chủ nhật không tăng ca). Ngày 26/10 có 2 nhân viên mới Assembly; 27/10 có 5 (Assembly 4 gồm 1 người dưới Mr.Dũng-Inlay, Raw Cutting Panel 1, người ghi "0" ở cột CT/TV vẫn tính); 28/10 và 29/10 chưa có nhân viên mới (chờ báo), từ 26/10 A2 của các nhóm Assembly cộng thêm 1 ngày vì Chủ nhật 25/10 đã tính. Chủ nhật 25/10 chỉ Assembly làm, không tăng ca: P20 và P26 của Assembly ghi 7:30-15:30; A2 của các nhóm Assembly vẫn ghi và tính là một ngày (theo tiền lệ 27/9). Từ 19/10 đến 23/10 ghi theo mẫu ngày thường có tăng ca đến 20:00 (cần xác nhận). Tài xế xe nâng (Machining 12/10, nhóm IDL-Forklift) không tính theo yêu cầu.',
'Thứ Bảy 10/10 không tăng ca: P26 và P20 Assembly ghi 7:30-15:30 (sớm 30 phút so với ca ngày thường 7:30-16:00); 4 dòng P20 ngắn ghi 7:00-7:30 theo mặc định, chờ xác nhận.']
for i,tx in enumerate(notes): st.cell(n0+i,1,tx).font=Font(name=F,italic=True)
for col,w in zip('ABCDEFGHIJKLMNOP',[22,20,6,16,16,16,16,16,16,16,16,16,16,16,16,26]): st.column_dimensions[col].width=w

# ---- A2 schedule sheet (full list of A2 dates per new-hire cohort)
sc=wb.create_sheet('Lịch A2 nhân viên mới')
sc['A1']='LỊCH A2 CỦA TỪNG NHÓM NHÂN VIÊN MỚI (12 ngày làm việc, bỏ Chủ nhật)';sc['A1'].font=Font(name=F,bold=True,size=14)
hd=['Nhóm (ngày vào làm)','Khu vực','HC','Ngày P22']+[f'A2 ngày {i}' for i in range(1,13)]
for i,h in enumerate(hd,1):
    c=sc.cell(3,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
sc.row_dimensions[3].height=32
r=4
for hdate in DAYS:
    for ar,hc in HIRE[hdate]:
        sc.cell(r,1,f'{hdate:%d/%m}');sc.cell(r,2,ar);sc.cell(r,3,hc)
        sc.cell(r,4,f'=DATE(2026,{hdate.month},{hdate.day})')
        sc.cell(r,5,f'=WORKDAY.INTL(D{r},1,11)')
        for j in range(2,13):
            sc.cell(r,4+j,f'=WORKDAY.INTL({L(3+j)}{r},1,11)')
        for cc in range(1,17):
            x=sc.cell(r,cc);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='center',vertical='center')
            if cc>=4: x.number_format='dd/mm'
        r+=1
sc.cell(r+1,1,'Mỗi ngày trong lịch ghi A2 7:30-12:00, Notes "HC mới- 50% hiệu suất- 2 tuần-dd/mm" (dd/mm = ngày vào làm). Khi có nhóm mới, thêm một dòng.').font=Font(name=F,italic=True)
sc.column_dimensions['A'].width=22;sc.column_dimensions['B'].width=20;sc.column_dimensions['C'].width=6
for cc in range(4,17): sc.column_dimensions[L(cc)].width=11
wb.save('Loss_time_05-10_den_29-10.xlsx')
print('co',len(co),'stalled',stalled)
for D in DAYS: print(D.date(),len(day_rows(D)))
