"""Đối chiếu file xuất hệ thống một tuần với bảng chuẩn -> danh sách sửa/thêm/xóa theo bộ phận + thống kê sai sót.
python build_fix_list_tuan.py <file xuất> dd-mm dd-mm <file ra> <Loss_time xlsx> <S32 xlsx> <A10 nhập xlsx>   (chạy trong outputs)"""
import openpyxl,sys,collections,datetime as dt,warnings
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
warnings.simplefilter('ignore')
U,a0,a1,OUT,LT,S32F,A10F=sys.argv[1:8]
d0=dt.datetime(2026,int(a0[3:]),int(a0[:2]));d1=dt.datetime(2026,int(a1[3:]),int(a1[:2]))
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
rows=list(openpyxl.load_workbook(U,data_only=True).active.iter_rows(values_only=True));H=rows[13]
def tm(x):
    if isinstance(x,str):h,m=x.split(':')[:2];return dt.time(int(h),int(m))
    return x
sysr=[];sysa10=[]
for r in rows[14:]:
    if not r[0]:continue
    d=dict(zip(H,r));day=dt.datetime.strptime(d['Loss Date'],'%d %b %Y')
    x=dict(day=day,code=d['LossTime Code'],notes=(d['Notes'] or '').strip(),proc=d['Process Name'],dept=(d['Dept In Charge'] or '').strip(),fr=tm(d['Loss From Time']),to=tm(d['Loss To Time']),hc=d['Quantity Workers'])
    (sysa10 if x['code']=='A10' else sysr).append(x)
PR={'RC - WOOD':'Raw Cutting','RC - PANEL':'Raw Cutting','RC':'Raw Cutting','Raw Cutting Panel':'Raw Cutting','RC WOOD':'Raw Cutting','Machining':'Machining','Assembly':'Assembly','INLAY':'INLAY'}
DP={'RC - PANEL':'RAW CUTTING','RC WOOD':'RAW CUTTING','RC - WOOD':'RAW CUTTING','Machining':'MACHINING','Assembly':'ASSEMBLY','HR/HSE':'HR/ HSE'}
mine=[]
wb=openpyxl.load_workbook(LT,data_only=True)
day=d0
while day<=d1:
    nm=day.strftime('%d-%m')
    if f'Loss time {nm}' in wb.sheetnames:
        for r in wb[f'Loss time {nm}'].iter_rows(min_row=3,values_only=True):
            if r[7] in('P20','P26','P22','A2') and r[6]:
                mine.append(dict(day=day,code=r[7],notes=(r[10] or '').strip(),proc=PR[r[6]],dept=DP.get(r[9],r[9]),fr=r[12],to=r[13],hc=r[15]))
    day+=dt.timedelta(days=1)
s32=openpyxl.load_workbook(S32F,data_only=True)
s32=next(w for w in s32 if w.title.startswith('S32 5-10'))
for r in s32.iter_rows(min_row=3,values_only=True):
    if r[1] and hasattr(r[1],'strftime') and d0<=r[1]<=d1:
        mine.append(dict(day=r[1],code='S32',notes=(r[10] or '').strip(),proc=r[6],dept=r[9],fr=r[12],to=r[13],hc=r[15]))
full=lambda x:(x['day'],x['code'],x['proc'],x['notes'][:45],x['dept'],x['fr'],x['to'],x['hc'])
loose=lambda x:(x['day'],x['code'],x['proc'],x['notes'][:45])
cs=collections.Counter(full(x) for x in sysr);cm=collections.Counter(full(x) for x in mine)
addm=list((cm-cs).elements());delm=list((cs-cm).elements())
res=[];errs=collections.Counter();byl=collections.defaultdict(list)
for k in delm: byl[k[:4]].append(k)
for k in addm:
    c=byl.get(k[:4],[])
    if c:
        o=c.pop(0);fl=[(n,a,b) for n,a,b in (('Dept',k[4],o[4]),('From',k[5],o[5]),('To',k[6],o[6]),('Worker Qty',k[7],o[7])) if a!=b]
        for n,a,b in fl: errs[('Sai '+n,k[2])]+=1
        res.append(('Sửa',k,'Hệ thống: '+', '.join(f'{n} {b}' for n,a,b in fl)+' -> sửa thành bảng chuẩn'))
    else:
        errs[('Thiếu dòng',k[2])]+=1;res.append(('Thêm',k,'Chưa có trên hệ thống'))
for lst in byl.values():
    for k in lst: errs[('Thừa/sai nhóm',k[2])]+=1;res.append(('Xóa',k,'Có trên hệ thống nhưng không có trong bảng chuẩn (nhập thừa/trùng hoặc sai nhóm)'))
# ghép Xóa + Thêm cùng nội dung nhưng khác mã -> "Sai mã"
dels=[x for x in res if x[0]=='Xóa' and x[1][1]!='A10'];adds=[x for x in res if x[0]=='Thêm' and x[1][1]!='A10']
merged=[]
for d in list(dels):
    for a in list(adds):
        kd,ka=d[1],a[1]
        if kd[0]==ka[0] and kd[2]==ka[2] and kd[3]==ka[3] and kd[5]==ka[5] and kd[6]==ka[6] and kd[7]==ka[7] and kd[1]!=ka[1]:
            res.remove(d);res.remove(a);adds.remove(a)
            errs[('Thừa/sai nhóm',kd[2])]-=1;errs[('Thiếu dòng',ka[2])]-=1;errs[('Sai mã',ka[2])]+=1
            ex=[]
            if kd[4]!=ka[4]: ex.append(f'Dept {kd[4]} -> {ka[4]}')
            res.append(('Sửa',ka,f'Sai mã: hệ thống ghi {kd[1]}, đúng là {ka[1]}'+(' ('+', '.join(ex)+')' if ex else '')));break
for k in [k for k,v in errs.items() if v<=0]: del errs[k]
# A10: so theo (ngày, process, mã máy, from, to)
a10=openpyxl.load_workbook(A10F,data_only=True);std10=[]
for n in ('Machining','Raw Cutting'):
    if n not in a10.sheetnames: continue
    for r in a10[n].iter_rows(min_row=5,values_only=True):
        if isinstance(r[0],dt.datetime) and d0<=r[0]<=d1: std10.append((r[0],r[3],str(r[2]).split('-')[0],r[5],r[6],r[2]))
def k10(x):return (x['day'],x['proc'],x['notes'].split('-')[0],x['fr'],x['to'])
sk=collections.Counter(k10(x) for x in sysa10);mk=collections.Counter((a,b,c,d,e) for a,b,c,d,e,_ in std10)
for k in (mk-sk).elements():
    n=next(s[5] for s in std10 if (s[0],s[1],s[2],s[3],s[4])==k);errs[('A10 thiếu dòng',k[1])]+=1
    res.append(('Thêm',(k[0],'A10',k[1],n,'MAINTENANCE',k[3],k[4],None),'A10 chưa có trên hệ thống (số công nhân cần tổ trưởng xác nhận)'))
for k in (sk-mk).elements():
    x=next(y for y in sysa10 if k10(y)==k);errs[('A10 thừa/sai giờ/sai khu vực',k[1])]+=1
    res.append(('Xóa',(x['day'],'A10',x['proc'],x['notes'],x['dept'],x['fr'],x['to'],x['hc']),'A10 trên hệ thống không khớp bảng bảo trì (sai giờ, sai khu vực hoặc thừa)'))
wbo=openpyxl.Workbook();wbo.remove(wbo.active)
red=PatternFill('solid',fgColor='C00000');gry=PatternFill('solid',fgColor='D9D9D9');col={'Thêm':'E2EFDA','Sửa':'FFF2CC','Xóa':'FCE4D6'}
ORD={'A10':0,'P20':1,'P26':2,'P22':3,'A2':4,'S32':5}
res.sort(key=lambda x:(x[1][2],x[1][0],ORD.get(x[1][1],9),x[0]))
# thống kê
ws=wbo.create_sheet('Thống kê')
tot_std=collections.Counter(x['proc'] for x in mine)+collections.Counter(a[3] for a in std10)
ws['A1']=f'THỐNG KÊ SAI SÓT {d0:%d/%m} - {d1:%d/%m}/2026 (file xuất hệ thống đối chiếu bảng chuẩn)';ws['A1'].font=Font(name=F,bold=True,size=14)
ws['A2']='Mỗi dòng sai tính một lần theo loại lỗi; một dòng có thể sai nhiều loại (ví dụ vừa sai Dept vừa sai To). Tỉ lệ = số việc cần làm / số dòng chuẩn của bộ phận.';ws['A2'].font=Font(name=F,italic=True)
ws['A4']='A. Theo bộ phận';ws['A4'].font=Font(name=F,bold=True)
hd=['Bộ phận','Số dòng chuẩn','Thêm','Sửa','Xóa','Tổng việc','Tỉ lệ']
for i,h in enumerate(hd,1):
    c=ws.cell(5,i,h);c.fill=red;c.font=Font(name=F,bold=True,color='FFFFFF');c.alignment=Alignment(horizontal='center');c.border=bd
cnt=collections.Counter((a,k[2]) for a,k,_ in res);r=6
for area in ('Assembly','Machining','Raw Cutting'):
    v=[area,tot_std.get(area,0),cnt[('Thêm',area)],cnt[('Sửa',area)],cnt[('Xóa',area)]]
    for i,x in enumerate(v,1):c=ws.cell(r,i,x);c.border=bd;c.font=Font(name=F)
    ws.cell(r,6,f'=C{r}+D{r}+E{r}').border=bd;ws.cell(r,7,f'=IF(B{r}=0,0,F{r}/B{r})').number_format='0.0%';ws.cell(r,7).border=bd;r+=1
ws.cell(r,1,'Tổng').font=Font(name=F,bold=True)
for i in range(2,7):ws.cell(r,i,f'=SUM({chr(64+i)}6:{chr(64+i)}{r-1})').font=Font(name=F,bold=True)
ws.cell(r,7,f'=IF(B{r}=0,0,F{r}/B{r})').number_format='0.0%'
r+=2;ws.cell(r,1,'B. Theo loại lỗi').font=Font(name=F,bold=True);r+=1
for i,h in enumerate(['Loại lỗi','Assembly','Machining','Raw Cutting','Tổng'],1):
    c=ws.cell(r,i,h);c.fill=red;c.font=Font(name=F,bold=True,color='FFFFFF');c.alignment=Alignment(horizontal='center');c.border=bd
types=sorted({k[0] for k in errs});r+=1
for tp in types:
    ws.cell(r,1,tp).border=bd
    for i,a in enumerate(('Assembly','Machining','Raw Cutting'),2):c=ws.cell(r,i,errs.get((tp,a),0));c.border=bd
    ws.cell(r,5,f'=SUM(B{r}:D{r})').border=bd;r+=1
ws.column_dimensions['A'].width=34
for c_ in 'BCDEFG':ws.column_dimensions[c_].width=14
for area in ('Assembly','Machining','Raw Cutting'):
    ws=wbo.create_sheet(area);ws['A1']=f'DANH SÁCH SỬA / THÊM / XÓA - {area.upper()} - {d0:%d/%m} đến {d1:%d/%m}';ws['A1'].font=Font(name=F,bold=True,size=14)
    ws['A2']='Xếp theo ngày, trong ngày theo mã A10, P20, P26, P22, A2, S32. Đối chiếu file xuất hệ thống với bảng chuẩn.';ws['A2'].font=Font(name=F,italic=True)
    hd=['Hành động','Ngày (Loss Date)','LossTime Code','Notes','Process','Dept In Charge','From','To','Worker Qty','Lý do / ghi chú','Đã xử lý']
    for i,h in enumerate(hd,1):
        c=ws.cell(4,i,h);c.fill=red if i<=9 else gry;c.font=Font(name=F,bold=True,color='FFFFFF' if i<=9 else '000000');c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True);c.border=bd
    r=5
    for act,k,why in res:
        if k[2]!=area:continue
        for i,v in enumerate([act,k[0],k[1],k[3],k[2],k[4],k[5],k[6],k[7],why,None],1):
            c=ws.cell(r,i,v);c.font=Font(name=F);c.border=bd;c.alignment=Alignment(horizontal='left' if i in(4,10) else 'center',vertical='center',wrap_text=i in(4,10))
        ws.cell(r,1).fill=PatternFill('solid',fgColor=col[act]);ws.cell(r,2).number_format='dd/mm/yyyy';ws.cell(r,7).number_format=ws.cell(r,8).number_format='hh:mm';r+=1
    for c_,w in zip('ABCDEFGHIJK',[10,14,10,46,13,22,8,8,10,60,9]):ws.column_dimensions[c_].width=w
    ws.freeze_panes='A5'
wbo.save(OUT);print(collections.Counter((a,k[2]) for a,k,_ in res),len(mine),len(std10),'dòng chuẩn')
