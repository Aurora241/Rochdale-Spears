"""Đối chiếu file xuất hệ thống W45 (2-8/11) với bảng chuẩn (P20, P26, P22, A2, S32) -> danh sách sửa/thêm/xóa theo bộ phận.
Chạy trong thư mục outputs: python ../scripts/build_fix_list_w45.py <file xuất>"""
import openpyxl,sys,collections,datetime as dt,warnings
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
warnings.simplefilter('ignore')
U=sys.argv[1]
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
rows=list(openpyxl.load_workbook(U,data_only=True).active.iter_rows(values_only=True));H=rows[13]
def tm(x):
    if isinstance(x,str):h,m=x.split(':')[:2];return dt.time(int(h),int(m))
    return x
sysr=[]
for r in rows[14:]:
    if not r[0]:continue
    d=dict(zip(H,r));day=dt.datetime.strptime(d['Loss Date'],'%d %b %Y')
    if d['LossTime Code'] in('A10',):continue
    sysr.append(dict(day=day,code=d['LossTime Code'],notes=(d['Notes'] or '').strip(),proc=d['Process Name'],dept=(d['Dept In Charge'] or '').strip(),fr=tm(d['Loss From Time']),to=tm(d['Loss To Time']),hc=d['Quantity Workers']))
PR={'RC - WOOD':'Raw Cutting','RC - PANEL':'Raw Cutting','RC':'Raw Cutting','Raw Cutting Panel':'Raw Cutting','RC WOOD':'Raw Cutting','Machining':'Machining','Assembly':'Assembly','INLAY':'INLAY'}
DP={'RC - PANEL':'RAW CUTTING','RC WOOD':'RAW CUTTING','RC - WOOD':'RAW CUTTING','Machining':'MACHINING','Assembly':'ASSEMBLY','HR/HSE':'HR/ HSE'}
mine=[]
wb=openpyxl.load_workbook('Loss_time_05-10_den_14-11.xlsx',data_only=True)
for d in ('02-11','03-11','04-11','05-11','06-11','07-11','08-11'):
    day=dt.datetime(2026,int(d[3:]),int(d[:2]))
    for r in wb[f'Loss time {d}'].iter_rows(min_row=3,values_only=True):
        if r[7] in('P20','P26','P22','A2') and r[6]:
            mine.append(dict(day=day,code=r[7],notes=(r[10] or '').strip(),proc=PR[r[6]],dept=DP.get(r[9],r[9]),fr=r[12],to=r[13],hc=r[15]))
s32=openpyxl.load_workbook('Loss_time_S32_05-10_den_08-11.xlsx',data_only=True)['S32 5-10 đến 08-11']
SD={'FINISHING':'FINISHING','TECHNICAL FINISHING':'TECHNICAL FINISHING','MACHINING':'MACHINING','SUBCON':'SUBCON','ME':'ME','OTHERS (INVESTIGATE)':'OTHERS (INVESTIGATE)','(cần xác nhận)':'(cần xác nhận)'}
for r in s32.iter_rows(min_row=3,values_only=True):
    if r[1] and hasattr(r[1],'strftime') and r[1]>=dt.datetime(2026,11,2):
        mine.append(dict(day=r[1],code='S32',notes=(r[10] or '').strip(),proc=r[6],dept=r[9],fr=r[12],to=r[13],hc=r[15]))
mine.append(dict(day=dt.datetime(2026,11,2),code='S32',notes='Cào cước hỗ trợ ME',proc='Machining',dept='ME',fr=dt.time(12,28),to=dt.time(13,48),hc=1))  # người dùng báo 2/11 (80 phút), chưa có trong bảng S32
# so khớp: key đầy đủ; phần dư ghép theo (ngày, code, process, notes) -> Sửa
full=lambda x:(x['day'],x['code'],x['proc'],x['notes'][:45],x['dept'],x['fr'],x['to'],x['hc'])
loose=lambda x:(x['day'],x['code'],x['proc'],x['notes'][:45])
cs=collections.Counter(full(x) for x in sysr);cm=collections.Counter(full(x) for x in mine)
addm=list((cm-cs).elements());delm=list((cs-cm).elements())
res=[]
byl={}
for k in delm: byl.setdefault(k[:4],[]).append(k)
for k in addm:
    cand=byl.get(k[:4],[])
    if cand:
        o=cand.pop(0);diff=[n for n,a,b in (('Dept',k[4],o[4]),('From',k[5],o[5]),('To',k[6],o[6]),('Worker Qty',k[7],o[7])) if a!=b]
        res.append(('Sửa',k,'Hệ thống: '+', '.join(f'{n} {b}' for n,a,b in (('Dept',k[4],o[4]),('From',k[5],o[5]),('To',k[6],o[6]),('Worker Qty',k[7],o[7])) if a!=b)+' -> sửa thành bảng chuẩn'))
    else: res.append(('Thêm',k,'Chưa có trên hệ thống'))
for lst in byl.values():
    for k in lst: res.append(('Xóa',k,'Có trên hệ thống nhưng không có trong bảng chuẩn (nhập thừa/trùng hoặc sai nhóm)'))
res.sort(key=lambda x:(x[1][2],x[1][0],{'S32':4,'P20':0,'P26':1,'P22':2,'A2':3}[x[1][1]],x[0]))
wbo=openpyxl.Workbook();wbo.remove(wbo.active)
red=PatternFill('solid',fgColor='C00000');gry=PatternFill('solid',fgColor='D9D9D9')
col={'Thêm':'E2EFDA','Sửa':'FFF2CC','Xóa':'FCE4D6'}
for area in ('Assembly','Machining','Raw Cutting'):
    ws=wbo.create_sheet(area);ws['A1']=f'DANH SÁCH SỬA / THÊM / XÓA - {area.upper()} - TUẦN W45 (2/11-8/11)';ws['A1'].font=Font(name=F,bold=True,size=14)
    ws['A2']='Xếp theo ngày, trong ngày theo mã P20, P26, P22, A2, S32. Đối chiếu file xuất hệ thống 9/11 với bảng chuẩn. A10 không đối chiếu ở đây (đã có bảng riêng).';ws['A2'].font=Font(name=F,italic=True)
    hd=['Hành động','Ngày (Loss Date)','LossTime Code','Notes','Process','Dept In Charge','From','To','Worker Qty','Lý do / ghi chú','Đã xử lý']
    for i,h in enumerate(hd,1):
        c=ws.cell(4,i,h);c.fill=red if i<=9 else gry;c.font=Font(name=F,bold=True,color='FFFFFF' if i<=9 else '000000');c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True);c.border=bd
    r=5
    for act,k,why in res:
        if k[2]!=area:continue
        for i,v in enumerate([act,k[0],k[1],k[3],k[2],k[4],k[5],k[6],k[7],why,None],1):
            c=ws.cell(r,i,v);c.font=Font(name=F);c.border=bd;c.alignment=Alignment(horizontal='left' if i in(4,10) else 'center',vertical='center',wrap_text=i in(4,10))
        ws.cell(r,1).fill=PatternFill('solid',fgColor=col[act]);ws.cell(r,2).number_format='dd/mm/yyyy';ws.cell(r,7).number_format=ws.cell(r,8).number_format='hh:mm'
        r+=1
    for c_,w in zip('ABCDEFGHIJK',[10,14,10,46,13,22,8,8,10,58,9]):ws.column_dimensions[c_].width=w
    ws.freeze_panes='A5'
wbo.save('Danh_sach_sua_them_W45.xlsx')
cnt=collections.Counter((x[0],x[1][2]) for x in res);print(cnt)
