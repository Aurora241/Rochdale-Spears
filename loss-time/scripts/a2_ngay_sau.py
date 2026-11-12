"""Thêm sheet 'A2 NGÀY SAU' vào file nhập liệu theo bộ phận: mọi dòng A2 của các ngày CHƯA tới (đến hết 12 ngày của từng nhóm) để người dùng không sót.
python a2_ngay_sau.py <Loss_time xlsx> <file theo bộ phận> dd-mm_ngày_cuối_đã_lập  (chạy trong outputs)"""
import openpyxl,sys,re,datetime as dt,warnings
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
warnings.simplefilter('ignore')
lt,out,last=sys.argv[1:4];lastd=dt.datetime(2026,int(last[3:]),int(last[:2]))
SUN={dt.datetime(2026,10,25),dt.datetime(2026,11,1),dt.datetime(2026,11,8)}
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
red=PatternFill('solid',fgColor='C00000');gry=PatternFill('solid',fgColor='D9D9D9')
nv=openpyxl.load_workbook(lt,data_only=True)['Nhân viên mới - A2']
G=[]
for r in range(4,nv.max_row+1):
    g=nv.cell(r,1).value
    if g and re.match(r'\d\d/\d\d$',str(g)):
        dd,mm=map(int,g.split('/'));G.append((dt.datetime(2026,mm,dd),nv.cell(r,2).value,nv.cell(r,3).value,g))
def k(g,a,d):
    n=0;x=g
    while x<d:
        x+=dt.timedelta(days=1)
        if x.weekday()!=6 or (a=='Assembly' and x in SUN): n+=1
    return n
PR=lambda a:'Assembly' if a=='Assembly' else ('Machining' if a=='Machining' else 'Raw Cutting')
rows=[]
for g,a,hc,txt in G:
    for off in range(1,30):
        d=lastd+dt.timedelta(days=off)
        if d.weekday()==6: continue
        kk=k(g,a,d)
        if 1<=kk<=12: rows.append((d,PR(a),hc,kk,txt))
rows.sort(key=lambda x:(x[0],['Assembly','Machining','Raw Cutting'].index(x[1]),x[4]))
wb=openpyxl.load_workbook(out)
if 'A2 NGÀY SAU' in wb.sheetnames: del wb['A2 NGÀY SAU']
ws=wb.create_sheet('A2 NGÀY SAU',1)
ws['A1']='A2 CỦA CÁC NGÀY SAU (CHƯA TỚI): ĐỪNG ĐỂ SÓT'
ws['A1'].font=Font(name=F,bold=True,size=14)
ws['A2']=f'Mỗi nhóm nhân viên mới phải có A2 đủ 12 ngày làm việc sau ngày P22. Bảng này liệt kê các dòng A2 từ {(lastd+dt.timedelta(days=1)):%d/%m} trở đi cho mọi nhóm đang chạy (Chủ nhật không tính; nếu Assembly làm Chủ nhật thì số ngày của nhóm Assembly rút ngắn 1). Nhóm mới có người vào làm thì chạy lại để bổ sung.'
ws['A2'].font=Font(name=F,italic=True,color='C00000');ws['A2'].alignment=Alignment(wrap_text=True,vertical='top');ws.merge_cells('A2:K2');ws.row_dimensions[2].height=44
H=['Ngày (Loss Date)','LossTime Code','Notes','Process','Dept In Charge','From','To','Worker Qty','Nhóm','Ngày A2','Đã nhập']
for i,h in enumerate(H,1):
    c=ws.cell(4,i,h);c.fill=red if i<=8 else gry;c.font=Font(name=F,bold=True,color='FFFFFF' if i<=8 else '000000');c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True);c.border=bd
r=5;prev=None
for d,pr,hc,kk,txt in rows:
    if prev is not None and d!=prev: r+=0
    prev=d
    vals=[d,'A2',f'HC mới- 50% hiệu suất- 2 tuần-{txt}',pr,'HR/ HSE',dt.time(7,30),dt.time(12,0),hc,f'nhóm {txt}',f'A2 ngày {kk}/12',None]
    for i,v in enumerate(vals,1):
        c=ws.cell(r,i,v);c.font=Font(name=F);c.border=bd;c.alignment=Alignment(horizontal='left' if i==3 else 'center',vertical='center',wrap_text=i==3)
    ws.cell(r,1).number_format='dd/mm/yyyy';ws.cell(r,6).number_format=ws.cell(r,7).number_format='hh:mm'
    if kk==12:
        for i in range(1,12): ws.cell(r,i).fill=PatternFill('solid',fgColor='FCE4D6')
    r+=1
ws.cell(r+1,1,'Nền cam = ngày A2 cuối (12/12) của nhóm.').font=Font(name=F,italic=True)
for col,w in zip('ABCDEFGHIJK',[16,12,44,14,16,8,8,10,14,12,9]): ws.column_dimensions[col].width=w
ws.freeze_panes='A5'
wb.save(out);print('A2 ngày sau:',len(rows),'dòng, từ',rows[0][0].strftime('%d/%m') if rows else '-','đến',rows[-1][0].strftime('%d/%m') if rows else '-')
