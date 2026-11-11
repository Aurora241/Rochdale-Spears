"""Thêm nhóm nhân viên mới vào sheet 'Nhân viên mới - A2' và 'Lịch A2 nhân viên mới' (trước khi chạy patch_w46.py cho ngày vào làm).
python add_group_nv.py <file> dd-mm Assembly:1,Machining:1  (chạy trong outputs)"""
import openpyxl,sys,re,copy,datetime as dt,warnings
warnings.simplefilter('ignore')
f,dm,spec=sys.argv[1:4];day=dt.datetime(2026,int(dm[3:]),int(dm[:2]))
SUN={dt.datetime(2026,10,25),dt.datetime(2026,11,1),dt.datetime(2026,11,8)}
def endday(g,area):
    x=g;n=0
    while n<12:
        x+=dt.timedelta(days=1)
        if x.weekday()!=6 or (area=='Assembly' and x in SUN): n+=1
    return x
wb=openpyxl.load_workbook(f);nv=wb['Nhân viên mới - A2']
items=[(a,int(h)) for a,h in (x.split(':') for x in spec.split(','))]
hd=[nv.cell(3,c).value for c in range(1,nv.max_column+1)]
nv.insert_rows(4,len(items))
for i,(ar,hc) in enumerate(items):
    r=4+i
    for c in range(1,nv.max_column+1): nv.cell(r,c)._style=copy.copy(nv.cell(4+len(items),c)._style)
    nv.cell(r,1).value=day.strftime('%d/%m');nv.cell(r,2).value=ar;nv.cell(r,3).value=hc
    for c,h in enumerate(hd,1):
        if h and str(h).startswith('Mã ngày'): nv.cell(r,c).value='-'
    e=endday(day,ar);nv.cell(r,nv.max_column).value=f'=DATE({e.year},{e.month},{e.day})'
lc=wb['Lịch A2 nhân viên mới']
last=next(r for r in range(4,lc.max_row+1) if not lc.cell(r,1).value)-1
lc.insert_rows(last+1,len(items))
def fix(v,row): return re.sub(r'([A-Z]+)\d+',lambda m:m.group(1)+str(row),v) if isinstance(v,str) and v.startswith('=') else v
for i,(ar,hc) in enumerate(items):
    r=last+1+i
    for c in range(1,lc.max_column+1): lc.cell(r,c)._style=copy.copy(lc.cell(last,c)._style);lc.cell(r,c).value=fix(lc.cell(last,c).value,r)
    lc.cell(r,1).value=day.strftime('%d/%m');lc.cell(r,2).value=ar;lc.cell(r,3).value=hc;lc.cell(r,4).value=f'=DATE({day.year},{day.month},{day.day})'
wb.save(f);print('đã thêm nhóm',dm,items)
