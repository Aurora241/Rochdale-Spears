"""Thêm nhân viên mới 5/11 (Machining 2, Assembly 3) vào Loss_time_05-10_den_06-11.xlsx:
P22 ngày 5/11, A2 từ 6/11. Chạy trong thư mục outputs."""
import openpyxl,copy,re,warnings,datetime as dt
warnings.simplefilter('ignore')
F='Loss_time_05-10_den_06-11.xlsx'
wb=openpyxl.load_workbook(F)
NEW=[('Machining',2),('Assembly',3)]
def cp(ws,src,dst,vals):
    for c in range(1,ws.max_column+1):
        s=ws.cell(src,c);d=ws.cell(dst,c);d._style=copy.copy(s._style)
    for c,v in vals.items(): ws.cell(dst,c).value=v
def fix_formula(v,row):
    return re.sub(r'([A-Z]+)\d+',lambda m:m.group(1)+str(row),v) if isinstance(v,str) and v.startswith('=') else v
def tail(ws,first):
    return [[ws.cell(r,c).value for c in range(1,ws.max_column+1)] for r in range(first,ws.max_row+1)],[[copy.copy(ws.cell(r,c)._style) for c in range(1,ws.max_column+1)] for r in range(first,ws.max_row+1)]
# ---- 05-11: chèn 2 dòng P22 trước dòng A2 đầu tiên
ws=wb['Loss time 05-11']
first=next(r for r in range(3,ws.max_row+1) if ws.cell(r,8).value=='A2')
vals,sty=tail(ws,first)
p22tpl=wb['Loss time 04-11']
p22row=next(r for r in range(3,p22tpl.max_row+1) if p22tpl.cell(r,8).value=='P22')
tot_old=ws.max_row
# ghi lại từ dòng first: 2 dòng P22 mới rồi phần cũ dịch xuống 2
for i,(ar,hc) in enumerate(NEW):
    r=first+i
    for c in range(1,ws.max_column+1):
        ws.cell(r,c)._style=copy.copy(p22tpl.cell(p22row,c)._style)
        ws.cell(r,c).value=fix_formula(p22tpl.cell(p22row,c).value,r)
    ws.cell(r,2).value=dt.datetime(2026,11,5);ws.cell(r,7).value=ar;ws.cell(r,16).value=hc
for k,(row,st) in enumerate(zip(vals,sty)):
    r=first+2+k
    for c in range(1,len(row)+1):
        ws.cell(r,c)._style=copy.copy(st[c-1]);ws.cell(r,c).value=fix_formula(row[c-1],r)
# tổng: SUM(Q3:Qn)
for r in range(3,ws.max_row+1):
    if ws.cell(r,16).value=='Tổng':
        last=r-1;ws.cell(r,17).value=f'=SUM(Q3:Q{last})';n05=last
# ---- 06-11: thêm 2 dòng A2 (nhóm 05/11, ngày 1/12) trước dòng Tổng
ws=wb['Loss time 06-11']
tr=next(r for r in range(3,ws.max_row+1) if ws.cell(r,16).value=='Tổng')
vals,sty=tail(ws,tr)
ref=tr-1  # dòng A2 cuối (mẫu)
for i,(ar,hc) in enumerate(NEW):
    r=tr+i
    for c in range(1,ws.max_column+1):
        ws.cell(r,c)._style=copy.copy(ws.cell(ref,c)._style);ws.cell(r,c).value=fix_formula(ws.cell(ref,c).value,r)
    ws.cell(r,7).value=ar;ws.cell(r,16).value=hc
    ws.cell(r,12).value='A2 ngày 1/12 - nhóm 05/11'
    li=ws.cell(ref,11).value;ws.cell(r,11).value=re.sub(r'\d{2}/\d{2}$','05/11',li)
for k,(row,st) in enumerate(zip(vals,sty)):
    r=tr+2+k
    for c in range(1,len(row)+1):
        ws.cell(r,c)._style=copy.copy(st[c-1]);ws.cell(r,c).value=fix_formula(row[c-1],r)
for r in range(3,ws.max_row+1):
    if ws.cell(r,16).value=='Tổng':
        last=r-1;ws.cell(r,17).value=f'=SUM(Q3:Q{last})';n06=last
# ---- Tổng hợp: cập nhật phạm vi SUMIFS
ws=wb['Tổng hợp']
for row in ws.iter_rows():
    for c in row:
        v=c.value
        if isinstance(v,str) and "'Loss time 05-11'" in v: c.value=re.sub(r'\$(Q|H|G)\$3:\$\1\$\d+',lambda m:f'${m.group(1)}$3:${m.group(1)}${n05}',v)
        elif isinstance(v,str) and "'Loss time 06-11'" in v: c.value=re.sub(r'\$(Q|H|G)\$3:\$\1\$\d+',lambda m:f'${m.group(1)}$3:${m.group(1)}${n06}',v)
# ---- Nhân viên mới - A2: chèn 2 nhóm mới lên đầu
ws=wb['Nhân viên mới - A2']
hdr=[ws.cell(3,c).value for c in range(1,ws.max_column+1)]
ws.insert_rows(4,2)
for i,(ar,hc) in enumerate(NEW[::-1]):   # Assembly trước, Machining sau (giống thứ tự hiển thị)
    r=4+i
    for c in range(1,ws.max_column+1): ws.cell(r,c)._style=copy.copy(ws.cell(6,c)._style)
    ws.cell(r,1).value='05/11';ws.cell(r,2).value=ar;ws.cell(r,3).value=hc
    for c in range(4,ws.max_column+1):
        h=hdr[c-1]
        if not h or not str(h).startswith('Mã ngày'):
            ws.cell(r,c).value=None;continue
        d=str(h).split()[-1]
        ws.cell(r,c).value='-' if d in([f'{x:02d}/10' for x in range(5,32)]+[f'{x:02d}/11' for x in range(1,5)]) else ('P22' if d=='05/11' else None)
# A2 ngày k: ngày 06/11 = A2 ngày 1; các ngày sau (nếu có cột)
for i in range(2):
    r=4+i
    for c in range(4,ws.max_column+1):
        h=hdr[c-1]
        if h and str(h).startswith('Mã ngày') and str(h).split()[-1]=='06/11': ws.cell(r,c).value='A2 ngày 1'
# ---- Lịch A2: thêm 2 dòng nhóm 05/11
ws=wb['Lịch A2 nhân viên mới']
last=next(r for r in range(4,ws.max_row+1) if not ws.cell(r,1).value)-1
ws.insert_rows(last+1,2)
for i,(ar,hc) in enumerate([('Assembly',3),('Machining',2)]):
    r=last+1+i
    for c in range(1,ws.max_column+1):
        ws.cell(r,c)._style=copy.copy(ws.cell(last,c)._style);ws.cell(r,c).value=fix_formula(ws.cell(last,c).value,r)
    ws.cell(r,1).value='05/11';ws.cell(r,2).value=ar;ws.cell(r,3).value=hc;ws.cell(r,4).value='=DATE(2026,11,5)'
wb.save(F);print('saved',n05,n06)
