"""Hoàn tất tuần W45: nhân viên mới 6/11 (Assembly 2), Thứ Bảy 7/11 (không tăng ca), Chủ nhật 8/11 (chỉ Assembly).
Chạy trong thư mục outputs sau patch_hire_05-11.py."""
import openpyxl,copy,re,warnings,datetime as dt
warnings.simplefilter('ignore')
F='Loss_time_05-10_den_08-11.xlsx'
wb=openpyxl.load_workbook(F)
D=lambda d,m:dt.datetime(2026,m,d)
SUN_OLD={D(25,10),D(1,11)};SUN_NEW=SUN_OLD|{D(8,11)}
def kday(g,area,day,suns):
    n=0;x=g
    while x<day:
        x+=dt.timedelta(days=1)
        if x.weekday()!=6 or (area=='Assembly' and x in suns): n+=1
    return n
def endday(g,area,suns):
    x=g;n=0
    while n<12:
        x+=dt.timedelta(days=1)
        if x.weekday()!=6 or (area=='Assembly' and x in suns): n+=1
    return x
def fix_formula(v,row):
    return re.sub(r'([A-Z]+)\d+',lambda m:m.group(1)+str(row),v) if isinstance(v,str) and v.startswith('=') else v
def clone_row(ws,src,dst,tpl=None):
    t=tpl or ws
    for c in range(1,ws.max_column+1):
        ws.cell(dst,c)._style=copy.copy(t.cell(src,c)._style)
# ---- nhóm nhân viên mới từ sheet "Nhân viên mới - A2"
nv=wb['Nhân viên mới - A2']
hd=[nv.cell(3,c).value for c in range(1,nv.max_column+1)]
col={str(h).split()[-1]:c for c,h in enumerate(hd,1) if h and str(h).startswith('Mã ngày')}
groups=[]
for r in range(4,nv.max_row+1):
    g=nv.cell(r,1).value
    if g and re.match(r'\d\d/\d\d$',str(g)):
        dd,mm=map(int,g.split('/'));groups.append((r,D(dd,mm),nv.cell(r,2).value,nv.cell(r,3).value,g))
groups.append((None,D(6,11),'Assembly',2,'06/11'))
AREA=lambda a:'Assembly' if a=='Assembly' else a
# kiểm tra hàm kday với sheet 06-11
s6=wb['Loss time 06-11'];chk=0
for r in range(3,s6.max_row+1):
    if s6.cell(r,8).value=='A2':
        k=int(re.search(r'ngày (\d+)/12',s6.cell(r,12).value).group(1));gd=re.search(r'nhóm (\d\d/\d\d)',s6.cell(r,12).value).group(1)
        ar=s6.cell(r,7).value
        match=[g for g in groups if g[4]==gd and AREA(g[2])==AREA(ar) or g[4]==gd and ar.startswith('Raw') and g[2].startswith(('Raw','RC'))]
        assert match and kday(match[0][1],AREA(ar),D(6,11),SUN_OLD)==k,(r,gd,ar,k); chk+=1
print('kiểm tra k 06/11 khớp',chk,'dòng')
# ---- 1) 06/11: thêm P22 Assembly 2
tpl=wb['Loss time 04-11'];p22=next(r for r in range(3,tpl.max_row+1) if tpl.cell(r,8).value=='P22')
first=next(r for r in range(3,s6.max_row+1) if s6.cell(r,8).value=='A2')
rows=[[s6.cell(r,c).value for c in range(1,s6.max_column+1)] for r in range(first,s6.max_row+1)]
sty=[[copy.copy(s6.cell(r,c)._style) for c in range(1,s6.max_column+1)] for r in range(first,s6.max_row+1)]
for c in range(1,s6.max_column+1):
    s6.cell(first,c)._style=copy.copy(tpl.cell(p22,c)._style);s6.cell(first,c).value=fix_formula(tpl.cell(p22,c).value,first)
s6.cell(first,2).value=D(6,11);s6.cell(first,7).value='Assembly';s6.cell(first,16).value=2
for k,(row,st) in enumerate(zip(rows,sty)):
    r=first+1+k
    for c in range(1,len(row)+1): s6.cell(r,c)._style=copy.copy(st[c-1]);s6.cell(r,c).value=fix_formula(row[c-1],r)
for r in range(3,s6.max_row+1):
    if s6.cell(r,16).value=='Tổng': s6.cell(r,17).value=f'=SUM(Q3:Q{r-1})';n06=r-1
# ---- 2) tạo sheet ngày mới từ mẫu
def a2rows(day,suns,sat):
    out=[]
    for (_,g,ar,hc,gtxt) in groups:
        k=kday(g,AREA(ar),day,suns)
        if 1<=k<=12 and (sat or AREA(ar)=='Assembly'): out.append((g,ar,hc,k,gtxt))
    order={'Assembly':0,'Machining':1}
    out.sort(key=lambda x:(x[0],order.get(x[1],2)))
    return out
def make(name,tplname,day,fixed_last,suns,sat,pos_after):
    t=wb[tplname];ws=wb.copy_worksheet(t);ws.title=name
    wb.move_sheet(ws,offset=wb.sheetnames.index(pos_after)+1-wb.sheetnames.index(name))
    ws['A1']=re.sub(r'\d\d/\d\d/2026',day.strftime('%d/%m/%Y'),t['A1'].value)
    tr=next(r for r in range(3,t.max_row+1) if t.cell(r,16).value=='Tổng')
    a2tpl=next(r for r in range(3,tr) if t.cell(r,8).value=='A2')
    for r in range(3,fixed_last+1): ws.cell(r,2).value=day
    # xóa phần cũ từ dòng fixed_last+1
    for r in range(fixed_last+1,ws.max_row+1):
        for c in range(1,ws.max_column+1): ws.cell(r,c).value=None
    rows=a2rows(day,suns,sat);r=fixed_last+1
    ref=[ [t.cell(a2tpl,c).value for c in range(1,ws.max_column+1)] ]
    for (g,ar,hc,k,gtxt) in rows:
        clone_row(ws,a2tpl,r,t)
        for c,v in enumerate(ref[0],1): ws.cell(r,c).value=fix_formula(v,r)
        ws.cell(r,2).value=day;ws.cell(r,7).value=ar;ws.cell(r,16).value=hc
        ws.cell(r,12).value=f'A2 ngày {k}/12 - nhóm {gtxt}'
        ws.cell(r,11).value=re.sub(r'\d\d/\d\d$',gtxt,ref[0][10]);r+=1
    last=r-1
    clone_row(ws,tr,r,t)
    for c in range(1,ws.max_column+1): ws.cell(r,c).value=fix_formula(t.cell(tr,c).value,r)
    ws.cell(r,17).value=f'=SUM(Q3:Q{last})'
    nr=tr+2
    for c in range(1,ws.max_column+1): ws.cell(r+2,c)._style=copy.copy(t.cell(nr,c)._style);ws.cell(r+2,c).value=t.cell(nr,c).value
    for rr in range(r+3,ws.max_row+1+3):
        for c in range(1,ws.max_column+1): ws.cell(rr,c).value=None
    return ws,last,rows
w7,n07,r7=make('Loss time 07-11','Loss time 31-10',D(7,11),22,SUN_NEW,True,'Loss time 06-11')
w8,n08,r8=make('Loss time 08-11','Loss time 01-11',D(8,11),5,SUN_NEW,False,'Loss time 07-11')
print('07/11 A2',[(x[2],x[1],x[3],x[4]) for x in r7]);print('08/11 A2',[(x[2],x[1],x[3],x[4]) for x in r8])
# ---- 3) Tổng hợp: thêm 2 khối
ws=wb['Tổng hợp']
note=ws.cell(281,1).value;start=273
for r in range(281,282):
    for c in range(1,ws.max_column+1): ws.cell(r,c).value=None
def block(top,sheet,last,day):
    for i in range(7):
        src=start+i;clone_row(ws,src,top+i)
        for c in range(1,ws.max_column+1):
            v=ws.cell(src,c).value
            if isinstance(v,str) and v.startswith('='):
                v=v.replace("'Loss time 06-11'",f"'{sheet}'");v=re.sub(r'\$(Q|H|G)\$3:\$\1\$\d+',lambda m:f'${m.group(1)}$3:${m.group(1)}${last}',v)
                v=re.sub(r'(?<![A-Z$])([A-Z]+)(27[3-9])\b',lambda m:m.group(1)+str(top+int(m.group(2))-start),v)
            ws.cell(top+i,c).value=v
    ws.cell(top,1).value=f'Ngày {day.strftime("%d/%m/%Y")}'
block(282,'Loss time 07-11',n07,D(7,11));block(291,'Loss time 08-11',n08,D(8,11))
ws.cell(300,1).value=note;ws.cell(300,1)._style=copy.copy(ws.cell(280,1)._style)
# ---- 4) Nhân viên mới - A2: thêm 2 cột ngày + cập nhật dòng nhóm 06/11
c0=nv.max_column  # cột "Ngày cuối"
nv.insert_cols(c0,2)
for c in (c0,c0+1):
    nv.cell(3,c)._style=copy.copy(nv.cell(3,c0-1)._style)
nv.cell(3,c0).value='Mã ngày 07/11';nv.cell(3,c0+1).value='Mã ngày 08/11'
newrow=None
# chèn dòng nhóm 06/11 Assembly lên đầu
nv.insert_rows(4,1)
for c in range(1,nv.max_column+1): nv.cell(4,c)._style=copy.copy(nv.cell(5,c)._style)
nv.cell(4,1).value='06/11';nv.cell(4,2).value='Assembly';nv.cell(4,3).value=2
for r in range(4,nv.max_row+1):
    g=nv.cell(r,1).value
    if not(g and re.match(r'\d\d/\d\d$',str(g))): continue
    dd,mm=map(int,g.split('/'));gd=D(dd,mm);ar=nv.cell(r,2).value
    for dcol,day in ((c0,D(7,11)),(c0+1,D(8,11))):
        c=nv.cell(r,dcol);c._style=copy.copy(nv.cell(r,c0-1)._style)
        if gd==day: c.value='P22'
        else:
            k=kday(gd,AREA(ar),day,SUN_NEW);kprev=kday(gd,AREA(ar),day-dt.timedelta(days=1),SUN_NEW)
            if day.weekday()==6 and AREA(ar)!='Assembly':
                done=kday(gd,AREA(ar),day,SUN_NEW)>=12 and kprev>=12
                c.value='Đã xong' if kprev>=12 else '-'
            elif k>12 or (kprev>=12): c.value='Đã xong'
            elif k==0: c.value='-'
            else: c.value=f'A2 ngày {k}'
    # sửa 06/11 của nhóm mới và ngày cuối
    if gd==D(6,11):
        c6=col['06/11'];nv.cell(r,c6).value='P22'
        for cc in range(4,c0):
            h=nv.cell(3,cc).value
            if h and str(h).startswith('Mã ngày') and nv.cell(r,cc).value is None: nv.cell(r,cc).value='-'
    endc=nv.cell(r,c0+2)
    if isinstance(endc.value,str) and endc.value.startswith('=DATE') and AREA(ar)=='Assembly':
        # nhóm Assembly còn đang chạy: tính lại ngày cuối có Chủ nhật 8/11
        kv=kday(gd,'Assembly',D(6,11),SUN_NEW)
        if kv<=12:
            e=endday(gd,'Assembly',SUN_NEW);endc.value=f'=DATE({e.year},{e.month},{e.day})'
    if gd==D(6,11) and endc.value is None:
        e=endday(gd,'Assembly',SUN_NEW);endc.value=f'=DATE({e.year},{e.month},{e.day})'
# nhóm 05/11 Machining/Assembly và 04/11,03/11: ngày cuối nếu thiếu
for r in range(4,nv.max_row+1):
    g=nv.cell(r,1).value
    if g in('05/11',) and nv.cell(r,c0+2).value is None:
        e=endday(D(5,11),AREA(nv.cell(r,2).value),SUN_NEW);nv.cell(r,c0+2).value=f'=DATE({e.year},{e.month},{e.day})'
# ---- 5) Lịch A2: thêm dòng nhóm 06/11
ws=wb['Lịch A2 nhân viên mới']
last=next(r for r in range(4,ws.max_row+1) if not ws.cell(r,1).value)-1
ws.insert_rows(last+1,1)
clone_row(ws,last,last+1)
for c in range(1,ws.max_column+1): ws.cell(last+1,c).value=fix_formula(ws.cell(last,c).value,last+1)
ws.cell(last+1,1).value='06/11';ws.cell(last+1,2).value='Assembly';ws.cell(last+1,3).value=2;ws.cell(last+1,4).value='=DATE(2026,11,6)'
wb.save(F);print('saved',n06,n07,n08)
