import openpyxl,datetime as dt,collections,sys
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
red=PatternFill('solid',fgColor='C00000');hf=Font(name=F,bold=True,color='FFFFFF');ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
band=PatternFill('solid',fgColor='1F4E78');grp=PatternFill('solid',fgColor='DDEBF7');gry=PatternFill('solid',fgColor='D9D9D9')
ORDER=['P20','P26','P22','A2']        # thứ tự mã cố định trong mọi bảng
AREAS=['Assembly','Machining','Raw Cutting']
WD=['Thứ Hai','Thứ Ba','Thứ Tư','Thứ Năm','Thứ Sáu','Thứ Bảy','Chủ nhật']
def tmin(x): return x.hour*60+x.minute if hasattr(x,'hour') else 0
def keyof(x): return (x[0],x[1],x[2],x[3],x[4],str(x[5]),str(x[6]),x[7])
NEWF=PatternFill('solid',fgColor='FFE699')
def sheet(wb,title,rows,marks=None):
    """rows: (date,code,notes,process,dept,from,to,hc)"""
    ws=wb.create_sheet(title)
    ws['A1']=f'BẢNG NHẬP LIỆU - {title.upper()}';ws['A1'].font=Font(name=F,bold=True,size=14)
    ws['A2']='Thứ tự cố định: theo ngày, trong mỗi ngày theo mã P20, P26, P22, A2. Nhập đúng thứ tự cột như form: LossTime Code | Notes | Process | Dept In Charge | From | To | Worker Qty. Nhập xong dòng nào thì đánh dấu X cột "Đã nhập".'
    ws['A2'].font=Font(name=F,italic=True)
    H=['Ngày (Loss Date)','LossTime Code','Notes','Process','Dept In Charge','From','To','Worker Qty','Giờ công (tham khảo)','Đã nhập']
    if marks is not None: H=H+['Ghi chú (so với lần gửi trước)']
    for i,h in enumerate(H,1):
        c=ws.cell(4,i,h);c.fill=red if i<=8 else gry;c.font=hf if i<=8 else Font(name=F,bold=True);c.alignment=ctr;c.border=bd
    ws.row_dimensions[4].height=32
    r=5;days=sorted({x[0] for x in rows});first=r
    for d in days:
        dr=[x for x in rows if x[0]==d]
        dr.sort(key=lambda x:(ORDER.index(x[1]) if x[1] in ORDER else 9,tmin(x[5]),str(x[2])))
        nnew=sum(1 for x in dr if marks is not None and keyof(x) in marks)
        ws.cell(r,1,f'{WD[d.weekday()]} {d:%d/%m/%Y}: {len(dr)} dòng'+(f' (có {nnew} dòng MỚI/SỬA)' if nnew else ''));
        for c in range(1,12 if marks is not None else 11):
            x=ws.cell(r,c);x.fill=band;x.font=Font(name=F,bold=True,color='FFFFFF')
        bandrow=r;r+=1;ds=r
        prev=None
        for x in dr:
            if prev is not None and x[1]!=prev: pass
            prev=x[1]
            row=[x[0],x[1],x[2],x[3],x[4],x[5],x[6],x[7],f'=IF(H{r}="","",(G{r}-F{r})*24*H{r})',None]
            for c,v in enumerate(row,1):
                cc=ws.cell(r,c,v);cc.font=Font(name=F);cc.border=bd;cc.alignment=Alignment(horizontal='left' if c==3 else 'center',vertical='center',wrap_text=c==3)
                if x[1] in('P22','A2'): cc.fill=grp
            ws.cell(r,1).number_format='dd/mm/yyyy';ws.cell(r,6).number_format=ws.cell(r,7).number_format='hh:mm';ws.cell(r,9).number_format='0.0'
            if marks is not None:
                m=marks.get(keyof(x));cc=ws.cell(r,11,m[0] if m else 'Đã gửi trước đó');cc.border=bd;cc.alignment=Alignment(wrap_text=True,vertical='center')
                cc.font=Font(name=F,bold=bool(m),color='C00000' if m else '7F7F7F')
                if m:
                    for c in range(1,10): ws.cell(r,c).fill=NEWF
                    ws.cell(r,11).fill=NEWF
            r+=1
        ws.cell(bandrow,9,f'=SUM(I{ds}:I{r-1})').number_format='0.0'
        ws.cell(bandrow,9).font=Font(name=F,bold=True,color='FFFFFF')
        r+=1
    ws.cell(r,8,'Tổng').font=Font(name=F,bold=True)
    ws.cell(r,9,f'=SUMIF(A{first}:A{r-1},">0",I{first}:I{r-1})').font=Font(name=F,bold=True);ws.cell(r,9).number_format='#,##0.0'
    for col,w in zip('ABCDEFGHIJK',[22,12,46,14,18,8,8,10,12,9,46]): ws.column_dimensions[col].width=w
    ws.freeze_panes='A5'
    return ws
def read_prev(path):
    wb=openpyxl.load_workbook(path,data_only=True);rows=[]
    for ws in wb.worksheets:
        if ws.title not in AREAS: continue
        for r in ws.iter_rows(min_row=5,values_only=True):
            if isinstance(r[0],dt.datetime): rows.append((r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7]))
    return rows
def diff(rows,prev):
    cs=collections.Counter(keyof(x) for x in prev);cm=collections.Counter(keyof(x) for x in rows)
    byk={keyof(x):x for x in rows};byp={keyof(x):x for x in prev}
    added=list((cm-cs).elements());removed=list((cs-cm).elements())
    loose=lambda k:(k[0],k[1],k[2],k[3])
    pool=collections.defaultdict(list)
    for k in removed: pool[loose(k)].append(k)
    marks={};acts=[]
    for k in added:
        c=pool.get(loose(k))
        if c:
            o=c.pop(0);ch=[f'{n} {a} -> {b}' for n,a,b in (('Dept',o[4],k[4]),('From',o[5],k[5]),('To',o[6],k[6]),('Worker Qty',o[7],k[7])) if a!=b]
            marks[k]=('SỬA: '+', '.join(ch),);acts.append(('Sửa',byk[k],'Dòng đã gửi trước đó, đổi: '+', '.join(ch)))
        else:
            marks[k]=('MỚI: thêm so với lần gửi trước',);acts.append(('Thêm',byk[k],'Chưa có ở lần gửi trước'))
    gone=[k for lst in pool.values() for k in lst]
    for k in gone: acts.append(('Bỏ',byp[k],'Có ở lần gửi trước, nay không còn (xóa dòng này nếu đã nhập)'))
    return marks,acts
def summary(wb,acts,label):
    ws=wb.create_sheet('CẦN NHẬP MỚI',0)
    cnt=collections.Counter(a[0] for a in acts)
    ws['A1']='CHỈ NHỮNG DÒNG CẦN LÀM SO VỚI LẦN GỬI TRƯỚC'+(f' ({label})' if label else '');ws['A1'].font=Font(name=F,bold=True,size=14)
    ws['A2']=f'Thêm: {cnt.get("Thêm",0)} dòng   |   Sửa: {cnt.get("Sửa",0)} dòng   |   Bỏ (xóa nếu đã nhập): {cnt.get("Bỏ",0)} dòng. Các sheet Assembly / Machining / Raw Cutting vẫn đủ cả tuần; dòng mới có nền vàng và cột Ghi chú.';ws['A2'].font=Font(name=F,italic=True,color='C00000')
    H=['Hành động','Bộ phận','Ngày (Loss Date)','LossTime Code','Notes','Process','Dept In Charge','From','To','Worker Qty','Ghi chú','Đã nhập']
    for i,h in enumerate(H,1):
        c=ws.cell(4,i,h);c.fill=red if i<=10 else gry;c.font=hf if i<=10 else Font(name=F,bold=True);c.alignment=ctr;c.border=bd
    colr={'Thêm':'E2EFDA','Sửa':'FFF2CC','Bỏ':'FCE4D6'}
    acts=sorted(acts,key=lambda a:(AREAS.index(a[1][3]) if a[1][3] in AREAS else 9,a[1][0],ORDER.index(a[1][1]) if a[1][1] in ORDER else 9,{'Thêm':0,'Sửa':1,'Bỏ':2}[a[0]]))
    r=5
    for act,x,why in acts:
        for i,v in enumerate([act,x[3],x[0],x[1],x[2],x[3],x[4],x[5],x[6],x[7],why,None],1):
            c=ws.cell(r,i,v);c.font=Font(name=F);c.border=bd;c.alignment=Alignment(horizontal='left' if i in(5,11) else 'center',vertical='center',wrap_text=i in(5,11))
        ws.cell(r,1).fill=PatternFill('solid',fgColor=colr[act]);ws.cell(r,3).number_format='dd/mm/yyyy';ws.cell(r,8).number_format=ws.cell(r,9).number_format='hh:mm'
        r+=1
    if not acts: ws.cell(5,1,'Không có thay đổi so với lần gửi trước').font=Font(name=F,italic=True)
    for col,w in zip('ABCDEFGHIJKL',[10,13,14,10,46,13,20,8,8,10,48,9]): ws.column_dimensions[col].width=w
    ws.freeze_panes='A5'
def build(rows,out,fixes=None,extra=None,prev=None,label=''):
    wb=openpyxl.Workbook();wb.remove(wb.active)
    marks=None;acts=[]
    if prev:
        marks,acts=diff(rows,prev)
    for a in AREAS:
        sheet(wb,a,[x for x in rows if x[3]==a],marks)
    if prev is not None: summary(wb,acts,label)
    wb.save(out)
if __name__=='__main__':
    # dùng: python byarea.py Nhap_lieu_xxx.xlsx ngay_dau ngay_cuoi out.xlsx  (ngày dạng dd-mm)
    src,d0,d1,out=sys.argv[1:5]
    wb=openpyxl.load_workbook(src,data_only=True)
    rows=[]
    for ws in wb.worksheets:
        if not ws.title.startswith('Nhập '): continue
        dd,mm=ws.title[5:].split('-');d=dt.datetime(2026,int(mm),int(dd))
        if not(dt.datetime(2026,int(d0[3:]),int(d0[:2]))<=d<=dt.datetime(2026,int(d1[3:]),int(d1[:2]))): continue
        for r in ws.iter_rows(min_row=2,values_only=True):
            if r[0]: rows.append((d,r[0],r[1],r[2],r[3],r[4],r[5],r[6]))
    prev=read_prev(sys.argv[5]) if len(sys.argv)>5 else None
    build(rows,out,prev=prev,label=sys.argv[6] if len(sys.argv)>6 else '')
