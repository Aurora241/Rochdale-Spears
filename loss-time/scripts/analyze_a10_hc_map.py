import pickle,pandas as pd,openpyxl,collections
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
R=pickle.load(open('hcmap2.pkl','rb'));df=pickle.load(open('a10_df.pkl','rb'));out=pickle.load(open('a10_hc_out.pkl','rb'))
G=['Mgr','Sr.Sup','Sup','TE','SL1','SL2','SL3','RL','OJT','Expeditor','Forklift'];OPS=['TE','SL1','SL2','SL3','RL','OJT']
s=lambda v: sum((x or 0) for i,x in enumerate(v) if G[i] in OPS)
c=df[df.clean&(df.family!='BC-LINE')]
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
red=PatternFill('solid',fgColor='C00000');hf=Font(name=F,bold=True,color='FFFFFF');ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
grn=PatternFill('solid',fgColor='E2EFDA');yel=PatternFill('solid',fgColor='FFF2CC');org=PatternFill('solid',fgColor='FCE4D6')
def table(ws,r0,headers,data,widths,fmt=None):
    for i,h in enumerate(headers,1):
        x=ws.cell(r0,i,h);x.fill=red;x.font=hf;x.alignment=ctr;x.border=bd
    for k,row in enumerate(data):
        for i,v in enumerate(row,1):
            x=ws.cell(r0+1+k,i,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='left' if isinstance(v,str) and len(v)>16 else 'center',vertical='center',wrap_text=True)
            if fmt and i in fmt: x.number_format=fmt[i]
    for i,w in enumerate(widths): ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=w
near={'EDB01':'Họ EDB trong lịch sử ghi 2-3 người, mapping 4 người (1 SL3 + 3 RL) có thể cao hơn thực tế; nên hỏi','EDB05':'Lịch sử 1 lần ghi 3 người, mapping thực tế 4 và tiêu chuẩn 3','CNCR09':'Tên gần giống "Máy chép cạnh thẳng" (60-0046): mapping 2 người, mã máy khác nên không dùng','HPTT02':'Tên "Máy cảo cong" (HF01=>HF08, nhóm 8 máy, 6 người cả nhóm): không dùng','WBS20':'Chưa có mã WBS20; máy WBS khác khu RC panel: WBS11 2, WBS09/10/21 3, WBS08 5 người','CNCR02':'Chưa có mã trong mapping','VC14':'Chưa có mã trong mapping','VC':'Chưa có mã trong mapping'}
rows=[]
for d,ar,code,name,fail,h,n,dist,share,fm,famtxt,prop,cf in out:
    m=R.get(code) if code!='(không mã)' else None
    mact=s(m['act']) if m else None;mstd=s(m['std']) if m else None
    hmode=None
    if n:
        x=c[c.code==code];hmode=int(x.hc.mode().iloc[0])
    if n and m: concl='Hai nguồn TRÙNG' if hmode==mact else 'Hai nguồn LỆCH'
    elif m: concl='Chỉ có mapping'
    elif n: concl='Chỉ có lịch sử'
    else: concl='Không có nguồn nào'
    if concl=='Hai nguồn TRÙNG': prop2=hmode
    elif concl=='Hai nguồn LỆCH': prop2=f'{hmode} (lịch sử) hoặc {mact} (mapping)'
    elif concl=='Chỉ có mapping': prop2=mact
    elif concl=='Chỉ có lịch sử': prop2=hmode
    else: prop2=None
    ask={'Hai nguồn TRÙNG':'Xác nhận nhanh','Hai nguồn LỆCH':'Hỏi tổ trưởng','Chỉ có mapping':'Xác nhận (độ tin cậy vừa)','Chỉ có lịch sử':'Xác nhận (độ tin cậy vừa)','Không có nguồn nào':'Phải hỏi'}[concl]
    rows.append((d,ar,code,name,h,n,dist if n else 'chưa từng có',mact,mstd,(m['group'] or m['area']) if m else '',concl,prop2,ask,near.get(code,'')))
wb=openpyxl.Workbook();wb.remove(wb.active)
ws=wb.create_sheet('1. Đề xuất theo 2 nguồn');ws['A1']='SỐ NGƯỜI ĐỀ XUẤT CHO 19 SỰ CỐ 19/10-30/10 VÀ WBS20: LỊCH SỬ A10 VÀ FILE HC MAPPING (CHỈ ĐỂ XÁC NHẬN, KHÔNG PHẢI SỐ NHẬP)';ws['A1'].font=Font(name=F,bold=True,size=13)
ws['A2']='Mapping: "Thực tế vận hành" = số người cấp TE/SL1/SL2/SL3/RL/OJT mapped cho máy đó (không tính quản lý). Hai nguồn trùng thì độ tin cậy cao hơn, lệch hoặc không có thì phải hỏi tổ trưởng.';ws['A2'].font=Font(name=F,italic=True,color='C00000')
H=['Ngày','Khu vực','Mã máy','Tên máy','Giờ máy dừng','Số sự cố cùng máy trong lịch sử','Số người đã ghi trong A10','Mapping: thực tế vận hành','Mapping: tiêu chuẩn','Nhóm máy trong mapping','Kết luận','Số người để xác nhận','Việc cần làm','Ghi chú']
table(ws,4,H,rows,[10,12,12,26,10,12,32,12,12,24,18,22,22,52],{5:'0.00'})
for k,row in enumerate(rows):
    ws.cell(5+k,11).fill=grn if row[10]=='Hai nguồn TRÙNG' else (org if row[10] in('Hai nguồn LỆCH','Không có nguồn nào') else yel)
cnt=collections.Counter(r[10] for r in rows)
ws.cell(6+len(rows),1,'Tóm tắt: '+'; '.join(f'{k}: {v}' for k,v in cnt.items())).font=Font(name=F,bold=True)
# so sánh 62 máy
ws2=wb.create_sheet('2. So sánh mapping - lịch sử');ws2['A1']='MAPPING VÀ LỊCH SỬ A10 CHO CÁC MÁY CÓ CẢ HAI NGUỒN';ws2['A1'].font=Font(name=F,bold=True,size=13)
data=[];exact=0;within=0;lines_eq=0;lines=0
for code,x in c.groupby('code'):
    if code in R:
        mode=int(x.hc.mode().iloc[0]);ma=s(R[code]['act']);ms=s(R[code]['std'])
        eq=(mode==ma);exact+=eq;within+=abs(mode-ma)<=1
        lines+=len(x);lines_eq+=int((x.hc==ma).sum())
        data.append((code,R[code]['group'] or '',R[code]['area'],len(x),mode,round(x.hc.mean(),2),ma,ms,R[code]['qty'],'Trùng' if eq else ('Lệch 1' if abs(mode-ma)==1 else 'Lệch >1')))
data.sort(key=lambda r:-r[3])
table(ws2,3,['Mã máy','Nhóm máy','Khu vực mapping','Số sự cố','Số người hay gặp (lịch sử)','Số người trung bình','Mapping thực tế vận hành','Mapping tiêu chuẩn','Mapping QTY (1 người/máy)','So sánh'],data,[12,30,16,10,14,12,14,14,14,12])
n=len(data)
txt=[f'Có {n} mã máy xuất hiện ở cả hai nguồn ({int(sum(r[3] for r in data))} dòng A10).',f'Số người hay gặp trùng mapping thực tế vận hành: {exact}/{n} máy ({100*exact/n:.0f}%). Lệch không quá 1 người: {within}/{n} máy ({100*within/n:.0f}%).',f'Tính theo từng dòng A10: số người ghi bằng mapping ở {lines_eq}/{lines} dòng ({100*lines_eq/lines:.0f}%).','Cột QTY của mapping luôn là 1 người mỗi máy nên không dùng được cho A10 (A10 thường 2-3 người vì cả người phụ cũng dừng).']
for k,tx in enumerate(txt): ws2.cell(5+n+k,1,tx).font=Font(name=F,bold=(k==0))
ws2.freeze_panes='A4'
# sheets cũ
old=openpyxl.load_workbook('A10_so_nguoi_theo_may.xlsx')
for nm in ('2. Số người theo mã máy','3. Theo họ máy'):
    src=old[nm];dst=wb.create_sheet(nm.replace('2.','3.').replace('3. Theo','4. Theo') if nm.startswith('3.') else '3. Số người theo mã máy')
    for row in src.iter_rows():
        for cell in row:
            if cell.value is not None:
                x=dst.cell(cell.row,cell.column,cell.value);x.font=Font(name=F,bold=cell.font.bold,color=cell.font.color.rgb if cell.font and cell.font.color and isinstance(cell.font.color.rgb,str) else None);x.border=bd;x.alignment=cell.alignment.copy();x.fill=cell.fill.copy();x.number_format=cell.number_format
    for k,dim in src.column_dimensions.items(): dst.column_dimensions[k].width=dim.width
wb.save('A10_so_nguoi_theo_may_v2.xlsx')
print(cnt);print(txt)
for r in rows: print(r[0],r[2],'|',r[5],r[6],'|map',r[7],r[8],'|',r[10],'|',r[11])
