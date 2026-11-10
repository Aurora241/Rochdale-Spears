"""Từ 10/11/2026: dòng "Không có bomtime sơn lòng trong" (thợ sơn lòng trong tủ) KHÔNG cần nhập (đồng nghiệp xác nhận).
Xóa các dòng này từ ngày <dd-mm> trở đi khỏi bảng S32 (sheet chính và Nhập liệu S32), sửa công thức, nguồn đối chiếu. Chạy trong outputs."""
import openpyxl,sys,re,datetime as dt,warnings
warnings.simplefilter('ignore')
f=sys.argv[1];frm=dt.datetime(2026,int(sys.argv[2][3:]),int(sys.argv[2][:2]))
wb=openpyxl.load_workbook(f)
ws=next(w for w in wb if w.title.startswith('S32 5-10'))
last=next(r for r in range(3,ws.max_row+1) if ws.cell(r,16).value=='Tổng')-1
dele=[r for r in range(3,last+1) if ws.cell(r,2).value and ws.cell(r,2).value>=frm and 'sơn lòng' in (ws.cell(r,11).value or '').lower()]
hours=0
for r in dele: hours+=((ws.cell(r,14).value.hour*60+ws.cell(r,14).value.minute)-(ws.cell(r,13).value.hour*60+ws.cell(r,13).value.minute))/60*ws.cell(r,16).value
for r in reversed(dele): ws.delete_rows(r)
nl=last-len(dele)
for r in range(3,nl+1):
    ws.cell(r,1).value=f'="W"&WEEKNUM(B{r},21)';ws.cell(r,15).value=f'=(N{r}-M{r})*24';ws.cell(r,17).value=f'=P{r}*O{r}'
ws.cell(nl+1,17).value=f'=SUM(Q3:Q{nl})'
for w in wb:
    for row in w.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith('='): c.value=re.sub(r'(\$[A-Z]+\$3:\$[A-Z]+\$)%d'%last,lambda m:m.group(1)+str(nl),c.value)
nlw=wb['Nhập liệu S32'];nd=[r for r in range(2,nlw.max_row+1) if nlw.cell(r,1).value and hasattr(nlw.cell(r,1).value,'year') and nlw.cell(r,1).value>=frm and 'sơn lòng' in (nlw.cell(r,3).value or '').lower()]
for r in reversed(nd): nlw.delete_rows(r)
dc=wb['Đối chiếu nguồn']
for r in range(4,16):
    if dc.cell(r,1).value=='Hình: thợ sơn': dc.cell(r,2).value=float(dc.cell(r,2).value)-hours
gc=wb['Ghi chú review'];g=gc.max_row
gc.cell(g+1,1)._style=gc.cell(g,1)._style
gc.cell(g+1,1).value=f'16. Từ 10/11/2026 (đồng nghiệp xác nhận): note "Không có bomtime sơn lòng trong" (thợ sơn lòng trong tủ) KHÔNG cần nhập. Đã xóa {len(dele)} dòng thợ sơn từ {frm:%d/%m} khỏi bảng này ({hours:.0f} giờ công); các tuần W42-W44 giữ nguyên như đã nhập trước đó.'
wb.save(f);print('đã xóa',len(dele),'dòng',hours,'giờ; dòng cuối',nl,'; Nhập liệu xóa',len(nd))
