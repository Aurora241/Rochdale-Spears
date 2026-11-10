"""Từ 10/11/2026: dòng P26 'Nhận hàng gỗ đầu vào từ RC' (Process Raw Cutting) ghi Dept In Charge RAW CUTTING (đồng nghiệp: Raw Cutting không nhập 2 bộ phận).
Đổi cột J (Chart / bộ phận) của dòng này từ 'Machining' thành 'RC - PANEL' trong mọi sheet ngày. Chạy trong outputs với tên file Loss_time."""
import openpyxl,sys,warnings
warnings.simplefilter('ignore')
f=sys.argv[1];wb=openpyxl.load_workbook(f);n=0
for ws in wb:
    if not ws.title.startswith('Loss time '):continue
    for r in range(3,ws.max_row+1):
        if (ws.cell(r,11).value or '').startswith('Nhận hàng gỗ đầu vào') and ws.cell(r,10).value=='Machining':
            ws.cell(r,10).value='RC - PANEL';n+=1
wb.save(f);print('đã đổi',n,'dòng')
