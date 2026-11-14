"""Ngày chỉ Assembly tăng ca (các khu vực khác không tăng ca): đổi To của P20 và P26 Assembly thành 20:00. python patch_assembly_ot.py <file> dd-mm (chạy trong outputs)"""
import openpyxl,sys,datetime as dt,warnings
warnings.simplefilter('ignore')
f,nm=sys.argv[1:3];wb=openpyxl.load_workbook(f);ws=wb[f'Loss time {nm}'];n=0
for r in range(3,ws.max_row+1):
    if ws.cell(r,7).value=='Assembly' and ws.cell(r,8).value in('P20','P26'):
        ws.cell(r,14).value=dt.time(20,0);n+=1
wb.save(f);print('đã đổi',n,'dòng Assembly sang 20:00')
