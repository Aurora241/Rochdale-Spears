"""Bảng nhập liệu S32 theo bộ phận cho một tuần: python build_nhap_s32_tuan.py dd-mm dd-mm <file ra>. Chạy trong outputs."""
import openpyxl,datetime as dt,sys,warnings
sys.path.insert(0,'../scripts')
import build_nhap_lieu_theo_bo_phan as B
warnings.simplefilter('ignore')
d0=dt.datetime(2026,int(sys.argv[1][3:]),int(sys.argv[1][:2]));d1=dt.datetime(2026,int(sys.argv[2][3:]),int(sys.argv[2][:2]))
ws=openpyxl.load_workbook('Loss_time_S32_05-10_den_08-11.xlsx',data_only=True)['S32 5-10 đến 08-11'];rows=[]
for r in ws.iter_rows(min_row=3,values_only=True):
    if r[1] and hasattr(r[1],'strftime') and d0<=r[1]<=d1: rows.append((r[1],'S32',r[10],r[6],r[9],r[12],r[13],r[15]))
B.ORDER.insert(0,'S32');B.build(rows,sys.argv[3]);print(len(rows),'dòng')
