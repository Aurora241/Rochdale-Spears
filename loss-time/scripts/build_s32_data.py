import openpyxl,datetime as dt,collections
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.utils import get_column_letter as L
Z='s32zip/Lost time S32/'
D0=dt.datetime(2026,10,5)
T=lambda h,m:dt.time(h,m)
def ld(f,sheet): return list(openpyxl.load_workbook(Z+f,read_only=True,data_only=True)[sheet].iter_rows(values_only=True))
R=[]   # dict rows
def add(date,src,item,qty,area,dept,lido,gc,fr,to,hc,flag='',items=None,chk=None):
    R.append(dict(date=date,src=src,item=item,qty=qty,area=area,dept=dept,lido=lido,gc=gc,fr=fr,to=to,hc=hc,flag=flag,items=items or []))
# 1. Beam saw (bản (5) mới nhất)
bs=[r for r in ld('LOST TIME BEMSAW ( NI ) (5).xlsx','data ') if isinstance(r[6],dt.datetime) and r[6]>=D0]
assert len(bs)==12
dm={'Machining CNC':'Machining','Machining':'Machining','Unassigned':'(chưa phân công)','QA do không test lú':'QC'}
for r in bs:
    ch=str(r[13]).strip()
    dept='QC' if ch.upper().startswith('QA') else dm[ch]
    flag='U' if ch=='Unassigned' else ''
    gc=f'Charge to: {ch}' if (ch=='Unassigned' or ch.upper().startswith('QA')) else ''
    add(r[6],'Beam saw',r[2],r[5],'RC - PANEL',dept,f'Beam saw - phát sinh (chi tiết {str(r[3]).strip()})',gc or f'STT {r[0]}',r[8],r[9],int(r[7]),flag)
# 2. UV hằng ngày (W41 từ 5/10): hàng làm việc 8h + OT 4h
uv=[r for r in ld('lost time khu vực UV (1).xlsx','LOST TIME') if isinstance(r[1],dt.datetime) and r[1]>=D0]
assert len(uv)==7
for r in uv:
    d=r[1];work=int(r[3] or 0);ot=int(r[4] or 0)
    if work: add(d,'UV hằng ngày',None,None,'RC - PANEL','Finishing','Không có bomtime chuyền UV','HC làm việc (ca thường)',T(7,30),T(15,30),work,'U' if d.weekday()==6 else '')
    if ot: add(d,'UV hằng ngày',None,None,'RC - PANEL','Finishing','Không có bomtime chuyền UV','HC tăng ca (OT 4 giờ)',T(15,30),T(19,30),ot,'')
# 3. UV tổng hợp phát sinh (Sheet1 bản (1))
s1=[r for r in ld('lost time khu vực UV (1).xlsx','Sheet1')[2:] if r[0] is not None]
g12=[r for r in s1 if r[6]==dt.datetime(2026,10,12)]
items12=[(r[2],r[3],r[5]) for r in g12]
hrs=[r for r in g12 if r[7] is not None]
assert len(g12)==23 and len(hrs)==2
for k,r in enumerate(hrs):
    add(dt.datetime(2026,10,12),'UV phát sinh',g12[0][2] if k==0 else None,None,'RC - PANEL','QC','Phát sinh khu vực UV - SLAT (23 mã hàng)',f'Charge to: QC A HIẾU; mã hàng xem sheet Chi tiết',r[8],r[9],int(r[7]),'',items12 if k==0 else [])
g13=[r for r in s1 if r[6]==dt.datetime(2026,10,13)]
assert len(g13)==1
r=g13[0];add(r[6],'UV phát sinh',r[2],r[5],'RC - PANEL','TF','Phát sinh khu vực UV - Back Panel','Charge to: TF',r[8],r[9],int(r[7]))
# 4. Ép nóng (bản (1))
ep=ld('FILE ÉP HÀNG SỬA ( KV ÉP NÓNG ) (1).xlsx','datta')
def grp(i0,i1): return [(ep[i-1][2],ep[i-1][3],ep[i-1][5]) for i in range(i0,i1+1)]
def eprow(i): return ep[i-1]
r=eprow(219);add(r[6],'Ép nóng',r[2],None,'RC - PANEL','P1','Hỗ trợ P1 ép nóng','Hỗ trợ P1 (3 mã hàng, 1 ca)',r[8],r[9],int(r[7]),'',grp(219,221))
r=eprow(222);add(r[6],'Ép nóng',r[2],None,'RC - PANEL','P1','Hỗ trợ P1 ép nóng - ép lót','Ép lót (2 mã hàng)',r[8],r[9],int(r[7]),'',grp(222,223))
r=eprow(224);add(r[6],'Ép nóng',r[2],None,'RC - PANEL','P1','Hỗ trợ P1 ép nóng - ép chính','Ép chính (2 mã hàng)',r[8],r[9],int(r[7]),'',grp(224,225))
# 5. Veneer hàng BS (bản 15/10)
ve=ld('26.10.15Lost time  veneer Hàng BS.xlsx','tổng')
v1,v2,v3=ve[1802],ve[1803],ve[1804]
add(v1[11],'Veneer',None,None,'RC - PANEL','TF','Làm mẫu (veneer)','2 mã (code 516, code 231), 1 ca',v1[8],v1[9],int(v1[7]),'',[(v1[1],v1[3],v1[5]),(v2[1],v2[3],v2[5])])
add(v3[11],'Veneer',None,None,'RC - PANEL','ME','Làm mẫu (veneer)','code 229',v3[8],v3[9],int(v3[7]),'',[(v3[1],v3[3],v3[5])])
# 6. Ba hình (Assembly)
for d in range(12,17): add(dt.datetime(2026,10,d),'Thợ sơn (hình)',None,None,'Assembly','Finishing','Không có bomtime sơn lòng trong','Thợ sơn lòng trong tủ - hỗ trợ Fin',T(7,30),T(19,30),12)
add(dt.datetime(2026,10,17),'Thợ sơn (hình)',None,None,'Assembly','Finishing','Không có bomtime sơn lòng trong','Thợ sơn lòng trong tủ - hỗ trợ Fin (thứ Bảy)',T(7,30),T(15,30),12,'P')
add(dt.datetime(2026,10,16),'Cào cước (hình)','RH893',None,'Assembly','Machining','Cào cước hỗ trợ MAC','Line 8 - hỗ trợ MAC',T(15,0),T(20,0),1,'P')
add(dt.datetime(2026,10,16),'Cào cước (hình)','RH5061',None,'Assembly','Machining','Cào cước hỗ trợ MAC','Line 7 - hỗ trợ MAC',T(8,0),T(12,0),1,'')
add(dt.datetime(2026,10,15),'Bắn cát (hình)','RH2252',734,'Assembly','Subcon','Bắn cát hỗ trợ Subcon Đại Xuyên (chân)','Line bắn cát - hỗ trợ Supcon Đại Xuyên',T(7,0),T(15,0),4,'A')
order=['Ép nóng','UV hằng ngày','Beam saw','Veneer','UV phát sinh','Thợ sơn (hình)','Cào cước (hình)','Bắn cát (hình)']
R.sort(key=lambda x:(x['date'],order.index(x['src']),x['fr']))
print('số dòng',len(R))
import pickle;pickle.dump(R,open('s32all.pkl','wb'))
hours=lambda x:((x['to'].hour*60+x['to'].minute)-(x['fr'].hour*60+x['fr'].minute))/60*x['hc']
by=collections.defaultdict(float)
for x in R: by[x['src']]+=hours(x)
for k,v in by.items(): print(k,round(v,3))
print('tổng',round(sum(by.values()),3))
