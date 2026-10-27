import openpyxl,re,collections,datetime as dt,pickle,warnings
import pandas as pd
warnings.simplefilter('ignore')
F='/root/.claude/uploads/64423226-a456-5e96-8225-6a5ffe123efb/04ede00f-Loss_time_White_Wood_2026_-_New_1.xlsx'
wb=openpyxl.load_workbook(F,data_only=True)
raw=[r for r in wb['Total loss time'].iter_rows(min_row=3,max_col=17,values_only=True) if r[1] and r[7]=='A10']
recs=[]
tmin=lambda x:(x.hour*60+x.minute) if hasattr(x,'hour') else None
for r in raw:
    recs.append(dict(date=r[1],area=r[6],text=(r[10] or '').strip(),dept=r[9],fr=r[12],to=r[13],hours=r[14] or 0,hc=r[15] or 0,labor=r[16] or 0,src='File Total loss time'))
S=pickle.load(open('sys.pkl','rb'))
for d in S:
    if d['LossTime Code']=='A10' and d['dt']>=dt.datetime(2026,10,4):
        fr=dt.datetime.strptime(d['Loss From Time'],'%H:%M').time();to=dt.datetime.strptime(d['Loss To Time'],'%H:%M').time()
        h=(tmin(to)-tmin(fr))/60
        recs.append(dict(date=d['dt'],area={'Raw Cutting':'RC - PANEL'}.get(d['Process Name'],d['Process Name']),text=(d['Description Loss Time'] or d['Notes'] or '').strip(),dept=d['Dept In Charge'],fr=fr,to=to,hours=h,hc=d['Quantity Workers'],labor=d['TotalLoss Hour'],src='Hệ thống (xuất 20/10)'))
df=pd.DataFrame(recs)
def code_of(t):
    seg=[s.strip() for s in t.split('-')]
    for s in seg:
        s2=s.replace(' ','')
        m=re.match(r'^([A-Za-z0-9]{2,5}?)0*(\d{1,2})$',s2)
        if m and re.search(r'[A-Za-z]',s2) and re.search(r'\d',s2) and not s2.lower().startswith(('wk','s1')):
            return (m.group(1).upper()+m.group(2).zfill(2))
    mm=re.search(r'(CNCR\d+)',t.replace(' ','').upper())
    if mm: 
        m=re.match(r'^(CNCR)0*(\d{1,2})$',mm.group(1)); return m.group(1)+m.group(2).zfill(2)
    tl=t.lower()
    mm=re.search(r'băng chuyền\s*line\s*(\d)',tl)
    if mm: return 'BC-LINE'+mm.group(1)
    return None
df['code']=df.text.map(code_of).astype(object).where(lambda x:x.notna(),None)
df['family']=df.code.map(lambda c: re.match(r'^([A-Z0-9]+?)\d{2}$',c).group(1) if isinstance(c,str) and re.match(r'^([A-Z0-9]+?)\d{2}$',c) else ('BC-LINE' if isinstance(c,str) and c.startswith('BC') else 'Không mã'))
df['code']=df.code.map(lambda c:c if isinstance(c,str) else 'Không có mã máy')
def cat(t):
    t=t.lower()
    if re.search(r'băng tải|con lăn|băng chuyền|cảo',t): return 'Băng tải - con lăn - băng chuyền'
    if re.search(r'dao|mũi|đạn|collet|kẹp',t): return 'Dao - dụng cụ cắt'
    if re.search(r'keo|nhiệt|lên nhiệt|ép',t): return 'Keo - nhiệt'
    if re.search(r'xì|dầu|van|hơi|hút|chân không|bụi|ống',t): return 'Khí nén - dầu - hút'
    if re.search(r'trục|động cơ|motor|hộp số|bạc|curoa|lò xo|đứt|gãy|sứt|bể|tuột|nhấc',t): return 'Cơ khí - truyền động'
    if re.search(r'nút|điện|biến tần|cảm biến|báo lỗi|lỗi|plc|màn hình|không lên|cúp',t): return 'Điện - điều khiển'
    return 'Chung chung (hư, không hoạt động...)'
df['cat']=df.text.map(cat)
df['fullday']=df.apply(lambda r: r.fr is not None and tmin(r.fr)<=5 and r.to is not None and tmin(r.to)>=23*60+55,axis=1)
df['dup']=df.duplicated(subset=['date','text','fr','to'],keep='first')
df['weekday']=df.date.dt.dayofweek
df['month']=df.date.dt.strftime('%Y-%m')
df['isoweek']=df.date.dt.isocalendar().week.astype(int)
df['start_hour']=df.fr.map(lambda x: x.hour if hasattr(x,'hour') else None)
df['clean']=~(df.fullday|df.dup)
df['hours_clean']=df.hours.where(df.clean,0);df['labor_clean']=df.labor.where(df.clean,0)
pickle.dump(df,open('a10_df.pkl','wb'))
print(len(df),df.src.value_counts().to_dict())
print('total labor',round(df.labor.sum(),1),'clean',round(df.labor_clean.sum(),1),'| hours',round(df.hours.sum(),1),'clean',round(df.hours_clean.sum(),1))
print('fullday rows',int(df.fullday.sum()),round(df.labor[df.fullday].sum(),1),'dup rows',int(df.dup.sum()),round(df.labor[df.dup].sum(),1))
print(df.groupby('family').agg(n=('labor','size'),labor=('labor','sum'),lab_c=('labor_clean','sum')).sort_values('lab_c',ascending=False).head(12).round(1))
print(df.groupby('code').agg(n=('labor','size'),lab_c=('labor_clean','sum'),h=('hours_clean','sum')).sort_values('lab_c',ascending=False).head(12).round(1))
print(df.groupby('cat').agg(n=('labor','size'),lab_c=('labor_clean','sum')).sort_values('lab_c',ascending=False).round(1))
print(df.groupby('month').agg(n=('labor','size'),lab=('labor','sum'),lab_c=('labor_clean','sum')).round(1))
print('no code',int((df.code=='Không có mã máy').sum()))
