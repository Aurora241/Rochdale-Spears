import json,subprocess,re
toc=json.load(open('toc.json'))
n=int(re.search(r'Pages:\s+(\d+)',subprocess.run(['pdfinfo','Bao_cao_thu_viec.pdf'],capture_output=True,text=True).stdout).group(1))
txt=[subprocess.run(['pdftotext','-f',str(i),'-l',str(i),'-layout','Bao_cao_thu_viec.pdf','-'],capture_output=True,text=True).stdout for i in range(1,n+1)]
pages={}
for lvl,t,name in toc:
    for i in range(2,n):
        if any(l.strip()==t or (len(t)>40 and l.strip().startswith(t[:40])) for l in txt[i].splitlines()):
            pages[name]=i-1;break
    else: print('MISSING',t)
json.dump(pages,open('pages.json','w'));print(pages,n)
