import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
weeks=['W41\n5-11/10','W42\n12-18/10','W43\n19-25/10','W44\n26/10-1/11','W45\n2-6/11 (đến 6/11)']
base=[1542,1542,1574,1574,1360]
train=[1994,1775,1504.5,1145,690.5]
s32=[880.8,1662.8,1782.8,1233.1,20.3]
fig,ax=plt.subplots(figsize=(7.0,3.9),dpi=200)
x=range(len(weeks));w=0.62
b1=ax.bar(x,base,w,color='#1F3A5F',label='P20 + P26 (mức nền cố định)')
b2=ax.bar(x,train,w,bottom=base,color='#4F81BD',label='P22 + A2 (nhân viên mới)')
b3=ax.bar(x,s32,w,bottom=[a+b for a,b in zip(base,train)],color='#C55A11',label='S32 (hỗ trợ sản xuất)')
for i in x:
    tot=base[i]+train[i]+s32[i]
    ax.text(i,tot+60,f'{tot:,.0f}'.replace(',','.'),ha='center',va='bottom',fontsize=9,fontweight='bold')
ax.set_xticks(list(x));ax.set_xticklabels(weeks,fontsize=8.5)
ax.set_ylabel('Giờ công',fontsize=9)
ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v,p:f'{v:,.0f}'.replace(',','.')))
ax.tick_params(axis='y',labelsize=8.5)
ax.spines[['top','right']].set_visible(False)
ax.legend(fontsize=8.5,frameon=False,loc='upper right')
ax.set_ylim(0,6600)
ax.text(4,base[4]+train[4]+s32[4]+420,'S32 tuần này\nchưa đủ nguồn',ha='center',fontsize=7.5,color='#7F7F7F')
fig.tight_layout();fig.savefig('chart.png')
print('ok')
