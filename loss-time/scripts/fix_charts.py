"""Sửa biểu đồ openpyxl để Excel hiện trục, tên cột và nhãn số."""
from openpyxl.chart import BarChart,LineChart,PieChart
from openpyxl.chart.label import DataLabelList
def fix_charts(wb):
    for ws in wb.worksheets:
        for ch in ws._charts:
            for c in getattr(ch,'_charts',[ch]):
                if isinstance(c,PieChart):
                    c.dataLabels=DataLabelList();c.dataLabels.showVal=True;c.dataLabels.showPercent=False
                    for k in ('showSerName','showCatName','showLegendKey'):setattr(c.dataLabels,k,False)
                    continue
                for ax in (c.x_axis,c.y_axis):ax.delete=False
                for s in c.series:
                    s.dLbls=DataLabelList();s.dLbls.showVal=True
                    for k in ('showSerName','showCatName','showLegendKey','showPercent'):setattr(s.dLbls,k,False)
                    s.dLbls.numFmt='0%' if isinstance(c,LineChart) else '#,##0.0';s.dLbls.position='t' if isinstance(c,LineChart) else 'outEnd'
            ch.legend.position='b' if ch.legend else None
