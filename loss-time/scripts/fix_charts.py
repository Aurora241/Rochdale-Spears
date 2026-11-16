"""Sửa biểu đồ openpyxl để Excel hiện trục và nhãn số (nhãn đặt ở cấp biểu đồ, không đặt từng series)."""
from openpyxl.chart import BarChart,LineChart,PieChart
from openpyxl.chart.label import DataLabelList
def _labels():
    d=DataLabelList();d.showVal=True
    for k in ('showSerName','showCatName','showLegendKey','showPercent'):setattr(d,k,False)
    return d
def fix_charts(wb):
    for ws in wb.worksheets:
        for ch in ws._charts:
            if isinstance(ch,PieChart):
                ch.dataLabels=_labels();continue
            for c in getattr(ch,'_charts',[ch]):
                for ax in (c.x_axis,c.y_axis):ax.delete=False
                if isinstance(c,BarChart):c.dataLabels=_labels()
            if ch.legend:ch.legend.position='b'
