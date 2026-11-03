import openpyxl,datetime as dt
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
F='Times New Roman';t=Side(style='thin');bd=Border(left=t,right=t,top=t,bottom=t)
red=PatternFill('solid',fgColor='C00000');hf=Font(name=F,bold=True,color='FFFFFF');ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
D=lambda d:dt.datetime(2026,10,d)
wb=openpyxl.Workbook();wb.remove(wb.active)
def sheet(name,title,headers,rows,widths):
    ws=wb.create_sheet(name);ws['A1']=title;ws['A1'].font=Font(name=F,bold=True,size=13)
    for i,h in enumerate(headers,1):
        c=ws.cell(3,i,h);c.fill=red;c.font=hf;c.alignment=ctr;c.border=bd
    for k,row in enumerate(rows):
        for i,v in enumerate(list(row)+[None],1):
            x=ws.cell(4+k,i,v);x.font=Font(name=F);x.border=bd;x.alignment=Alignment(horizontal='left' if isinstance(v,str) and len(v)>18 else 'center',vertical='center',wrap_text=True)
            if isinstance(v,dt.datetime): x.number_format='dd/mm/yyyy'
    for i,w in enumerate(widths): ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=w
    return ws
H=['Bộ phận','Ngày','Việc','Nhận ra dòng bằng','Ô cần sửa','Đang có','Sửa thành','Giờ công thay đổi','Đã làm']
rows=[
('INLAY',D(26),'SỬA','Hai dòng giống nhau "25 full sét" AT655 (7:30-12:30, 5 người, 25 giờ). Sửa MỘT trong hai dòng','Notes và To','"...25 full sét", To 12:30','"...25 sét đuôi + hồng 3 đầu", To 11:30',-5),
('Machining',D(28),'SỬA','Hỗ trợ sửa hàng MAC - chạy lại tay nắm (RH12-3791, 15:30-19:30, 3 người)','Process','Raw Cutting','Machining',0),
('Raw Cutting',D(26),'ĐIỀN','Ép hàng sửa (chi tiết A5-S02-1) AT655F535, 7:00-7:30, 4 người','Dept In Charge','(trống)','OTHERS (INVESTIGATE) - đề xuất, chờ xác nhận',0),
('Raw Cutting',D(29),'ĐIỀN','Làm lại (bắn cát) hàng bong tróc sơn Resin, AT11-1008-0535-R0000, 7:30-19:30, 4 người','Dept In Charge','(trống)','Chờ xác nhận bộ phận (nguồn ghi "làm lại")',0),
('Raw Cutting',D(30),'ĐIỀN','Làm lại (bắn cát) AT11-1008-0535-R0000, 7:30-19:30, 4 người','Dept In Charge','(trống)','Chờ xác nhận bộ phận',0),
('Raw Cutting',D(31),'ĐIỀN','3 dòng làm lại AT25-1008-0535, 0888, 1089 (7:30-10:30, 4 người)','Dept In Charge','(trống)','Chờ xác nhận bộ phận',0),
]
sheet('Sửa RC - Machining - INLAY','DÒNG S32 ĐÃ NHẬP CẦN SỬA (đối chiếu file xuất 3/11, 25/10-31/10)',H,rows,[14,12,9,70,18,26,36,14,9])
rows2=[
('Assembly',D(31),'XÓA','Thợ sơn lòng trong tủ, 7:30-15:30, 12 người, 96 giờ (đang có thêm dòng 7:30-19:30, 144 giờ)','xóa 1 dòng','','',-96),
('Assembly',D(26),'NHẬP THÊM','Thợ sơn lòng trong tủ 26/10 đến 30/10: mỗi ngày 7:30-19:30, 12 người, 144 giờ','','chưa có','5 dòng (720 giờ)',720),
('Assembly',dt.datetime(2026,11,1),'NHẬP THÊM','Thợ sơn Chủ nhật 1/11: 7:30-15:30, 12 người, 96 giờ (cần xác nhận Chủ nhật có làm)','','chưa có','1 dòng (96 giờ)',96),
]
sheet('Assembly (bạn tự nhập)','S32 ASSEMBLY (chỉ để bạn tham khảo, bạn tự nhập khi cần)',H,rows2,[14,12,12,80,14,14,18,14,9])
ws=wb.create_sheet('UV chưa có nguồn');ws['A1']='UV LINE TUẦN 44 (26/10-1/11): CHƯA CÓ BÁO CÁO UV';ws['A1'].font=Font(name=F,bold=True,size=13)
for k,tx in enumerate(['Hệ thống chưa có dòng UV nào từ 26/10 đến 31/10, và file UV mình có chỉ đến 24/10 (25/10 Chủ nhật 0 giờ).','Cần xin file "lost time khu vực UV" cập nhật đến 1/11 (số HC làm việc và HC tăng ca mỗi ngày).','Mỗi ngày 2 dòng: ca thường 7:30-15:30 (HC làm việc) và tăng ca 15:30-19:30 (HC OT), Process Raw Cutting, bộ phận FINISHING, Notes "Không có bomtime chuyền UV".','Tuần 43 (19/10-24/10) còn UV thiếu 19/10, 20/10, 21/10 ca thường và dòng 12 người thừa 24/10: xem file Nhap_lieu_con_lai_theo_bo_phan_27-10.xlsx.']):
    ws.cell(3+k,1,tx).font=Font(name=F)
ws.column_dimensions['A'].width=150
wb.save('Nhap_lieu_S32_con_lai_theo_bo_phan_03-11.xlsx')
