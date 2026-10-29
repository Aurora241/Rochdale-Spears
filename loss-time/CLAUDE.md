# Quy tắc làm việc: bảng Loss time White Wood

Tài liệu này tóm tắt những gì đã chốt trong phiên làm việc với người dùng (nhân viên văn phòng, 2 tuần đầu đi làm). Đọc hết trước khi làm. Trả lời bằng tiếng Việt, ngắn gọn.

## 1. Mục tiêu hằng ngày
Người dùng yêu cầu: "lập bảng loss time ngày dd/mm". Bạn tạo:
1. Bảng loss time của ngày đó (file Excel) theo cấu trúc sheet "Total loss time".
2. Bảng nhập liệu xếp theo thứ tự ô trên form (mục 7).
3. Với mỗi lần người dùng đưa danh sách nhân viên mới: tạo dòng P22 cho ngày vào làm VÀ danh sách A2 cho 12 ngày làm việc tiếp theo (không được quên A2), kèm ngày kết thúc của nhóm.

Phạm vi: mã cố định (P20, P26), mã nhân viên mới (P22, A2) và, từ tuần 42, mã S32 theo báo cáo của Assembly (mục 11). Mã P24 (Kaizen) chỉ thêm khi người dùng xác nhận lịch. BỎ QUA mã lỗi máy móc (A10 và liên quan) và các mã khác (O15, P24...) vì người dùng chưa được hướng dẫn.

## 2. Dữ liệu và file
- File gốc: "Loss time White Wood 2026" (sheet "Total loss time", cột: Week, Date, Item's, Code Máy Hư, Q'ty, Số NCR/ME, Khu vực phát sinh, Code, Name code, Chart/bộ phận chịu trách nhiệm, Lí Do, Ghi chú, From, To, LossTime (Hr), HC, Loss time labor (Hrs)). Week = "W"&WEEKNUM(Date,21). LossTime = (To-From)*24. Loss time labor = LossTime*HC.
- Sheet "Loss time categoty" là bảng 35 mã (A1-A11, O11-O19, P20-P26, Q25-Q30, S31-S32).
- Dữ liệu trong file chỉ có đến 3/10/2026. Các ngày sau do bạn lập theo quy tắc. Bảng 5/10 đến 17/10 đã được lập (xem mục 9).
- Danh sách nhân viên mới (có tên, mã nhân viên, lương) là dữ liệu cá nhân, người dùng nói chỉ để tham khảo. KHÔNG đưa tên, mã, lương vào bất kỳ file nào, không commit lên git, chỉ dùng ngày vào làm, bộ phận, số người để đếm.

## 3. Mã cố định hằng ngày (20 dòng ở đầu mỗi ngày làm việc: 5 P20 + 15 P26)
Mẫu lấy từ thứ Sáu 2/10/2026. Giờ trong bảng là giờ ngày thường (có tăng ca đến 20:00 từ 28/9):

| # | Code | Process (khu vực) | Dept In Charge | Notes (lí do) | Ghi chú cũ | From | To | HC |
|---|---|---|---|---|---|---|---|---|
| 1 | P20 | RC - WOOD | RC WOOD | Vệ sinh khu vực máy móc RC WOOD |  | 07:00 | 07:30 | 7 |
| 2 | P20 | RC - PANEL | RC - PANEL | Vệ sinh khu vực máy móc RC PANEL |  | 07:00 | 07:30 | 18 |
| 3 | P20 | Machining | Machining | Vệ sinh khu vực máy móc MAC FINAL |  | 07:00 | 07:30 | 7 |
| 4 | P20 | Machining | Machining | Vệ sinh khu vực máy móc PRE MAC |  | 07:00 | 07:30 | 12 |
| 5 | P20 | Assembly | Assembly | Quét nhà toàn xưởng khu vực ASSY |  | 07:30 | 20:00 | 2 |
| 6 | P26 | RC - PANEL | RC - PANEL | Shop Floor data RC PANEL | Recorder | 07:30 | 20:00 | 4 |
| 7 | P26 | Machining | Machining | Shop Floor data- STOCK ROOM | Recorder | 07:30 | 20:00 | 1 |
| 8 | P26 | Machining | Machining | Shop Floor data MAC FINAL | Recorder | 07:30 | 20:00 | 1 |
| 9 | P26 | RC - PANEL | RC - PANEL | Theo dõi Hàng venner làm xong | Recorder | 07:30 | 20:00 | 1 |
| 10 | P26 | RC - PANEL | RC - PANEL | Nhận venner cho ép nóng | Supporter | 07:30 | 20:00 | 1 |
| 11 | P26 | RC - WOOD | RC WOOD | Lựa hàng tinh giao mac- gỗ đẹp xấu | Supporter | 07:30 | 20:00 | 1 |
| 12 | P26 | RC - PANEL | RC - PANEL | Ghi nhận hàng giao kho | Recorder | 07:30 | 20:00 | 1 |
| 13 | P26 | RC - PANEL | RC - PANEL | Kiểm hàng chất lượng vào kho | Supporter | 07:30 | 20:00 | 1 |
| 14 | P26 | RC - PANEL | Machining | Nhận hàng gỗ đầu vào từ RC | Supporter | 07:30 | 20:00 | 1 |
| 15 | P26 | Machining | Machining | Ghi nhận hàng giao kho | Recorder | 07:30 | 20:00 | 1 |
| 16 | P26 | Machining | Machining | quét nhà toàn xưởng khu vực ASSY | Supporter | 07:30 | 20:00 | 1 |
| 17 | P26 | Machining | Machining | Theo hàng pilot -hàng test-spo-subcon-NCR | Recorder | 07:30 | 20:00 | 1 |
| 18 | P26 | Machining | Machining | Giao nhận hàng hộc kéo | Supporter | 07:30 | 20:00 | 1 |
| 19 | P26 | Assembly | Assembly | Báo output- sửa máy- giao nhận vật tư | Recorder | 07:30 | 20:00 | 1 |
| 20 | P26 | Assembly | Assembly | Giao nhận HW- phát HW cho Line sx | Supporter | 07:30 | 20:00 | 1 |

Lưu ý về mẫu: từ 23/9/2026 dòng P26 Machining "Nhận ván - venner ráp cụm & cảo bịch" (Supporter) được thay bằng "quét nhà toàn xưởng khu vực ASSY" (Supporter); trước đó dùng dòng cũ. Dòng "Nhận hàng gỗ đầu vào từ RC" ghi khu vực RC - PANEL, bộ phận Machining từ 14/3. Khi lập bảng mới luôn kiểm tra mẫu 20 dòng của ngày làm việc gần nhất trong file, vì nội dung có thể đổi.

Quy ước giờ ngày thường của bộ phận sản xuất (RC, Machining, Assembly): ca chính thức 7:30-16:00; các dòng P26 và dòng P20 Assembly ghi 7:30-20:00 từ 28/9 (trước đó 7:00-19:30). Người dùng là nhân viên văn phòng (8:00-16:30), bỏ qua giờ đó.

Thứ Bảy: số dòng giống ngày thường (20 dòng), nhưng bộ phận sản xuất về sớm nửa tiếng so với ngày thường: không tăng ca = 7:30-15:30 cho P26 và P20 Assembly. Thứ Bảy tăng ca (về 20:00) là hiếm (ví dụ 3/10). 4 dòng P20 ngắn mặc định 7:00-7:30 (thứ Bảy 3/10 từng ghi 7:30-8:00, cần xác nhận). Hỏi người dùng thứ Bảy đó có tăng ca không.

Chủ nhật: thường nghỉ (từ tháng 9 hầu như không có dữ liệu). Chỉ lập khi người dùng nói có làm. Chủ nhật không tính vào 12 ngày A2.

## 4. Nhân viên mới: P22 và A2
- Ngày đầu: ghi P22 "HR Training New Workers", 8:00-16:00 (8 giờ), Dept In Charge HR/HSE, Notes "Đào tạo công nhân mới", một dòng mỗi khu vực với HC = số người. KHÔNG ghi thêm A2 cùng ngày (nhóm 4/3 từng bị ghi trùng).
- 12 ngày làm việc sau đó (thứ Hai đến thứ Bảy, bỏ Chủ nhật): ghi A2 "Training" 7:30-12:00 (giờ gần nhất; trước 28/9 là 7:00-12:00), Dept In Charge HR/HSE, Notes "HC mới- 50% hiệu suất- 2 tuần-dd/mm" (dd/mm = ngày vào làm). Ghi chú cột "A2 ngày k/12 - nhóm dd/mm" để theo dõi.
- Người ghi "0" ở cột CT/TV của danh sách vẫn tính vào HC (đã chốt tạm). NGOẠI LỆ: tài xế xe nâng (Grade "IDL-Forklift", lao động gián tiếp) KHÔNG tính vào P22/A2 (người dùng đã chốt; 12/10 có 1 người Machining như vậy bị loại).
- Tên khu vực dùng: RC - WOOD, RC - PANEL, Machining, Assembly, Raw Cutting Panel (cách ghi gần nhất cho Raw Cutting). File cũ còn có "RC".
- Mỗi khi có danh sách mới: thêm dòng vào sheet "Lịch A2 nhân viên mới" (đủ 12 ngày A2 của nhóm, dùng WORKDAY.INTL(ngày,1,11)) và lập sheet danh sách theo nhóm (ngày vào làm, khu vực, HC, mã từng ngày P22 hoặc A2 ngày k, ngày cuối A2 bằng WORKDAY.INTL(ngày, 12, 11)).

## 5. Quy luật mã P24 (chỉ tham khảo, chưa đưa vào bảng)
Kaizen thường thứ Hai và thứ Ba hằng tuần, 8:00-16:30, mỗi khu vực một dòng 1-4 người, Dept HR/HSE, Name "Planned Losstime". Gần nhất: 21-22/9 và 28-29/9: Machining 3, Assembly 2, RC - PANEL 2. Hỏi người dùng ngày nào có Kaizen và số người.

## 6. Việc cần hỏi người dùng khi có dữ liệu mới
- Danh sách nhân viên mới của ngày cần lập (hoặc "không có").
- Thứ Bảy có tăng ca không; Chủ nhật có làm không; Kaizen (P24).

## 7. Thứ tự cột nhập form (từ trái sang phải)
LossTime Code, Notes, Process, Dept In Charge, From, To, Worker Qty. (Loss Date, LossTime Name, TotalLoss Hour do form tự điền.) Ghép cột: Notes = cột "Lí Do", Process = "Khu vực phát sinh", Dept In Charge = "Chart/bộ phận chịu trách nhiệm". Cột Ghi chú cũ (Recorder/Supporter, số ngày A2) để tham khảo, không nhập.

## 8. Quy định còn mở (cần hỏi production engineer)
- A2 kéo dài 12 hay 13 ngày làm việc (trong file: 12 ngày 43%, 13 ngày 28%).
- Giờ A2: 7:00-12:00 hay 7:30-12:00; ghi chú "50% hiệu suất" nhưng 4,5-5 giờ trên 8 giờ là 56-62,5%.
- Phân loại cho mã P26 trong bảng mã (chưa có nhóm loss, 33,7% giờ công).
- Số A2 các nhóm cuối tháng 9 trong file có thể ghi thiếu ngày.

## 9. Tình trạng đã lập (đến 17/10/2026)
Đã lập bảng 5/10 đến 29/10 (trừ Chủ nhật 11/10 và 18/10). Ngày 17/10 (Thứ Bảy) không tăng ca; 19/10 đến 23/10 ghi theo mẫu ngày thường (24/10 Thứ Bảy mặc định không tăng ca) có tăng ca (cần xác nhận); nhân viên mới 21/10: Raw Cutting 1, Machining 2, Assembly 2 (1 người dưới Mr.Dũng-Inlay), A2 từ 22/10; nhân viên mới 22/10: Raw Cutting 3, Machining 1, Assembly 2 (A2 từ 23/10); 23/10 chưa có nhân viên mới (chờ báo). Nhân viên mới 19/10: Assembly 2, Machining 1; 20/10: Assembly 4 (2 người dưới Mr.Dũng-Inlay), Machining 2 (theo danh sách admin ngày 22/10). Nhân viên mới: 9/10 Assembly 6 + Raw Cutting 2 (lấy theo P22 trên hệ thống, A2 từ 10/10), 10/10 chưa thấy có, 16/10 Assembly 1 (theo hệ thống). Đối chiếu với file xuất hệ thống ngày 20/10: A2 của nhóm 9/10 chưa được nhập từ 10/10; 10/10 nhóm 8/10 bị nhập Raw Cutting 4 thay vì Raw Cutting 1 + Machining 3; 19/10 nhóm 8/10 Machining nhập 1 thay vì 3.

Nhân viên mới đã ghi nhận (chỉ số người theo khu vực):
| Ngày vào làm | Khu vực | Số người |
|---|---|---|
| 5/10 | Assembly | 6 |
| 6/10 | Assembly | 7 |
| 6/10 | Raw Cutting Panel | 4 |
| 6/10 | Machining | 3 |
| 7/10 | Assembly | 6 |
| 7/10 | Raw Cutting Panel | 2 |
| 8/10 | Raw Cutting Panel | 1 |
| 8/10 | Machining | 3 |
| 9/10, 10/10 | chờ danh sách | |
| 12/10 | Assembly | 2 |
| 12/10 | Raw Cutting Panel | 1 |
| 13/10 | Assembly | 2 |
| 13/10 | Raw Cutting Panel | 4 |
| 13/10 | Machining | 1 |
| 14/10 | Assembly | 1 |
| 15/10 | Assembly | 4 |
| 16/10 | chưa rõ khu vực | 1 |
| 17/10 | chờ người dùng báo | ? |

Các nhóm A2 cũ còn chạy (lấy từ file Loss time đến 3/10; "n" = số ngày A2 đã ghi trong file, HC = số người ngày ghi cuối). Ngày A2 hôm nay = n + số ngày làm việc (T2-T7) từ 4/10 đến hôm nay:
| Nhóm | Khu vực | n | HC |
|---|---|---|---|
| 18/9 | Assembly | 4 | 7 |
| 23/9 | Assembly | 9 | 3 |
| 24/9 | Assembly | 4 | 5 |
| 25/9 | Assembly | 4 | 1 |
| 28/9 | Assembly | 5 | 3 |
| 21/9 | Machining | 4 | 2 |
| 22/9 | Machining | 4 | 4 |
| 23/9 | Machining | 9 | 9 |
| 24/9 | Machining | 4 | 3 |
| 25/9 | Machining | 4 | 1 |
| 22/9 | RC | 4 | 1 |
| 23/9 | RC | 4 | 1 |
| 24/9 | RC | 4 | 2 |
| 26/9 | RC | 4 | 1 |
| 23/9 | Raw Cutting Panel | 3 | 3 |
| 25/9 | Raw Cutting Panel | 3 | 1 |
Nhóm 19/9 (Assembly 13 ngày, Machining 12 ngày) đã xong A2 trước 5/10. Nhóm 23/9 Assembly và Machining xong 7/10. Nhóm 28/9 Assembly xong 12/10; các nhóm còn lại (k=11 hôm 12/10) kết thúc 13/10, nên từ 14/10 danh sách A2 chỉ còn các nhóm vào làm từ 5/10 trở đi và Raw Cutting Panel 23/9, 25/9.

## 10. Chuẩn chất lượng khi xuất file
- Font Times New Roman; công thức Excel thật (không gõ cứng kết quả): Week, LossTime, Loss time labor, tổng hợp bằng SUMIFS.
- Tính lại công thức (LibreOffice), không có lỗi công thức, đối chiếu tổng với tính tay độc lập, và nêu rõ các chỗ chưa chắc.
- Nói rõ phần nào là suy ra theo quy tắc, phần nào lấy trực tiếp từ file.

## 11. Mã S32 hằng tuần (Assembly hỗ trợ bộ phận khác) - việc mới nhận từ tuần 42
Người dùng mới được bàn giao việc này (chưa quen cách hoạt động), nên giải thích ngắn gọn và nêu các điểm cần xác nhận.
- S32 = Other Support of Production (nhóm Support): người của một khu vực làm cho bộ phận khác hoặc nhà thầu. Có trong 37/38 tuần của file (418 dòng, 19,5% giờ công, lớn thứ ba sau P26 và A2). Không lặp cố định mỗi ngày; số dòng, HC và giờ thay đổi theo từng tuần, cần báo cáo của người gửi.
- Khu vực Assembly không có S32 từ tuần 13/7 đến 28/9; tuần 21/9 và 28/9 gần như không có S32 (27 và 12 giờ). Tuần 38-41 người dùng nói đã có người nhập, KHÔNG cần bù.
- Tiền lệ cách ghi: thợ sơn lòng trong tủ (Assembly, bộ phận Finishing, lí do "Không có bomtime sơn lòng trong"), tuần W23 ghi từng ngày HC gộp, thứ Hai-thứ Sáu 7:30-19:30, thứ Bảy và Chủ nhật 7:30-15:30; nửa đầu năm có kiểu "Tổng kết tuần" một dòng 00:00-23:59. Cào cước giống các dòng "Inlay hỗ trợ MAC" (Assembly, bộ phận Machining). Bắn cát dùng bộ phận Subcon (file viết "Subcon", báo cáo viết "Supcon").
- Yêu cầu tuần 42 (do đồng nghiệp nhờ qua tin nhắn): thợ sơn lòng trong tủ 12 người (12 giờ/ngày T2-T6, 8 giờ T7, CN nghỉ, tổng 816 giờ, hỗ trợ Fin); cào cước 16/10 (2 dòng, 1 người mỗi dòng, 15:00-20:00 và 8:00-12:00, mã hàng RH893 và RH5061, hỗ trợ MAC, tổng 9 giờ); bắn cát Supcon Đại Xuyên 15/10 (4 người, mã hàng RH2252, số lượng 734 chân, báo cáo ghi 7:00-16:00 nhưng 8 giờ/người, tổng 32 giờ). Tổng 857 giờ công.
- Đã lập bảng S32 tuần 42 (9 dòng, 857 giờ, khớp báo cáo), có sheet đối chiếu, nhập liệu theo form và giải thích. Các điểm còn chờ xác nhận: bắn cát ghi 7:00-15:00 để ra 8 giờ (7:00-16:00 sẽ thành 36 giờ); cào cước 15:00-20:00 ngày 16/10 và thợ sơn thứ Bảy 17/10 có thể là giờ dự kiến; giờ thợ sơn 7:30-19:30 hay 8:00-20:00; mã S32 và bộ phận chịu trách nhiệm cần production engineer xác nhận.
- Form nhập chỉ có ô Notes (dùng cho lí do). Mã hàng và số lượng: chưa biết ô nào trên form tương ứng (có các ô Pa, WO, Suppl, Desc), người dùng sẽ xem form.
- Mỗi tuần khi người dùng gửi báo cáo Assy mới (thợ sơn, cào cước, bắn cát, hỗ trợ khác), làm tương tự: lập bảng S32, đối chiếu với tổng giờ trong báo cáo, và hỏi giờ thực tế/dự kiến.

### Quy tắc đã chốt cho S32 (cập nhật 19/10/2026)
- LUÔN làm S32 RIÊNG: một file/bảng S32 riêng, không trộn S32 vào bảng hằng ngày P20/P26/P22/A2 (và không ghi S32 vào bảng ngày). Người dùng sẽ gửi cập nhật, phải lập lại bảng S32 theo dữ liệu mới nhất.
- Nguồn S32 là gói zip "Lost time S32" (từ 5/10 đến nay) cộng báo cáo dạng hình: Beam saw (6 bản, mỗi bản là bản chụp tích lũy, lấy bản đầy đủ nhất), UV (báo cáo ngày + sheet phát sinh), Ép nóng (ÉP HÀNG SỬA), Veneer hàng BS (sheet "tổng"), và 3 hình của Assembly (thợ sơn lòng trong tủ tuần 42, cào cước 16/10, bắn cát 15/10). Luôn dùng bản mới nhất và chỉ lấy dòng từ 5/10.
- Cách ghi: Code S32; Process thường RC - PANEL cho các file khu vực (UV, Ép nóng, Beam saw, Veneer) và Assembly cho 3 hình; Dept In Charge lấy từ cột Charge to/Hỗ trợ (P1, TF, ME, Finishing, Machining, Subcon; "Machining CNC" ghi "Machining"; "QA..." ghi "QC"; "Unassigned" phải hỏi). Nhiều mã hàng cùng một ca = một dòng, mã hàng nối bằng "+", không nhân giờ theo từng mã. UV báo cáo ngày: giờ = 8 x HC làm việc + 4 x HC tăng ca, tách 2 dòng/ngày (7:30-15:30 và 15:30-19:30) để đúng giờ báo cáo.
- Luôn có sheet đối chiếu giờ công với số giờ trong nguồn (phải "Khớp") và sheet Chi tiết mã hàng.
- Điểm còn chờ xác nhận: UV Chủ nhật 11/10; cách ghi 1 dòng hay 2 dòng/ngày; bộ phận của Beam saw "Unassigned" 16/10; Beam saw và Veneer đã xác nhận là khu vực Raw Cutting (RC - PANEL) ngày 20/10; giờ dự kiến (thợ sơn thứ Bảy 17/10, cào cước 15:00-20:00 ngày 16/10); bắn cát ghi 7:00-15:00 để ra 8 giờ.
- Trạng thái: đã lập bảng S32 từ 5/10 đến 19/10: 42 dòng, 1.761,93 giờ công, khớp nguồn.


## Quy tắc nhập vào hệ thống (đối chiếu 20/10/2026)
- Tên trong dropdown hệ thống: Process = Raw Cutting / Machining / Assembly / INLAY (không có "RC - PANEL"); Dept In Charge = MACHINING, RAW CUTTING, ASSEMBLY, HR/ HSE, FINISHING (có khoảng trắng cuối), TECHNICAL FINISHING, OTHERS (INVESTIGATE), ME, MAINTENANCE, QA/QC/TESTING, Prototype, SUBCON. File nhập liệu đã đổi sang các tên này.
- P22 LUÔN charge HR/ HSE. A2 hiện charge HR/ HSE (Tiến nhập từ 21/9; Hùng trước đó ghi theo khu vực).
- Lỗi đã gặp: nhập nhầm mã (P26 thành P20 hoặc P22), nhập trùng cả ngày (9/10, 13/10). Mỗi tuần cộng P26 và P20 để đối chiếu: tuần đủ 6 ngày có tăng ca P26 1.350, P20 282; tuần có Thứ Bảy không tăng ca P26 1.269, P20 273. Nếu tuần lệch phải tìm dòng sai.
- Beam saw và Veneer thuộc Raw Cutting. Cách ghi S32 của người trước: hỗ trợ P1 ép nóng, dept OTHERS (INVESTIGATE); veneer/bảng màu/làm mẫu, dept TECHNICAL FINISHING; bắn cát, dept SUBCON (process Assembly); UV, dept FINISHING, process Raw Cutting.

## Cập nhật S32 ngày 21/10/2026
- Bảng S32 hiện đến 20/10: `loss-time/outputs/Loss_time_S32_05-10_den_20-10.xlsx` (47 dòng, 1.766,33 giờ, mọi nguồn khớp). Nguồn mới: Beam saw bản 20.10new (14 dòng từ 8/10) và Veneer bản 20/10 (thêm 3 dòng ngày 20/10, 4 giờ).
- Điểm cần xác nhận: Veneer walnut cắt ghép 20/10 Charge to "bs lại" (bộ phận chưa rõ); Beam saw "Unassigned" 16/10 và 20/10; 3 dòng Beam saw 19/10 trên hệ thống đang nhập ngày 18/10.
- UV (báo cáo ngày đến 11/10, phát sinh đến 13/10) và Ép nóng (đến 6/10) chưa có số liệu mới; bắn cát 70 giờ chưa được giải thích.
- Tên bộ phận và khu vực trong bảng S32 đã đổi theo dropdown hệ thống (xem mục "Quy tắc nhập vào hệ thống").

## Cập nhật S32 ngày 24/10/2026
- Bảng S32 hiện đến 24/10: `loss-time/outputs/Loss_time_S32_05-10_den_29-10.xlsx` (97 dòng, 4.343,63 giờ). Nguồn mới (gói update.zip): Beam saw bản (2) (thêm 22/10 và 24/10), Veneer bản (2) (thêm 2 dòng 24/10, sau đó bản Veneer mới nhất đổi dòng sửa hàng hộc kéo 48 giờ từ 24/10 sang 23/10 và thêm dòng 24/10 2 người 8:00-15:00, 14 giờ), và file mới "hàng sửa khu vực MAC Mr.Hiệp" (14/10 và 23/10, khu vực Machining, bộ phận ME).
- Khi gửi nhiều phiên bản của cùng một file Beam saw/Veneer, dùng bản đầy đủ nhất (bản có nhiều dòng nhất và chứa mọi dòng của các bản cũ; đã kiểm tra bản cũ là tập con).
- Chưa đưa vào bảng: 7 dòng Beam saw 22/10 "HÀNG BẢNG MÀU" không có giờ bắt đầu/kết thúc. Dòng MAC 23/10 ghi tổng 1:55 nhưng 9:32-11:20 là 1:48 (đối chiếu lệch 0,23 giờ). Beam saw "Unassigned" nay có 5 dòng (16/10, 20/10, 24/10).

## Cập nhật UV ngày 24/10/2026
- File UV mới (báo cáo ngày đến 24/10) thêm 1.504 giờ công cho 12/10-24/10 (UV tổng 2.280 giờ). File có các khối tuần W42 lặp hai lần, giống hệt nhau: chỉ tính một lần theo ngày. Chủ nhật 18/10 và 25/10 không có HC làm việc nên không ghi. Cách ghi UV giữ như cũ: mỗi ngày 2 dòng (ca thường 7:30-15:30 = HC làm việc, tăng ca 15:30-19:30 = HC OT), kể cả Thứ Bảy theo số HC OT của báo cáo. Thợ sơn, cào cước, bắn cát, ép nóng chưa có số liệu mới.

## Cập nhật S32 ngày 26/10/2026 (3 báo cáo Assembly tuần 43)
- Thợ sơn lòng trong tủ tuần 43: 12 người, 912 giờ (T2-T6 12 giờ, T7 8 giờ, CN 8 giờ: báo cáo tuần này có CN làm). Cào cước: 20/10 RH109F1 2 người (báo cáo ghi 7:30-20:00 nhưng 24 giờ, bảng ghi 7:30-19:30), 24/10 RH133F1 16 giờ, 25/10 (CN) RH133F1 6 giờ; tổng 46. Bắn cát 24/10: 3 người 8:00-16:00, 4 mã hàng (732 chân), 24 giờ.
- Bảng S32 hiện đến 25/10: `loss-time/outputs/Loss_time_S32_05-10_den_29-10.xlsx` (97 dòng, 4.343,63 giờ). Các dòng Chủ nhật 25/10 (thợ sơn 96 giờ, cào cước 6 giờ) tô cam, chờ xác nhận có nhập CN không. Bắn cát tổng 32 (15/10) + 24 (24/10) = 56 giờ, vẫn chưa khớp con số 70 giờ bên kia báo.

## Chủ nhật và phản hồi đồng nghiệp ngày 26/10/2026
- Chủ nhật KHÔNG phải lúc nào cũng nghỉ. Trong file có 10 Chủ nhật có đủ 20 dòng P20/P26 (và A2) của mọi khu vực (17/5, 24/5, 7/6, 14/6, 21/6, 28/6, 19/7, 2/8, 23/8, riêng 27/9 chỉ có A2 Assembly 3 người). Đồng nghiệp xác nhận 26/10: "lắp ráp làm cả chủ nhật". Tuần 43 báo cáo Assembly ghi CN 25/10 làm (thợ sơn 8 giờ, cào cước 3 giờ); 11/10 có sự cố A10 Assembly (12 người) vào Chủ nhật. Khi lập bảng một Chủ nhật phải HỎI người dùng khu vực nào có làm việc, không mặc định nghỉ.
- Đồng nghiệp kiểm tra pivot tuần 43: Assembly P20 thiếu 24/10 và 25/10 (mỗi ngày 25 giờ); P26 Assembly chỉ có 125 giờ (T2-T6, mỗi ngày 2 dòng x 12,5); Machining P26 375 giờ (T2-T6) cũng thiếu ngày. Lỗi nhập 20/10: dòng P20 Assembly "Quét nhà toàn xưởng khu vực ASSY" bị ghi bộ phận MACHINING, phải là ASSEMBLY.
- Không điền giờ không có thật để chỉnh k-factor; chỉ nhập giờ có báo cáo hoặc có người xác nhận.

## Danh sách sửa/thêm tuần 43 (26/10/2026)
- `loss-time/outputs/Danh_sach_sua_them_19-10_den_26-10.xlsx` (script `build_fix_list.py`): đối chiếu file xuất hệ thống 19-24/10 với bảng chuẩn. Gồm nhập thêm (19/10 P22, 20/10 A2, 21/10 P22, cả ngày 24/10), sửa dòng sai (20/10 bộ phận P20 Assembly; 21/10 P22 HC và Process P20; 23/10 bộ phận P26 "Nhận hàng gỗ"), xóa dòng trùng (21/10 A2 nhóm 9/10), và S32 khu vực khác (UV thiếu 19-21/10, xóa dòng UV thừa 24/10, đổi ngày Beam saw 18/10 thành 19/10, điền bộ phận trống). S32 của Assembly (thợ sơn, cào cước, bắn cát) do người dùng nhập khi cần.
- Lỗi lặp lại cần soát mỗi tuần: nhập trùng cả nhóm A2, sai Process/Dept của P20/P26, thiếu P22/A2 của nhóm vào làm.

## Quy tắc Chủ nhật chỉ Assembly làm (26/10/2026)
- Bảng 25/10: chỉ Assembly làm, không tăng ca. Chỉ ghi dòng cố định của Assembly: P20 "Quét nhà toàn xưởng khu vực ASSY" (2 người) và P26 Recorder + Supporter (mỗi dòng 1 người), giờ 7:30-15:30 (như Thứ Bảy không tăng ca) = 32 giờ công. Machining và Raw Cutting không có dòng cố định ngày đó. Nhân viên mới của Assembly: P22 8:00-16:00 nếu có; A2 của các nhóm Assembly vẫn ghi 7:30-12:00 và TÍNH là một ngày A2 (tiền lệ 27/9: nhóm 17/9 và 19/9 của Assembly có A2 Chủ nhật 27/9 và đủ 13 ngày). A2 của Machining/Raw Cutting không ghi và không tính Chủ nhật đó.
- Hệ quả: từ 26/10 trở đi, số ngày A2 của các nhóm Assembly phải cộng thêm 1 (Chủ nhật 25/10 đã tính). Khi lập bảng tiếp, sửa hàm đếm ngày cho Assembly. Chủ nhật 11/10 và 18/10: cần hỏi Assembly có làm không; nếu có thì phải bổ sung dòng cố định và tính lại A2 Assembly.
- Bảng 25/10: 12 dòng, 122 giờ công (P20 16, P26 16, A2 90 = 20 người Assembly x 4,5).

## Cập nhật 27/10/2026
- Nhân viên mới 23/10: Raw Cutting 1 (người ghi "0" ở cột CT/TV vẫn tính). Bảng 23/10 thêm P22 8 giờ, 24/10 thêm A2 4,5 giờ. Thứ Hai 26/10 ghi theo mẫu ngày thường có tăng ca (P20 47, P26 225; lần đầu mình ghi nhầm không tăng ca 15:30 rồi 16:00, người dùng sửa: Thứ Hai-Thứ Sáu LUÔN tăng ca); bảng 26/10: 429,5 giờ công. Từ 26/10, A2 của các nhóm Assembly vào làm trước 25/10 cộng thêm 1 ngày (Chủ nhật 25/10 đã tính); nhóm Machining/Raw Cutting không cộng. Đã kiểm tra lại độc lập số người A2 mỗi khu vực từ 15/10 đến 26/10, khớp bảng.
- Danh sách sửa/thêm mở rộng đến 26/10 (96 dòng nhập thêm gồm cả 24/10, 25/10, 26/10).
- Quy tắc tăng ca đã chốt (27/10): Thứ Hai-Thứ Sáu LUÔN tăng ca, giữ nguyên mẫu P20 và P26 đến 20:00 (P20 47, P26 225 mỗi ngày). Chỉ Thứ Bảy và Chủ nhật không tăng ca (P20 và P26 Assembly 7:30-15:30; P20 38, P26 144 khi cả xưởng làm). Không tự đoán ngày thường không tăng ca.

## Định dạng bảng nhập liệu từ 27/10/2026 (yêu cầu của người dùng)
- Phần mềm nhập liệu tự sắp xếp lại dòng sau mỗi lần nhập nên người dùng dễ nhầm. Vì vậy từ nay MỌI bảng nhập liệu phải lập RIÊNG TỪNG BỘ PHẬN (sheet Assembly, Machining, Raw Cutting; theo cột Process), trong mỗi sheet xếp theo ngày, trong mỗi ngày xếp cố định theo mã P20, P26, P22, A2 (rồi theo giờ bắt đầu). Có dòng tiêu đề từng ngày (số dòng và tổng giờ) và cột "Đã nhập" để người dùng đánh dấu X.
- Script: `scripts/build_nhap_lieu_theo_bo_phan.py` (chuyển file `Nhap_lieu_05-10_den_dd-mm.xlsx` thành bảng theo bộ phận: `python build_nhap_lieu_theo_bo_phan.py <file nhập liệu> dd-mm dd-mm <file ra>`), và `build_fix_list_theo_bo_phan.py` cho danh sách sửa/thêm. File mẫu: `outputs/Nhap_lieu_theo_bo_phan_19-10_den_26-10.xlsx` (96 dòng nhập thêm, 7 dòng sửa/xóa, 12 dòng S32 Raw Cutting).
- Khi lập bảng ngày mới, luôn gửi bản theo bộ phận này (có thể gửi thêm bản cũ theo ngày để đối chiếu).

## A10 từ báo cáo bảo trì (27/10/2026)
- Người dùng gửi báo cáo máy dừng của bảo trì (Mr.Phong - Mr.Khiêm) cho 19/10-25/10: 9 sự cố, cột: ngày, khu vực (Machining/RC), mã máy, tên máy, hiện tượng, loại "Machine stop", giờ bắt đầu/kết thúc, giờ dừng (2 cột giống nhau), một cột số (24 cho CNCR02, 12 cho các dòng còn lại; CHƯA RÕ nghĩa, nghi là số công nhân bị ảnh hưởng), trạng thái Done.
- Cách ghi vào form: Code A10, Notes "MãMáy-Tên máy-Hiện tượng", Process Machining hoặc Raw Cutting (RC), Dept In Charge MAINTENANCE, From/To theo báo cáo. Worker Qty KHÔNG được điền từ cột số cuối (12/24) của báo cáo và KHÔNG được ước lượng: A10 rất nhạy cảm, chỉ ghi số công nhân thật đứng máy do tổ trưởng khu vực xác nhận (ghi tên người xác nhận). File `outputs/Nhap_lieu_A10_19-10_den_25-10_NHAP.xlsx` để trống Worker Qty (9 sự cố, tổng 11,95 giờ máy dừng). Đừng bao giờ tự điền số công nhân cho A10.

## Đối chiếu lần 3 (file xuất 27/10, 19/10-25/10)
- Người dùng đã sửa gần hết: 19/10, 20/10, 22/10, 23/10 khớp bảng chuẩn; 21/10 còn lỗi nhập A2 nhóm 9/10 (Assembly 6 bị lặp, thiếu Raw Cutting 2); 24/10 thiếu P26 "Ghi nhận hàng giao kho" Machining, thiếu A2 nhóm 23/10 Raw Cutting, và 4 dòng Machining ghi nhầm bộ phận RAW CUTTING; 25/10 sai Worker Qty A2 nhóm 13/10 (4 thay vì 2), 15/10 (2 thay vì 4), 19/10 (1 thay vì 2).
- S32 Raw Cutting vẫn còn 12 việc (UV 19-21/10, xóa dòng UV thừa 24/10, đổi ngày Beam saw, điền bộ phận). File: `outputs/Nhap_lieu_con_lai_theo_bo_phan_27-10.xlsx`. Lỗi lặp lại: nhầm Process/Dept giữa các khu vực, nhầm số người giữa các nhóm A2 cùng ngày, nhập trùng.

## Phân tích A10 (27/10/2026)
- File `outputs/Phan_tich_A10.xlsx` (scripts `analyze_a10_data.py`, `analyze_a10_report.py`, chạy trong thư mục có `sys.pkl`): 191 sự cố (186 dòng file + 5 dòng hệ thống 4-11/10), 1.402 giờ công ghi, 917 giờ sau khi loại 9 dòng bất thường (máy dừng cả ngày 00:00-23:59 và dòng trùng). CNC router chiếm 37% giờ công A10; CNCR01 (17 sự cố, 216 giờ công), CNCR03, CNCR14 là máy cần chú ý; băng chuyền Assembly ít lần nhưng giờ công lớn (11-12 người); 37% dòng chỉ ghi chung chung nên không rõ nguyên nhân; bảo trì tuần 43 báo 9 sự cố (11,95 giờ máy dừng) mà hệ thống chưa nhập.
- A10 rất nhạy cảm: chỉ nhập số công nhân thật đứng máy do tổ trưởng xác nhận, không ước lượng.

## Cập nhật 28/10/2026
- Nhân viên mới 26/10: Assembly 2 (P22 16 giờ, A2 từ 27/10). Đã lập 26/10 (445,5 giờ), 27/10 (425 giờ) và 28/10 (398 giờ), cả ba theo mẫu ngày thường có tăng ca; 27/10 và 28/10 chưa có nhân viên mới (chờ báo). Số người A2 mỗi khu vực đã kiểm tra độc lập, khớp.
- Bảng nhập liệu theo bộ phận cho 26-28/10: `outputs/Nhap_lieu_theo_bo_phan_26-10_den_28-10.xlsx`. Danh sách sửa/thêm cũ đến 26/10 (`Danh_sach_sua_them_19-10_den_26-10.xlsx`) đã lỗi thời ở ngày 26/10 (chưa có P22 Assembly 2); dùng `Nhap_lieu_con_lai_theo_bo_phan_27-10.xlsx` và file theo bộ phận mới.

## Cập nhật 29/10/2026
- Nhân viên mới 27/10: Assembly 4 (1 người dưới Mr.Dũng-Inlay), Raw Cutting 1 (người ghi "0" ở cột CT/TV vẫn tính); P22 40 giờ, A2 từ 28/10. Đã lập 27/10 (465 giờ), 28/10 (420,5 giờ), 29/10 (402,5 giờ), mẫu ngày thường có tăng ca; 28/10 và 29/10 chưa có nhân viên mới (chờ báo). Số người A2 mỗi khu vực đã kiểm tra độc lập, khớp. File theo bộ phận: `outputs/Nhap_lieu_theo_bo_phan_27-10_den_29-10.xlsx`; file 26/10-28/10 cũ đã thay (26/10 vẫn nằm trong `Nhap_lieu_theo_bo_phan_26-10_den_28-10.xlsx`, 27/10 trở đi dùng file mới vì 27/10 đã thêm P22).

## Cập nhật S32 ngày 29/10/2026 (update_1.zip)
- Gói mới: Beam saw bản (2) (30 dòng, thêm 24/10, 28/10, 29/10: 3,3 giờ công), Ép nóng (thêm dòng 26/10 AT655F535, 4 người 7:00-7:30, Charge to Unassigned, 2 giờ công), hàng sửa MAC (thêm 28/10 RH12-3791 chạy lại tay nắm, 3 người, 15:30-20:00 nhưng tổng 4:00 nên bảng ghi 15:30-19:30 = 12 giờ công). Bảng S32 hiện đến 29/10: 97 dòng, 4.343,63 giờ công; mọi nguồn khớp trừ dòng MAC 23/10 lệch 0,23 giờ.
- Bảng INLAY (RESIN) gửi qua ảnh (4 dòng AT654, AT655 x2, AT384, 5 người mỗi dòng, 25 + 20 + 25 + 14 = 84 giờ công) KHÔNG có ngày và giờ: chưa nhập, phải hỏi ngày và giờ bắt đầu/kết thúc. Process INLAY có trong hệ thống; dòng "Hỗ trợ hàng inlay p1" cũ dùng bộ phận OTHERS (INVESTIGATE).
