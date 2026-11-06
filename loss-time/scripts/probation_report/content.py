B=[]
a=B.append
a(heading(1,'LỜI MỞ ĐẦU'))
a(para('Một tháng đầu tiên (từ 5/10 đến 6/11/2026) là khoảng thời gian tôi làm quen với công việc và từng bước nhận phần việc được giao. Qua thời gian này, tôi hiểu rõ hơn về quy trình sản xuất, về cách các bộ phận phối hợp xử lý hàng lỗi và về cách theo dõi máy móc, loss time tại nhà máy.'))
a(para('Báo cáo này tổng hợp các công việc được giao và mức hoàn thành, các công việc theo dõi tôi đã tham gia, công việc nhập liệu loss time, cùng với tự đánh giá những điểm còn chưa tốt và kế hoạch trong thời gian tới.'))
a(heading(1,'LỜI CẢM ƠN'))
a(para('Tôi xin chân thành cảm ơn anh Nguyễn Đình Luân cùng anh/chị và các thành viên trong team đã tận tình hướng dẫn, hỗ trợ và chia sẻ kinh nghiệm trong thời gian qua. Đây là nền tảng quý báu để tôi tiếp tục học hỏi và đóng góp tốt hơn cho công việc chung.'))
a(PB)
# ---------------- CHƯƠNG 1
a(heading(1,'CHƯƠNG 1. NỘI DUNG CÔNG VIỆC ĐƯỢC GIAO VÀ MỨC HOÀN THÀNH'))
a(heading(2,'1.1. Tổng quan mức hoàn thành'))
a(para('Tôi được giao 19 công việc chính. Sau một tháng, **10 việc đã hoàn thành (100%)**, **3 việc đang làm ở mức 50%** (chưa thành thạo hoặc cần thông qua người hướng dẫn) và **6 việc chưa học, chưa thực hiện**. Chi tiết như bảng 1.1.'))
jobs=[
('1','Báo cáo loss time cho RC Wood, RC Panel, Machining, Assembly, Inlay','100%'),
('2','Ghi nhận downtime và quản lý máy hư sửa giữa sản xuất và Bảo trì','100%'),
('3','Theo dõi thông tin và cập nhật bảo dưỡng máy','100%'),
('4','Theo dõi hàng chất lượng trả sửa giữa MAC và Assembly','100%'),
('5','Theo dõi hàng NCR trả về White Wood','100%'),
('6','Theo dõi chất lượng hàng lỗi báo cáo QC','Chưa học, chưa thực hiện'),
('7','Đề xuất phương án và phân bổ sửa hàng lỗi ASSY + NCR Finishing','50% (có ý tưởng, cần thông qua Mr. Luân)'),
('8','Ghi nhận vấn đề phát sinh lỗi và ý tưởng cải tiến bản vẽ quy trình ở bộ phận sản xuất','Chưa học, chưa thực hiện'),
('9','Làm NCR/MRR bổ sung vật tư cho sản xuất','50% (chưa thành thạo)'),
('10','Hỗ trợ quản lý đội ngũ làm rập/jig, kho chứa rập/jig','100%'),
('11','Hỗ trợ cung cấp MCP, mẫu veneer, mẫu gỗ, rập/jig cho lắp ráp; hỗ trợ kĩ thuật jig cho ASSY','100%'),
('12','Hỗ trợ làm jig và dữ liệu jig cho dự án JIG – runner của ME','Chưa học, chưa thực hiện'),
('13','Theo dõi và follow hàng bổ sung','100%'),
('14','Làm báo cáo và cập nhật báo cáo Pilot run','Chưa học, chưa thực hiện'),
('15','Hỗ trợ review Pilot run','100%'),
('16','In routing và triển khai bổ sung chi tiết hư hỏng nặng/thất lạc/lỗi kĩ thuật','50% (chưa thành thạo)'),
('17','Làm SOP, WI nội bộ','Chưa học, chưa thực hiện'),
('18','Điều phối kế hoạch tiến độ sửa hàng của sản xuất','100%'),
('19','Điều phối các phòng ban để đánh giá các vấn đề nghiêm trọng','Chưa học, chưa thực hiện'),
]
a(tcap('Bảng 1.1. Nội dung công việc được giao và mức hoàn thành'))
def shade(ri,ci,c):
    if ci!=2:return None
    return 'E2EFDA' if c.startswith('100') else ('FFF2CC' if c.startswith('50') else 'F2F2F2')
a(table([['STT','Công việc chính','Mức hoàn thành']]+[list(j) for j in jobs],[600,5671,2800],center=(0,),sz=20,shade=shade))
a(para('Phần lớn các việc đã hoàn thành thuộc ba nhóm: theo dõi hàng lỗi trả sửa (việc 4, 5, 18), ghi nhận downtime và bảo dưỡng máy (việc 2, 3), và báo cáo loss time (việc 1). Các nhóm này được trình bày ở chương 2 và chương 3.'))
a(PB)
# ---------------- CHƯƠNG 2
a(heading(1,'CHƯƠNG 2. CÁC CÔNG VIỆC THEO DÕI ĐÃ THAM GIA'))
a(heading(2,'2.1. Theo dõi hàng Assembly trả về MAC'))
a(para('**Mục đích:** theo dõi được tiến độ của hàng trả, đồng thời nắm được bộ phận nào phụ trách khắc phục lỗi, đảm bảo đúng tiến độ đề ra.'))
a(para('**Nội dung theo dõi:** mỗi dòng trong bảng báo cáo chất lượng ghi line, TE, WOM, tuần Assembly, mã hàng, tổng số lượng, số lượng set lỗi, vấn đề lỗi, bộ phận gây lỗi, bộ phận phụ trách, ngày trả về MAC, hình ảnh và ghi chú. Tôi liên hệ các bộ phận để nhận sửa, ghi thời gian dự kiến hoàn thành và cập nhật tình trạng cho đến khi hàng xong.'))
a(para('**Kết quả:** trong file theo dõi, hàng trả MAC từ 5/10 đến 15/10 có 39 dòng đã xử lý xong (348 set lỗi), trong đó 35 dòng do MAC xử lý và 4 dòng do ME. Thời gian từ lúc trả về đến lúc MAC xong trung bình 1,4 ngày; 36 trong 39 dòng (92%) xong trong 2 ngày. Tại ngày 16/10 còn 11 dòng đang xử lý, gồm 247 sản phẩm và 124 set lỗi (bảng 2.1).'))
a(tcap('Bảng 2.1. Hàng Assembly trả về MAC đang xử lý tại ngày 16/10'))
a(table([
 ['Mã hàng','Tổng SL','SL set lỗi','Vấn đề lỗi','Tình trạng'],
 ['AR18-1042-0892-R0000','33','33','Chi tiết A1S02-2 chưa dán veneer 2 cạnh','Hứa 11h30, 17/10'],
 ['RH7371F381TR','132','16','MHK dư số đo chiều dài 5–7 mm so với routing','Chờ Mr. Luân review'],
 ['SH974F1477','20','16','Chân tiện A5-S2 bị cắt dư số đo','Hết 17/10'],
 ['SH911AF1265','50','50','Chỉ A3S01-1 bị lệch khi ráp','Hết 17/10'],
 ['CZ065F191','10','7','Cánh cửa chạy rãnh lớn hơn routing','Done'],
 ['AR1363F1330-7066','2','2','Ván hông thiếu cắt góc 45 độ','Done'],
 ['Tổng','247','124','',''],
],[2200,900,1000,3271,1700],center=(1,2),sz=20))
a(heading(2,'2.2. Theo dõi hàng trả từ Finishing về White Wood'))
a(para('**Mục đích:** theo dõi và nắm được thông tin của hàng trả từ bộ phận Finishing (FIN) về White Wood (WW) để xử lý.'))
a(para('**Vì sao cần theo dõi sát:** FIN là công đoạn hoàn thành sản phẩm, nên hàng lỗi trả về cần được theo dõi sát và cập nhật liên tục tình trạng để đảm bảo hoàn thành đúng tiến độ.'))
a(para('**Nội dung theo dõi:** danh sách hàng NCR ghi WOF, mã hàng, chi tiết, số lượng, tuần giao, số NCR, nguyên nhân, bộ phận chịu trách nhiệm, giờ và ngày nhận hàng, cam kết hoàn thành của MAC + Panel và ASSY, và giờ giao lại cho FIN. Ở bảng ngày 5/11 có 4 dòng: 3 dòng tủ AR1368F1330-7066 (nứt kết cấu, rung lắc, bộ phận Technical, nhận ngày 27/10, làm mới) và 1 dòng hộc kéo RH12-4706 (vân veneer không khớp, Final Machining, nhận 18h19 ngày 5/11, MAC hứa hoàn thành lúc 10h ngày 7/11).'))
a(figure('rId21','ncr.png','Hình 2.1. Danh sách hàng NCR trả từ Finishing về White Wood (bảng ngày 5/11)'))
a(heading(2,'2.3. Ghi nhận downtime và theo dõi bảo trì máy hàng tuần'))
a(para('Kế hoạch bảo trì phòng ngừa (Maintenance Preventive) của nhà máy P2 được cập nhật hằng tuần. Với mỗi máy, bảng ghi khu vực, tên máy, mã máy, thời gian dừng máy kế hoạch, ngày thực hiện và giờ bắt đầu. Trong tuần W45 (từ 2/11 đến 8/11) có 19 máy cần bảo trì, gồm 9 máy Machining, 9 máy Panel và 1 máy Prototype, với tổng thời gian dừng máy kế hoạch là 75 giờ. Trong đó 5 máy có ngày cụ thể (2, 3, 4, 5 và 7/11), 3 máy làm ca đêm và 11 máy linh động trong tuần.'))
a(figure('rId22','bt.png','Hình 2.2. Kế hoạch bảo trì phòng ngừa P2 tuần W45 (2/11 – 8/11)'))
a(para('Máy hư ngoài kế hoạch được ghi nhận trong báo cáo Stop line + Repair WW, gồm ngày nộp, bộ phận, mã và tên máy, vấn đề, người phụ trách, giờ hư và giờ sửa xong, tổng giờ dừng máy. Ví dụ ngày 2/11 đến 4/11 có 9 dòng (máy dán cạnh, CNC 5 trục, máy thấm cạnh Veneer, máy cắt veneer, máy chà nhám băng, máy may veneer), phần lớn dừng dưới 1 giờ, dài nhất là 1,52 giờ (máy may veneer VES09, ngày 4/11).'))
a(figure('rId23','stop.png','Hình 2.3. Báo cáo Stop line + Repair WW (trích 2/11 – 4/11)'))
a(heading(2,'2.4. Các công việc phối hợp và hỗ trợ khác'))
a(para('Ngoài các việc theo dõi trên, tôi đã hoàn thành các việc phối hợp và hỗ trợ sau (đạt 100%):'))
a(bullet('Điều phối kế hoạch tiến độ sửa hàng của sản xuất: liên hệ các bộ phận nhận sửa và ghi thời gian dự kiến hoàn thành (việc 18).'))
a(bullet('Theo dõi và follow hàng bổ sung (việc 13).'))
a(bullet('Hỗ trợ quản lý đội ngũ làm rập/jig và kho chứa rập/jig (việc 10); hỗ trợ cung cấp MCP, mẫu veneer, mẫu gỗ, rập/jig cho lắp ráp và hỗ trợ kĩ thuật jig cho ASSY (việc 11).'))
a(bullet('Hỗ trợ review Pilot run (việc 15).'))
a(para('Ba việc đang ở mức 50% gồm đề xuất phương án và phân bổ sửa hàng lỗi ASSY + NCR Finishing (có ý tưởng, cần thông qua Mr. Luân), làm NCR/MRR bổ sung vật tư, và in routing kèm triển khai bổ sung chi tiết hư hỏng/thất lạc (chưa thành thạo).'))
a(PB)
# ---------------- CHƯƠNG 3
a(heading(1,'CHƯƠNG 3. LOSS TIME VÀ CÔNG VIỆC NHẬP LIỆU'))
a(heading(2,'3.1. Khái niệm loss time và quy tắc nhập liệu'))
a(para('Loss time là thời gian công nhân hoặc máy không tạo ra sản lượng. Mỗi dòng dữ liệu gồm ngày, mã, khu vực, bộ phận chịu trách nhiệm, lí do, giờ bắt đầu (From), giờ kết thúc (To) và số người (HC). Giờ công mất được tính theo công thức:'))
a('<w:p><w:pPr><w:spacing w:before="60" w:after="120"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>Giờ công mất = (To − From) × 24 × HC</w:t></w:r></w:p>')
a(para('Các quy tắc nhập liệu tôi đã nắm và áp dụng hằng ngày được tóm tắt trong bảng 3.1.'))
a(tcap('Bảng 3.1. Quy tắc nhập liệu theo từng mã'))
a(table([
 ['Mã','Quy tắc'],
 ['P20, P26','20 dòng mỗi ngày làm việc (5 dòng P20, 15 dòng P26). Thứ Hai đến thứ Sáu luôn tăng ca đến 20:00; thứ Bảy, Chủ nhật không tăng ca (7:30 đến 15:30).'],
 ['P22, A2','Ngày đầu nhân viên mới ghi P22 (8:00 đến 16:00); 12 ngày làm việc sau ghi A2 (7:30 đến 12:00).'],
 ['S32','Hỗ trợ khác từ sản xuất: lập bảng riêng, đối chiếu với báo cáo nguồn của từng khu vực.'],
 ['A10','Máy/dụng cụ hư: chỉ tính giờ trong khung 7:30 đến 20:00; số người tại máy do trưởng bộ phận xác nhận.'],
],[1300,7771],center=(0,),sz=22))
a(heading(2,'3.2. Công việc đã thực hiện và kết quả'))
for t in ['Lập bảng loss time và bảng nhập liệu cho từng ngày: 20 dòng cố định, P22 và A2 theo số nhân viên mới từng khu vực.',
 'Sắp xếp bảng nhập liệu theo từng bộ phận (Assembly, Machining, Raw Cutting) và đúng thứ tự ô trên form để nhập nhanh, ít sai.',
 'Lập bảng S32 riêng từ 11 nguồn báo cáo và đối chiếu số liệu với hệ thống.',
 'Đối chiếu dữ liệu đã nhập trên hệ thống với bảng chuẩn, lập danh sách chỗ cần sửa/thêm.',
 'Phân tích mã A10 (máy hư) và lập bản nháp để nhập khi có số người xác nhận.']:
    a(item(t,20))
a(para('Kết quả: 31 bảng loss time ngày (1.170 dòng) từ 5/10 đến 6/11, với 100 nhân viên mới; bảng S32 gồm 124 dòng, tổng 5.579,83 giờ công. Giờ công theo tuần (chỉ gồm P20, P26, P22, A2) như bảng 3.2.'))
a(tcap('Bảng 3.2. Giờ công loss time theo tuần'))
a(table([
 ['Tuần','Nhân viên mới (P22)','Giờ công P20+P26+P22+A2','Giờ công S32'],
 ['W41 (5 – 11/10)','40','3.536,0','880,8'],
 ['W42 (12 – 18/10)','16','3.317,0','1.662,8'],
 ['W43 (19 – 25/10)','21','3.078,5','1.782,8'],
 ['W44 (26/10 – 1/11)','16','2.719,0','1.233,1'],
 ['W45 (2 – 6/11)','7','2.050,5','20,3 (chưa đủ nguồn)'],
 ['Tổng','100','14.701,0','5.579,8'],
],[2300,1900,2571,2300],center=(1,2,3),sz=22))
a(heading(2,'3.3. Vấn đề dữ liệu và đề xuất'))
a(tcap('Bảng 3.3. Các vấn đề dữ liệu và đề xuất xử lý'))
a(table([
 ['Vấn đề','Bằng chứng','Đề xuất'],
 ['Thiếu báo cáo nguồn S32','UV từ 26/10, Veneer sau 29/10, MAC sau 28/10 chưa có','Xin báo cáo còn thiếu và cập nhật S32'],
 ['Chênh lệch bắn cát tuần W43','Hệ thống 75 giờ, báo cáo 24 giờ','Xác nhận với người phụ trách'],
 ['Số người A10 chưa xác nhận','14 trong 22 dòng A10 nháp chưa có số người','Trưởng bộ phận xác nhận số người tại máy'],
 ['Nhập sai trên hệ thống','Có ngày nhập sai (ví dụ giờ tăng ca thứ Hai), phải sửa và nhập lại','Cuối ngày đối chiếu số dòng và giờ công với bảng chuẩn'],
],[2300,3771,3000],sz=22))
# ---------------- CHƯƠNG 4
a(heading(1,'CHƯƠNG 4. TỰ ĐÁNH GIÁ VÀ KẾ HOẠCH TIẾP THEO'))
a(heading(2,'4.1. Những điều đã học nhưng chưa làm tốt'))
a(para('Bên cạnh những gì đã làm được, tôi nhận thấy còn các điểm chưa tốt sau:'))
a(bullet('**Công việc chưa học, chưa thực hiện (6 việc):** theo dõi chất lượng hàng lỗi báo cáo QC; ghi nhận vấn đề và cải tiến bản vẽ quy trình; hỗ trợ jig cho dự án JIG – runner của ME; báo cáo Pilot run; làm SOP, WI nội bộ; điều phối các phòng ban đánh giá vấn đề nghiêm trọng.'))
a(bullet('**Công việc chưa thành thạo (mức 50%):** làm NCR/MRR bổ sung vật tư; in routing và triển khai bổ sung chi tiết hư hỏng, thất lạc.'))
a(bullet('**Công việc cần thông qua:** đề xuất phương án và phân bổ sửa hàng lỗi ASSY + NCR Finishing cần được Mr. Luân thông qua trước khi triển khai.'))
a(heading(2,'4.2. Kế hoạch tiếp theo'))
for t in ['Duy trì theo dõi hàng trả sửa giữa Assembly, MAC và Finishing, cập nhật tiến độ hằng ngày.',
 'Theo dõi kế hoạch bảo trì và downtime máy hằng tuần.',
 'Học làm NCR/MRR bổ sung vật tư và in routing cho thành thạo để đạt 100%.',
 'Học dần các việc chưa thực hiện, ưu tiên báo cáo QC, SOP/WI và báo cáo Pilot run.',
 'Duy trì lập bảng loss time hằng ngày; cập nhật S32 và số người A10 khi có số liệu xác nhận.']:
    a(item(t,21))
a(heading(2,'4.3. Nội dung cần hỗ trợ'))
a(para('Tôi mong được anh/chị hướng dẫn các việc chưa học (báo cáo QC, SOP/WI, Pilot run), hỗ trợ thông qua phương án sửa hàng lỗi, cung cấp báo cáo nguồn S32 còn thiếu và xác nhận số người tại máy cho mã A10.'))
a(heading(1,'KẾT LUẬN'))
a(para('Sau một tháng, tôi đã làm quen và hoàn thành 10 trên 19 công việc được giao, đặc biệt là theo dõi hàng trả sửa giữa các bộ phận, ghi nhận downtime và bảo trì máy, cùng công việc nhập liệu loss time hằng ngày. Ba việc đang ở mức 50% và sáu việc còn lại tôi sẽ tiếp tục học và thực hiện trong thời gian tới.'))
a(para('Một lần nữa, tôi xin chân thành cảm ơn anh/chị và các thành viên trong team đã kiên nhẫn hướng dẫn, hỗ trợ tôi trong thời gian qua. Trong thời gian tới, tôi xin hứa sẽ:'))
a(bullet('Chủ động học các việc còn lại, hỏi ngay khi chưa rõ.'))
a(bullet('Cải thiện tốc độ, hoàn thành công việc đúng hạn và chính xác hơn.'))
a(bullet('Nắm vững quy trình hàng lỗi trả về để theo dõi và hỗ trợ các bộ phận tốt hơn.'))
BODY=''.join(B)
