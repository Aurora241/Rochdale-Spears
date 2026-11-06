import re,shutil,os,zipfile
from xml.sax.saxutils import escape as X
SRC='/home/user/lt-work/tpl/x'
d=open(SRC+'/word/document.xml',encoding='utf8').read()
head=d[:d.find('<w:body>')+8]
# title + keep template's second paragraph style (find what follows title)
def P(style,text,extra=''):
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/>{extra}</w:pPr><w:r><w:t xml:space="preserve">{X(text)}</w:t></w:r></w:p>'
def H(t): return P('Heading2',t)
def FP(t): return P('FirstParagraph',t)
def BP(t): return P('BodyText',t)
def bullet(t,lead=None,num=1001):
    r=''
    if lead: r+=f'<w:r><w:rPr><w:bCs/><w:b/></w:rPr><w:t xml:space="preserve">{X(lead)}</w:t></w:r>'
    r+=f'<w:r><w:t xml:space="preserve">{X(t)}</w:t></w:r>'
    return f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr><w:pStyle w:val="Compact"/></w:pPr>{r}</w:p>'
def num(t): return bullet(t,None,1002)
def cell(t,hdr=False,w=None):
    rpr='<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman"/>'+('<w:b/><w:color w:val="FFFFFF"/>' if hdr else '')+'<w:sz w:val="21"/>'
    tcpr='<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="1F3A5F"/></w:tcPr>' if hdr else '<w:tcPr/>'
    return f'<w:tc>{tcpr}<w:p><w:pPr><w:pStyle w:val="Compact"/><w:spacing w:after="40"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{X(t)}</w:t></w:r></w:p></w:tc>'
def table(rows,widths):
    g=''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    out=('<w:tbl><w:tblPr><w:tblStyle w:val="Table"/><w:tblW w:type="pct" w:w="5000"/><w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0" w:val="0020"/><w:jc w:val="start"/><w:tblBorders>'
     +''.join(f'<w:{s} w:val="single" w:sz="4" w:color="A6A6A6"/>' for s in('top','left','bottom','right','insideH','insideV'))
     +f'</w:tblBorders></w:tblPr><w:tblGrid>{g}</w:tblGrid>')
    for i,r in enumerate(rows):
        out+='<w:tr>'+('<w:trPr><w:tblHeader w:val="true"/></w:trPr>' if i==0 else '')+''.join(cell(c,i==0) for c in r)+'</w:tr>'
    return out+'</w:tbl>'
cap="Giờ công mất mỗi tuần theo nhóm mã, W41–W45 (5/10 đến 6/11/2026). Nguồn: bảng loss time đã lập (cột Loss time labor), S32 từ 11 nguồn báo cáo"
fig=('<w:p><w:pPr><w:pStyle w:val="CaptionedFigure"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/></w:rPr><w:drawing><wp:inline><wp:extent cx="5334000" cy="2971800"/><wp:effectExtent b="0" l="0" r="0" t="0"/>'
 f'<wp:docPr descr="{X(cap)}" title="" id="25" name="Picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr descr="chart.png" id="26" name="Picture"/><pic:cNvPicPr><a:picLocks noChangeArrowheads="1" noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId24"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="5334000" cy="2971800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525"><a:noFill/><a:headEnd/><a:tailEnd/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
 +P('ImageCaption',cap))
B=[]
B.append(P('Title','Báo cáo thử việc'))
B.append(H('Mục tiêu và phạm vi'))
B.append(FP('Trong khoảng 1 tháng (5/10 đến 6/11/2026), tôi đã làm quen và đảm nhận phần lớn công việc trong bảng nội dung công việc được giao. Báo cáo này tóm tắt ba mảng chính tôi thực hiện, mức hoàn thành từng việc và những điều còn cần học thêm.'))
B.append(bullet('Báo cáo loss time (P20, P26, P22, A2, S32) và ghi nhận downtime máy hư (A10).',None))
B.append(bullet('Theo dõi hàng lỗi trả sửa giữa MAC và Assembly, hàng NCR trả về White Wood, điều phối thời gian sửa dự kiến.',None))
B.append(bullet('Theo dõi kế hoạch và cập nhật bảo dưỡng máy (bảo trì phòng ngừa P2).',None))
B.append(H('Nội dung công việc và mức hoàn thành'))
B.append(FP('Trong 19 việc được giao, tôi hoàn thành 10 việc (100%), đang làm 3 việc ở mức 50% và 6 việc chưa học, chưa thực hiện.'))
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
B.append(table([['STT','Công việc chính','Mức hoàn thành']]+[list(j) for j in jobs],[600,5320,2000]))
B.append(H('Theo dõi hàng lỗi trả sửa'))
B.append(FP('Tôi theo dõi hàng lỗi trả giữa Assembly, MAC và Finishing: liên hệ bộ phận nhận sửa, ghi thời gian dự kiến hoàn thành và cập nhật tình trạng cho đến khi hàng xong.'))
B.append(bullet('Trong file theo dõi, hàng trả MAC từ 5/10 đến 15/10 có 39 dòng đã xử lý xong (348 set lỗi): 35 dòng MAC xử lý, 4 dòng ME. Thời gian từ lúc trả về đến lúc MAC xong trung bình 1,4 ngày; 36 trong 39 dòng (92%) xong trong 2 ngày.','Hàng trả MAC: '))
B.append(bullet('Tại ngày 16/10 còn 11 dòng đang xử lý (247 sản phẩm, 124 set lỗi), gồm chi tiết chưa dán veneer 2 cạnh, MHK dư số đo chiều dài, chân tiện cắt dư số đo, chỉ bị lệch khi ráp, cánh cửa chạy rãnh lớn hơn routing, ván hông thiếu cắt góc 45 độ. Mỗi dòng đều ghi thời gian hứa hoàn thành (ví dụ 11h30 17/10, hết 17/10).','Đang xử lý: '))
B.append(bullet('Bảng ngày 5/11 có 4 dòng: 3 dòng tủ AR1368F1330-7066 (nứt kết cấu, rung lắc, bộ phận Technical, nhận 27/10, làm mới) và 1 dòng hộc kéo RH12-4706 (vân veneer không khớp, Final Machining, nhận 18h19 5/11, MAC hứa hoàn thành 10h ngày 7/11).','Hàng NCR từ Finishing về White Wood: '))
B.append(BP('Phần đề xuất phương án và phân bổ sửa hàng lỗi ASSY và NCR Finishing (mục 7) tôi mới đạt 50%: đã có ý tưởng nhưng cần thông qua Mr. Luân.'))
B.append(H('Loss time và downtime'))
B.append(FP('Giờ công mất = (To − From) × số người (HC). Tôi lập bảng loss time và bảng nhập liệu mỗi ngày theo các quy tắc đã chốt với quản lý.'))
B.append(table([
 ['Nội dung','Quy tắc'],
 ['P20, P26 (mã cố định)','20 dòng mỗi ngày làm việc. Thứ Hai đến thứ Sáu luôn tăng ca đến 20:00 (P20 47 giờ công, P26 225 giờ công mỗi ngày); thứ Bảy, Chủ nhật không tăng ca (7:30 đến 15:30). Chủ nhật hỏi lại khu vực nào làm việc.'],
 ['P22, A2 (nhân viên mới)','Ngày đầu ghi P22 (HR/ HSE, 8:00 đến 16:00); 12 ngày làm việc sau ghi A2 (HR/ HSE, 7:30 đến 12:00). Chủ nhật chỉ tính cho Assembly khi Assembly làm việc.'],
 ['S32 (hỗ trợ khác từ sản xuất)','Lập bảng riêng, đối chiếu với báo cáo nguồn (UV, Beam saw, Veneer, ép nóng, MAC, RC WOOD, INLAY, sơn, cào cước, bắn cát).'],
 ['A10 (máy/dụng cụ hư)','Bộ phận MAINTENANCE; chỉ tính giờ ngừng máy trong khung 7:30 đến 20:00. Số người tại máy phải do trưởng bộ phận xác nhận, tôi không tự ước lượng.'],
],[2400,5520]))
B.append(BP('Kết quả: 31 bảng loss time ngày với 1.170 dòng (5/10 đến 6/11); 100 nhân viên mới; S32 gồm 124 dòng, 5.579,83 giờ công từ 11 nguồn báo cáo.'))
B.append(table([
 ['Tuần','Nhân viên mới (P22)','Giờ công P20 + P26 + P22 + A2','Giờ công S32'],
 ['W41 (5 đến 11/10)','40','3.536,0','880,8'],
 ['W42 (12 đến 18/10)','16','3.317,0','1.662,8'],
 ['W43 (19 đến 25/10)','21','3.078,5','1.782,8'],
 ['W44 (26/10 đến 1/11)','16','2.719,0','1.233,1'],
 ['W45 (2 đến 6/11)','7','2.050,5','20,3 (chưa đủ nguồn)'],
 ['Tổng','100','14.701,0','5.579,8'],
],[2400,1700,2300,1520]))
B.append(fig)
B.append(bullet('P20 và P26 là mức nền (1.542 đến 1.574 giờ công mỗi tuần). Giờ công P22 và A2 giảm từ 1.994 xuống 690,5 giờ khi số người vào làm giảm.',None))
B.append(bullet('S32 lớn nhất ở W42 và W43 (khoảng 1.660 đến 1.780 giờ công) và phụ thuộc báo cáo từ nhiều nguồn.',None))
B.append(bullet('A10: đã lập bản nháp 22 dòng; số người tại máy có thể lệch 1 đến 2 người so với thực tế nên chỉ điền khi được xác nhận (hiện còn 14 dòng chờ).',None))
B.append(H('Bảo trì và downtime máy'))
B.append(FP('Tôi theo dõi kế hoạch bảo trì phòng ngừa P2 hằng tuần và ghi nhận downtime đã được sản xuất xác nhận.'))
B.append(bullet('Tuần W45 (2 đến 8/11) có 19 máy cần bảo trì (9 máy Machining, 9 máy Panel, 1 máy Prototype), tổng thời gian dừng máy kế hoạch 75 giờ. Trong đó 5 máy có ngày cụ thể (2, 3, 4, 5 và 7/11), 3 máy làm ca đêm, 11 máy linh động trong tuần.','Kế hoạch bảo trì: '))
B.append(bullet('Báo cáo Stop line + Repair WW ghi ngày nộp, máy, vấn đề, người phụ trách, giờ hư và giờ sửa xong. Ví dụ 2 đến 4/11 có 9 dòng (máy dán cạnh, CNC 5 trục, máy thấm cạnh Veneer, cắt veneer, chà nhám băng, máy may veneer), phần lớn dừng dưới 1 giờ, dài nhất 1,52 giờ (máy may veneer VES09, 4/11).','Ghi nhận downtime: '))
B.append(H('Vấn đề dữ liệu và đề xuất'))
B.append(table([
 ['Vấn đề','Bằng chứng','Đề xuất'],
 ['Thiếu báo cáo nguồn S32','UV từ 26/10, Veneer sau 29/10, MAC sau 28/10 chưa có; 8 dòng Beam saw chưa có số giờ','Xin các báo cáo còn thiếu và cập nhật bảng S32'],
 ['Chênh lệch bắn cát tuần W43','Hệ thống 75 giờ, báo cáo 24 giờ','Xác nhận lại với người phụ trách'],
 ['Số người A10 chưa xác nhận','14 trong 22 dòng A10 nháp chưa có số người','Trưởng bộ phận xác nhận số người tại máy'],
 ['Bộ phận “Unassigned”','Một số dòng S32 chưa gán bộ phận chịu trách nhiệm','Thống nhất cách gán bộ phận khi nhập'],
 ['Nhập sai trên hệ thống','Đã có ngày nhập sai (ví dụ giờ tăng ca thứ Hai), phải sửa và nhập lại','Cuối ngày đối chiếu số dòng và giờ công với bảng chuẩn'],
],[2200,3200,2520]))
B.append(H('Những điều chưa làm tốt'))
B.append(bullet('6 việc chưa học và chưa thực hiện: theo dõi chất lượng hàng lỗi báo cáo QC (6), ghi nhận vấn đề và cải tiến bản vẽ quy trình (8), hỗ trợ jig cho dự án JIG – runner của ME (12), báo cáo Pilot run (14), làm SOP, WI nội bộ (17), điều phối đánh giá vấn đề nghiêm trọng (19).','Chưa học: '))
B.append(bullet('Làm NCR/MRR bổ sung vật tư (9) và in routing, triển khai bổ sung chi tiết (16) đã làm nhưng chưa thành thạo, đạt 50%.','Chưa thành thạo: '))
B.append(bullet('Đề xuất phương án và phân bổ sửa hàng lỗi (7) cần thông qua Mr. Luân trước khi triển khai.','Chờ phê duyệt: '))
B.append(H('Kế hoạch tiếp theo'))
B.append(FP('Tôi sẽ giữ nhịp công việc hằng ngày và học nốt các việc chưa thực hiện.'))
for t in ['Lập và nhập loss time hằng ngày từ 7/11; hỏi lại thứ Bảy, Chủ nhật khu vực nào làm việc.',
 'Cập nhật S32 khi có báo cáo UV, Veneer, MAC; hoàn tất số người A10 sau khi được xác nhận.',
 'Học làm NCR/MRR và in routing cho thành thạo để đạt 100%.',
 'Học dần các việc chưa thực hiện, ưu tiên báo cáo QC, SOP/WI và báo cáo Pilot run.',
 'Duy trì theo dõi hàng trả sửa và kế hoạch bảo trì hằng tuần.']:
    B.append(bullet(t))
B.append(BP('Cần hỗ trợ: hướng dẫn các việc chưa học (QC, SOP/WI, Pilot run), báo cáo nguồn S32 còn thiếu và xác nhận số người A10.'))
sect='<w:sectPr><w:pgMar w:left="1247" w:right="1247" w:top="1134" w:bottom="1134"/></w:sectPr></w:body></w:document>'
out='/home/user/lt-work/probation/x'
shutil.rmtree(out,ignore_errors=True);shutil.copytree(SRC,out)
open(out+'/word/document.xml','w',encoding='utf8').write(head+''.join(B)+sect)
shutil.copy('chart.png',out+'/word/media/rId24.png')
fn='Bao_cao_thu_viec_Loss_time.docx'
if os.path.exists(fn):os.remove(fn)
with zipfile.ZipFile(fn,'w',zipfile.ZIP_DEFLATED) as z:
    z.write(out+'/[Content_Types].xml','[Content_Types].xml')
    for root,_,fs in os.walk(out):
        for f in fs:
            p=os.path.join(root,f);a=os.path.relpath(p,out)
            if a!='[Content_Types].xml':z.write(p,a)
print('built')
