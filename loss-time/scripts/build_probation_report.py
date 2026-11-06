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
B.append(P('Title','Báo cáo thử việc: Loss time và Downtime'))
B.append(H('Mục tiêu và phạm vi'))
B.append(FP('Trong khoảng 1 tháng đầu (5/10 đến 6/11/2026), tôi đã nắm hệ thống mã loss time, tự lập và nhập loss time hằng ngày cho xưởng Whitewood (Plant 2), đồng thời bắt đầu đảm nhận mã hỗ trợ sản xuất S32 và tìm hiểu mã lỗi máy A10.'))
B.append(bullet('31 bảng loss time hằng ngày với 1.170 dòng (5/10 đến 6/11), theo đúng thứ tự ô trên form nhập liệu.'))
B.append(bullet('Mã tôi phụ trách: mã cố định P20, P26; mã nhân viên mới P22, A2; mã hỗ trợ S32 (lập bảng riêng).'))
B.append(bullet('Mã A10 (lỗi máy): đã lập bản nháp để phân tích; số người tại máy chỉ điền khi trưởng bộ phận xác nhận, tôi không tự ước lượng.'))
B.append(H('Kiến thức và quy tắc đã nắm'))
B.append(FP('Giờ công mất = (To − From) × số người (HC). Tôi đã chốt với quản lý các quy tắc nhập sau và ghi lại thành tài liệu để áp dụng nhất quán mỗi ngày.'))
B.append(table([
 ['Nội dung','Quy tắc'],
 ['P20, P26 (mã cố định)','20 dòng mỗi ngày làm việc (5 dòng P20, 15 dòng P26). Thứ Hai đến thứ Sáu luôn tăng ca đến 20:00 (P20 47 giờ công, P26 225 giờ công mỗi ngày); chỉ thứ Bảy, Chủ nhật không tăng ca (7:30 đến 15:30). Chủ nhật hỏi lại khu vực nào làm việc.'],
 ['P22, A2 (nhân viên mới)','Ngày đầu ghi P22 (HR/ HSE, 8:00 đến 16:00); 12 ngày làm việc sau ghi A2 (HR/ HSE, 7:30 đến 12:00). Chủ nhật chỉ tính cho Assembly khi Assembly làm việc. Số người theo danh sách nhân viên mới hằng ngày, chỉ dùng số lượng theo khu vực.'],
 ['S32 (hỗ trợ khác từ sản xuất)','Luôn lập bảng riêng, đối chiếu với báo cáo nguồn (UV, Beam saw, Veneer, Ép nóng, MAC, RC WOOD, INLAY, sơn, cào cước, bắn cát).'],
 ['A10 (máy/dụng cụ hư)','Bộ phận MAINTENANCE; thời gian ngừng máy chỉ tính trong khung 7:30 đến 20:00 (12,5 giờ mỗi ngày). Số người tại máy phải do trưởng bộ phận xác nhận.'],
 ['Cách lập bảng nhập liệu','Tách theo bộ phận (Assembly, Machining, Raw Cutting), sắp theo thứ tự mã P20, P26, P22, A2, có cột “Đã nhập”. Trong tuần luôn gửi đầy đủ từ thứ Hai đến ngày hiện tại, chỉ bắt đầu file mới khi sang tuần mới.'],
],[2400,5520]))
B.append(H('Công việc đã thực hiện'))
B.append(FP('Tôi đã chuyển từ đọc hiểu file Loss time sang vận hành quy trình hằng ngày và mở rộng sang S32, A10.'))
for t in [
 'Lập bảng loss time và bảng nhập liệu mỗi ngày từ 5/10; ngày nào có nhân viên mới thì cập nhật P22 và A2 theo số người đã xác nhận.',
 'Kiểm tra độc lập số người A2 theo từng khu vực sau mỗi lần lập bảng (tính lại từ danh sách nhân viên mới).',
 'Lập bảng S32 riêng: 124 dòng, 5.579,83 giờ công, tổng hợp từ 11 nguồn báo cáo (5/10 đến 4/11); chỉ lệch 0,23 giờ ở một nguồn (MAC 23/10).',
 'Đối chiếu dữ liệu xuất từ hệ thống với bảng của tôi, lập danh sách chỗ cần sửa/thêm/xóa và bảng nhập liệu cho phần đó.',
 'Phân tích mã A10: lập bản nháp 22 dòng, tính giờ ngừng máy nhiều ngày theo khung 7:30 đến 20:00, và lập bảng số công nhân tại máy từ dữ liệu quá khứ để trưởng bộ phận xác nhận.',
 'Ghi toàn bộ quy tắc đã chốt vào tài liệu dự án để áp dụng nhất quán; dùng công cụ AI hỗ trợ tính toán và luôn kiểm tra lại số liệu trước khi nhập.',
]: B.append(num(t))
B.append(table([
 ['Tuần','Nhân viên mới (P22)','Giờ công P20 + P26 + P22 + A2','Giờ công S32'],
 ['W41 (5 đến 11/10)','40','3.536,0','880,8'],
 ['W42 (12 đến 18/10)','16','3.317,0','1.662,8'],
 ['W43 (19 đến 25/10)','21','3.078,5','1.782,8'],
 ['W44 (26/10 đến 1/11)','16','2.719,0','1.233,1'],
 ['W45 (2 đến 6/11)','7','2.050,5','20,3 (chưa đủ nguồn)'],
 ['Tổng','100','14.701,0','5.579,8'],
],[2400,1700,2300,1520]))
B.append(BP('Cột giờ công chỉ gồm P20, P26, P22 và A2. S32 tính riêng từ 11 nguồn báo cáo; tuần W45 mới có nguồn đến 4/11 và còn thiếu báo cáo.'))
B.append(H('Số liệu và phát hiện chính'))
B.append(FP('Trong 5 tuần, mã cố định và mã nhân viên mới tạo ra 14.701 giờ công; giờ công của nhân viên mới (P22, A2) giảm dần từ 1.994 xuống 690,5 giờ khi số người vào làm giảm.'))
B.append(fig)
B.append(bullet('P20 và P26 là mức nền (1.542 đến 1.574 giờ công mỗi tuần), gần như không đổi.',None))
B.append(bullet('P22 và A2 biến động theo từng đợt nhân viên mới: 100 người vào làm trong 5 tuần, đông nhất tuần W41 (40 người).',None))
B.append(bullet('S32 lớn nhất ở W42 và W43 (khoảng 1.660 đến 1.780 giờ công), đáng theo dõi vì phụ thuộc báo cáo từ nhiều nguồn.',None))
B.append(bullet('A10: số người tại máy ảnh hưởng trực tiếp giờ công mất nên cần xác nhận thực tế; bảng số người ước theo máy khớp lịch sử 48% chính xác và 89% lệch tối đa 1 người, chỉ dùng để tham khảo.',None))
B.append(H('Vấn đề dữ liệu và đề xuất'))
B.append(FP('Khi đối chiếu với dữ liệu hệ thống và các báo cáo nguồn, tôi ghi nhận những chỗ còn mở sau đây.'))
B.append(table([
 ['Vấn đề','Bằng chứng','Đề xuất'],
 ['Thiếu báo cáo nguồn S32','UV từ 26/10, Veneer sau 29/10, MAC sau 28/10 chưa có; 8 dòng Beam saw chưa có số giờ','Xin các báo cáo còn thiếu và cập nhật bảng S32'],
 ['Chênh lệch bắn cát tuần W43','Hệ thống 75 giờ, báo cáo 24 giờ','Xác nhận lại với người phụ trách'],
 ['Số người A10 chưa xác nhận','14 trong 22 dòng A10 nháp chưa có số người','Trưởng bộ phận xác nhận số người tại máy; tôi không điền ước lượng'],
 ['Bộ phận “Unassigned” trong dữ liệu','Một số dòng S32 chưa gán bộ phận chịu trách nhiệm','Thống nhất cách gán bộ phận khi nhập'],
 ['Nhập sai trên hệ thống','Đã có các ngày nhập sai (ví dụ giờ tăng ca thứ Hai), phải sửa và nhập lại','Cuối ngày đối chiếu số dòng và giờ công với bảng chuẩn'],
],[2200,3200,2520]))
B.append(H('Đối chiếu nội dung công việc được giao'))
B.append(FP('Theo bảng nội dung công việc trong biên bản thử việc, tôi tự đánh giá mức hoàn thành như sau. Các mục [Bổ sung] cần tôi điền thêm bằng chứng trước khi nộp.'))
B.append(table([
 ['Nhóm mục','Mức hoàn thành tự đánh giá','Bằng chứng / ghi chú'],
 ['10 mục','100%','Loss time hằng ngày, S32, đối chiếu hệ thống, A10 nháp (phần liên quan loss time); [Bổ sung] bằng chứng cho các mục còn lại'],
 ['3 mục','50%','Sửa hàng lỗi ASSY + NCR Finishing (chờ Mr. Luân); NCR/MRR bổ sung vật tư; in routing'],
 ['6 mục','[Bổ sung]','Chưa ghi mức hoàn thành / chưa phát sinh'],
],[2000,2200,3720]))
B.append(H('Kế hoạch tiếp theo'))
B.append(FP('Tôi sẽ giữ nhịp lập bảng hằng ngày và hoàn thiện các phần còn mở với bộ phận liên quan.'))
for t in ['Lập và nhập loss time hằng ngày từ 7/11; hỏi lại thứ Bảy/Chủ nhật khu vực nào làm việc.',
 'Cập nhật S32 khi có báo cáo UV, Veneer, MAC và các dòng thiếu giờ.',
 'Hoàn tất số người cho các dòng A10 sau khi trưởng bộ phận xác nhận, rồi nhập vào hệ thống.',
 'Sửa các chỗ lệch với dữ liệu hệ thống theo danh sách đã lập.',
 'Chuyển dần sang báo cáo tuần bằng email (phân tích giờ công theo mã, khu vực).']:
    B.append(bullet(t))
B.append(BP('Cần hỗ trợ: báo cáo nguồn S32 còn thiếu, xác nhận số người A10 và danh sách nhân viên mới sau 4/11.'))
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
