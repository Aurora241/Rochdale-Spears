"""Slide bảng 19 công việc (đảm nhận, hỗ trợ, tiến độ) cùng phong cách bộ slide thử việc (navy + teal). Chạy: python build_slide_cong_viec.py <file pptx ra>"""
import sys
from pptx import Presentation
from pptx.util import Inches,Pt,Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN,MSO_ANCHOR
NAVY=RGBColor(0x14,0x28,0x4B);TEAL=RGBColor(0x12,0xA1,0x8F);BG=RGBColor(0xF4,0xF6,0xFA);INK=RGBColor(0x3B,0x45,0x5C);WHITE=RGBColor(255,255,255)
GRN=RGBColor(0xE1,0xF3,0xEF);YEL=RGBColor(0xFD,0xF1,0xD0);GRY=RGBColor(0xEA,0xED,0xF3);LINE=RGBColor(0xD6,0xDB,0xE6)
FONT='Be Vietnam Pro'
jobs=[
('Báo cáo loss time cho RC Wood + RC Panel + Machining + Assembly + Inlay','Tiến','100%'),
('Ghi nhận downtime và quản lý máy hư sửa giữa sản xuất và Bảo trì','Tiến','100%'),
('Theo dõi thông tin và cập nhật bảo dưỡng máy','Tiến','100%'),
('Theo dõi hàng chất lượng trả sửa giữa MAC và Assembly','Luân','100%'),
('Theo dõi hàng NCR trả về White Wood','Luân','100%'),
('Theo dõi chất lượng hàng lỗi báo cáo QC','Luân','Chưa học, chưa thực hiện'),
('Đề xuất phương án và phân bổ sửa hàng lỗi ASSY + NCR Finishing','Luân','50% (có ý tưởng, phải thông qua Mr. Luân)'),
('Ghi nhận vấn đề phát sinh lỗi, ý tưởng cải tiến bản vẽ quy trình ở bộ phận sản xuất','Luân','Chưa học, chưa thực hiện'),
('Làm NCR/MRR bổ sung vật tư cho sản xuất','Tú','50% (chưa thành thạo)'),
('Hỗ trợ quản lý đội ngũ làm rập/jig, kho chứa rập/jig','Cơ','100%'),
('Hỗ trợ cung cấp MCP, mẫu veneer, mẫu gỗ, rập/jig cho lắp ráp; hỗ trợ kĩ thuật jig cho ASSY','Cơ','100%'),
('Hỗ trợ làm jig và dữ liệu jig cho dự án JIG – runner của ME','Cơ','Chưa học, chưa thực hiện'),
('Theo dõi và follow hàng bổ sung','Tú','100%'),
('Làm báo cáo và cập nhật báo cáo Pilot run','Cơ','Chưa học, chưa thực hiện'),
('Hỗ trợ review Pilot run','Cơ','100%'),
('In routing và triển khai bổ sung chi tiết hư hỏng nặng/thất lạc/lỗi kĩ thuật','Tú','50% (chưa thành thạo)'),
('Làm SOP, WI nội bộ','Tiến','Chưa học, chưa thực hiện'),
('Điều phối kế hoạch tiến độ sửa hàng của sản xuất','Luân','100%'),
('Điều phối các phòng ban để đánh giá các vấn đề nghiêm trọng','Luân','Chưa học, chưa thực hiện'),
]
prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5)
s=prs.slides.add_slide(prs.slide_layouts[6])
bg=s.background.fill;bg.solid();bg.fore_color.rgb=BG
def tb(x,y,w,h,text,size,bold=False,color=INK,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP):
    sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));tf=sh.text_frame;tf.word_wrap=True;tf.vertical_anchor=anchor
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
    p=tf.paragraphs[0];p.alignment=align;r=p.add_run();r.text=text;r.font.size=Pt(size);r.font.bold=bold;r.font.color.rgb=color;r.font.name=FONT
    return sh
bar=s.shapes.add_shape(1,Inches(0.67),Inches(0.42),Inches(0.5),Inches(0.05));bar.fill.solid();bar.fill.fore_color.rgb=TEAL;bar.line.fill.background()
tb(0.67,0.58,12,0.55,'19 việc được giao: 10 việc hoàn thành, 3 việc đạt 50%',26,True,NAVY)
# chú giải
lx=9.0
for c,t in ((GRN,'100%'),(YEL,'50%'),(GRY,'Chưa học')):
    r=s.shapes.add_shape(1,Inches(lx),Inches(1.27),Inches(0.16),Inches(0.16));r.fill.solid();r.fill.fore_color.rgb=c;r.line.color.rgb=LINE
    tb(lx+0.22,1.24,1.1,0.22,t,10,False,INK);lx+=1.35
rows=len(jobs)+1
cw=[0.45,6.75,1.15,1.15,2.5]
tbl=s.shapes.add_table(rows,5,Inches(0.67),Inches(1.55),Inches(sum(cw)),Inches(0.27*rows)).table
for i,w in enumerate(cw):tbl.columns[i].width=Inches(w)
tbl.first_row=True
def cell(r,c,text,size=10.5,bold=False,color=INK,fill=None,align=PP_ALIGN.LEFT):
    ce=tbl.cell(r,c);ce.margin_left=ce.margin_right=Inches(0.08);ce.margin_top=ce.margin_bottom=Inches(0.02);ce.vertical_anchor=MSO_ANCHOR.MIDDLE
    tf=ce.text_frame;tf.word_wrap=True;p=tf.paragraphs[0];p.alignment=align;p.text=''
    ru=p.add_run();ru.text=text;ru.font.size=Pt(size);ru.font.bold=bold;ru.font.color.rgb=color;ru.font.name=FONT
    ce.fill.solid();ce.fill.fore_color.rgb=fill if fill else WHITE
for c,h in enumerate(['STT','Công việc chính','Đảm nhận','Hỗ trợ','Tiến độ']):
    cell(0,c,h,10.5,True,WHITE,NAVY,PP_ALIGN.CENTER if c!=1 else PP_ALIGN.LEFT)
tbl.rows[0].height=Inches(0.32)
for i,(job,sup,pr) in enumerate(jobs,1):
    st=GRN if pr.startswith('100') else (YEL if pr.startswith('50') else GRY)
    stc=RGBColor(0x0B,0x6B,0x5E) if pr.startswith('100') else (RGBColor(0x8A,0x5A,0x00) if pr.startswith('50') else RGBColor(0x5A,0x64,0x7C))
    zebra=WHITE if i%2 else RGBColor(0xFA,0xFB,0xFD)
    cell(i,0,str(i),10,False,INK,zebra,PP_ALIGN.CENTER)
    cell(i,1,job,10.5,False,NAVY,zebra)
    cell(i,2,'Anh',10.5,False,INK,zebra,PP_ALIGN.CENTER)
    cell(i,3,sup,10.5,False,INK,zebra,PP_ALIGN.CENTER)
    cell(i,4,pr,10,True,stc,st,PP_ALIGN.CENTER)
    tbl.rows[i].height=Inches(0.255)
tb(0.67,7.03,9,0.25,'Cột "Hỗ trợ" là người hỗ trợ đảm nhận công việc. Nguồn: bảng nội dung công việc trong biên bản thử việc.',9,False,RGBColor(0x6B,0x75,0x8C))
tb(11.4,7.03,1.3,0.25,'Báo cáo thử việc',9,False,RGBColor(0x6B,0x75,0x8C),PP_ALIGN.RIGHT)
prs.save(sys.argv[1]);print('saved')
