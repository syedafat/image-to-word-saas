import os, sys, json
sys.path.insert(0, '.')
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client
import pythoncom
import pypdfium2

sys.stdout.reconfigure(encoding='utf-8')

# Test trHeight atLeast
doc = Document()
for s in doc.sections:
    s.page_width = Inches(8.5)
    s.page_height = Inches(11.0)
    s.top_margin = Inches(0.35)
    s.bottom_margin = Inches(0.35)
    s.left_margin = Inches(0.35)
    s.right_margin = Inches(0.35)

tbl = doc.add_table(rows=1, cols=2)
trPr = tbl.rows[0]._tr.get_or_add_trPr()
trPr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))
# 10.2 inches = 10.2 * 1440 = 14688 dxa
trPr.append(parse_xml(r'<w:trHeight %s w:val="14400" w:hRule="atLeast"/>' % nsdecls('w')))

c_l = tbl.cell(0, 0)
c_r = tbl.cell(0, 1)
c_l.width = Inches(2.6)
c_r.width = Inches(5.2)

# Blue sidebar
shd = parse_xml(r'<w:shd %s w:fill="163853"/>' % nsdecls('w'))
c_l._tc.get_or_add_tcPr().append(shd)

p_l = c_l.paragraphs[0]
r_l = p_l.add_run("SIDEBAR CONTENT")
r_l.font.color.rgb = RGBColor(255, 255, 255)

p_r = c_r.paragraphs[0]
r_r = p_r.add_run("MAIN BODY CONTENT")

p_tail = doc.add_paragraph()
p_tail.paragraph_format.line_spacing = Pt(1)
p_tail.paragraph_format.space_before = Pt(0)
p_tail.paragraph_format.space_after = Pt(0)
r_t = p_tail.add_run()
r_t.font.size = Pt(1)

test_docx = os.path.abspath('scratch/test_trheight.docx')
test_pdf = os.path.abspath('scratch/test_trheight.pdf')
doc.save(test_docx)

pythoncom.CoInitialize()
word = win32com.client.DispatchEx("Word.Application")
try:
    word.Visible = False
    word.DisplayAlerts = 0
    wdoc = word.Documents.Open(test_docx)
    wdoc.SaveAs(test_pdf, FileFormat=17)
    wdoc.Close(SaveChanges=0)
finally:
    try:
        word.Quit()
    except Exception:
        pass
    del word
    pythoncom.CoUninitialize()

pdf = pypdfium2.PdfDocument(test_pdf)
print(f"Page count with trHeight 14400: {len(pdf)}")
pdf[0].render(scale=1.5).to_pil().save('scratch/test_trheight_page1.png')
print("Rendered scratch/test_trheight_page1.png")
