import os, sys, json
sys.path.insert(0, '.')
from core.docx_builder import build_advanced_docx, Document
import pypdfium2

sys.stdout.reconfigure(encoding='utf-8')

with open('scratch/rs_ir.json', 'r', encoding='utf-8') as f:
    ir = json.load(f)

docx_path = os.path.abspath('outputs/test_rs_reconstructed.docx')
pdf_path = os.path.abspath('outputs/test_rs_reconstructed.pdf')
png_path = os.path.abspath('outputs/test_rs_page1.png')

print("1. Building advanced docx...")
build_advanced_docx(ir, docx_path, mode="high_fidelity")
print(f"Created {docx_path} ({os.path.getsize(docx_path)} bytes)")

print("2. Converting to PDF via Word COM...")
import win32com.client
import pythoncom
pythoncom.CoInitialize()
word = win32com.client.DispatchEx("Word.Application")
try:
    word.Visible = False
    word.DisplayAlerts = 0
    wdoc = word.Documents.Open(docx_path)
    wdoc.SaveAs(pdf_path, FileFormat=17)
    wdoc.Close(SaveChanges=0)
finally:
    try:
        word.Quit()
    except Exception:
        pass
    del word
    pythoncom.CoUninitialize()

print(f"Created {pdf_path}")
pdf = pypdfium2.PdfDocument(pdf_path)
print(f"PDF Page Count: {len(pdf)}")
pdf[0].render(scale=1.5).to_pil().save(png_path)
print(f"Rendered {png_path}")

print("\n3. Inspecting DOCX Structure:")
doc = Document(docx_path)
t = doc.tables[0]
print("Layout table rows:", len(t.rows), "cols:", len(t.columns))
cell_l = t.cell(0, 0)
cell_r = t.cell(0, 1)

print("\n--- LEFT SIDEBAR TEXT ---")
for p in cell_l.paragraphs:
    if p.text.strip():
        print("  L:", repr(p.text[:60]))

print("\n--- RIGHT COLUMN TEXT & NESTED TABLES ---")
for p in cell_r.paragraphs:
    if p.text.strip():
        print("  R:", repr(p.text[:60]))

print(f"\nNested tables in right cell: {len(cell_r.tables)}")
for i, nt in enumerate(cell_r.tables):
    print(f"  Table {i}: {len(nt.rows)}x{len(nt.columns)}")
    for r in range(len(nt.rows)):
        for c in range(len(nt.columns)):
            txt = nt.cell(r, c).text.strip().replace('\n', ' | ')
            print(f"    Cell({r},{c}): {repr(txt[:70])}")

