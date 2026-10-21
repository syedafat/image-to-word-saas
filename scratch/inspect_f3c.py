import os, sys, json
sys.path.insert(0, '.')
from core.ocr_engine import extract_structured_ir

sys.stdout.reconfigure(encoding='utf-8')
ir = extract_structured_ir('uploads/f3cba159622e4990aa6eed8749f7a0a2_images_1.jpg', job_id='inspect_f3c', page_num=1)
with open('scratch/f3c_ir.json', 'w', encoding='utf-8') as f:
    json.dump(ir, f, indent=2)

print("Saved scratch/f3c_ir.json")
print("has_sidebar:", ir.get("has_sidebar"))
print("sidebar_pct:", ir.get("sidebar_width_percent"))
print("sidebar_bg:", ir.get("sidebar_bg_color"))
print("elements:", len(ir.get("elements", [])))
for i, el in enumerate(ir.get("elements", [])):
    t = el.get("text") or el.get("content") or ""
    print(f"{i:2d}: type={el.get('type'):12s} col={str(el.get('column')):7s} bbox={str(el.get('bbox'))[:20]:20s} text={repr(t[:45])}")

