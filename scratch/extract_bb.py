import json, sys, os
sys.path.insert(0, '.')
from core.ocr_engine import extract_structured_ir

sys.stdout.reconfigure(encoding='utf-8')
ir = extract_structured_ir('uploads/19f44a7ee0854c30b8c2356de4a1e440_images_1.jpg', job_id='inspect_bb', page_num=1)
with open('scratch/bb_ir.json', 'w', encoding='utf-8') as f:
    json.dump(ir, f, indent=2)
print('Saved scratch/bb_ir.json')
print('Elements count:', len(ir.get('elements', [])))
for i, el in enumerate(ir.get('elements', [])):
    t = el.get('text') or el.get('content') or ''
    print(f"{i}: type={el.get('type')} col={el.get('column')} text={repr(t[:40])}")

