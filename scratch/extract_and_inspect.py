import json, sys, os
sys.path.insert(0, '.')
from core.ocr_engine import extract_structured_ir

sys.stdout.reconfigure(encoding='utf-8')

ir = extract_structured_ir('tests/fixtures/sample_resume.jpg', job_id='inspect_rs', page_num=1)
with open('scratch/rs_ir.json', 'w', encoding='utf-8') as f:
    json.dump(ir, f, indent=2)
print('Saved scratch/rs_ir.json')
print('Elements count:', len(ir.get('elements', [])))
for i, el in enumerate(ir.get('elements', [])):
    t = el.get('text') or el.get('content') or ''
    print(f"{i}: type={el.get('type')} col={el.get('column')} bbox={el.get('bbox')} text={repr(t[:40])}")
