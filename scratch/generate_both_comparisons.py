import os
from PIL import Image, ImageDraw, ImageFont

def make_comp(orig_path, rend_path, out_path, title):
    img_orig = Image.open(orig_path).convert('RGB')
    img_rend = Image.open(rend_path).convert('RGB')

    target_h = 1000
    scale_o = target_h / img_orig.height
    w_o = int(img_orig.width * scale_o)
    img_orig_resized = img_orig.resize((w_o, target_h), Image.LANCZOS)

    scale_r = target_h / img_rend.height
    w_r = int(img_rend.width * scale_r)
    img_rend_resized = img_rend.resize((w_r, target_h), Image.LANCZOS)

    banner_h = 50
    canvas_w = w_o + w_r + 30
    canvas_h = target_h + banner_h

    comp = Image.new('RGB', (canvas_w, canvas_h), (245, 247, 250))
    draw = ImageDraw.Draw(comp)

    try:
        font = ImageFont.truetype("arial.ttf", 20)
    except Exception:
        font = ImageFont.load_default()

    draw.text((w_o // 2 - 80, 15), f"ORIGINAL ({title})", fill=(30, 41, 59), font=font)
    draw.text((w_o + 30 + w_r // 2 - 100, 15), "RECONSTRUCTED (WORD .DOCX)", fill=(30, 41, 59), font=font)

    comp.paste(img_orig_resized, (0, banner_h))
    comp.paste(img_rend_resized, (w_o + 30, banner_h))
    draw.line([(w_o + 15, 0), (w_o + 15, canvas_h)], fill=(203, 213, 225), width=2)

    comp.save(out_path)
    print(f"Saved {out_path} ({comp.size})")

make_comp('tests/fixtures/sample_resume.jpg', 'outputs/test_rs_page1.png', 'outputs/visual_comparison_richard_sanchez.png', 'Richard Sanchez')
make_comp('uploads/f3cba159622e4990aa6eed8749f7a0a2_images_1.jpg', 'outputs/test_f3c_page1.png', 'outputs/visual_comparison_brian_baxter.png', 'Brian Baxter')

