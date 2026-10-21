import os
import uuid
import json
import warnings
from PIL import Image
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

from config import Config

warnings.filterwarnings("ignore", category=UserWarning)

class VisualElement(BaseModel):
    type: str = Field(
        default="paragraph",
        description="Element type: 'heading', 'paragraph', 'bullet', 'table', 'image', 'divider', 'skill_bar', 'contact_item'"
    )
    text: str = Field(
        default="",
        description="The exact extracted text content. Must not be empty for headings, paragraphs, and bullets. For skill_bar, put skill name if skill_name is empty."
    )
    column: str = Field(
        default="single",
        description="Layout column: 'left' for left sidebar/column, 'right' for right main body, 'single' for full-width across page."
    )
    font_size_pt: float = Field(
        default=10.5,
        description="Estimated font size in points (e.g. 24 for main title, 14 for section heading, 10 for body text, 8 for captions)."
    )
    is_bold: bool = Field(
        default=False,
        description="True if text is bold or heavy weight."
    )
    is_italic: bool = Field(
        default=False,
        description="True if text is italicized."
    )
    text_color: str = Field(
        default="#0F172A",
        description="Hex text color code (e.g. #000000, #FFFFFF, #1E3A5F, #334155)."
    )
    bg_color: str = Field(
        default="",
        description="Hex background color if this element sits inside a colored card or container."
    )
    text_align: str = Field(
        default="left",
        description="Text alignment: 'left', 'center', or 'right'."
    )
    bbox: list[int] = Field(
        default_factory=list,
        description="[ymin, xmin, ymax, xmax] coordinates normalized on a 0 to 1000 scale. MANDATORY for images, photos, logos, signatures!"
    )
    image_label: str = Field(
        default="",
        description="If type='image', specify category: 'photo', 'logo', 'signature', 'icon'."
    )
    table_rows: list[list[str]] = Field(
        default_factory=list,
        description="2D array of rows and cells if type='table'."
    )
    table_header_bg: str = Field(
        default="#F1F5F9",
        description="Hex background color for table header row if shaded."
    )
    icon_type: str = Field(
        default="",
        description="For contact info or bullet icons: 'phone', 'email', 'website', 'location', 'bullet_dot', 'star'."
    )
    skill_name: str = Field(
        default="",
        description="Skill name if type='skill_bar', e.g. 'Adobe Photoshop', 'Python', 'Leadership'."
    )
    skill_percent: int = Field(
        default=0,
        description="Proficiency percentage between 10 and 100 for skill level / progress rating bars."
    )

class PageLayout(BaseModel):
    has_sidebar: bool = Field(
        default=False,
        description="True if page has a vertical sidebar panel (e.g. CV sidebar, colored vertical column)."
    )
    sidebar_position: str = Field(
        default="none",
        description="'left', 'right', or 'none'."
    )
    sidebar_width_percent: int = Field(
        default=33,
        description="Sidebar width percentage, e.g. 30 to 40."
    )
    sidebar_bg_color: str = Field(
        default="",
        description="Hex background color of the sidebar panel, e.g. #2B2B2B, #17304B, #1E3A5F, #0D9488."
    )
    accent_color: str = Field(
        default="",
        description="Primary accent or highlight color in hex, e.g. #E5A914, #F59E0B, #0D9488, #2563EB used for icons, bullet dots, lines, and skill bars."
    )
    header_bg_color: str = Field(
        default="",
        description="Hex background color of header banner/card if present, e.g. #F1F3F5."
    )
    has_corner_decorations: bool = Field(
        default=False,
        description="True if page has decorative geometric corner shapes or triangles (e.g. yellow corner triangles in top-left or bottom-right)."
    )
    corner_colors: list[str] = Field(
        default_factory=list,
        description="Hex colors of corner decorative shapes if present."
    )
    elements: list[VisualElement] = Field(
        default_factory=list,
        description="Ordered list of visual elements on this page."
    )

def crop_image_assets(elements: list[dict], original_image_path: str, job_id: str, page_num: int) -> list[dict]:
    """
    Crops detected image elements (photos, logos, signatures) from the original
    high-resolution page image and stores their asset paths in the element metadata.
    Includes fallback bounding-box logic for profile photos if coordinates were omitted.
    """
    if not os.path.exists(original_image_path):
        return elements

    try:
        page_img = Image.open(original_image_path)
        img_w, img_h = page_img.size
        assets_dir = Config.ASSET_FOLDER
        os.makedirs(assets_dir, exist_ok=True)

        for idx, el in enumerate(elements):
            el_type = el.get("type", "")
            bbox = el.get("bbox", [])

            # Fallback for photo/profile if bbox is missing
            if el_type == "image" and len(bbox) != 4:
                lbl = str(el.get("image_label", "")).lower()
                col = el.get("column", "left")
                if "photo" in lbl or "profile" in lbl or col == "left":
                    # Top-left sidebar portrait area fallback (0..1000 scale)
                    bbox = [80, 80, 310, 330]
                    el["bbox"] = bbox

            if el_type == "image" and len(bbox) == 4:
                ymin, xmin, ymax, xmax = bbox
                # Convert normalized 0..1000 coordinates to actual pixel coordinates
                x1 = max(0, min(img_w - 1, int((xmin / 1000.0) * img_w)))
                y1 = max(0, min(img_h - 1, int((ymin / 1000.0) * img_h)))
                x2 = max(x1 + 10, min(img_w, int((xmax / 1000.0) * img_w)))
                y2 = max(y1 + 10, min(img_h, int((ymax / 1000.0) * img_h)))

                if (x2 - x1) >= 15 and (y2 - y1) >= 15:
                    crop = page_img.crop((x1, y1, x2, y2))
                    asset_filename = f"{job_id}_p{page_num}_img{idx}_{uuid.uuid4().hex[:6]}.png"
                    asset_path = os.path.join(assets_dir, asset_filename)
                    crop.save(asset_path, "PNG")
                    el["image_asset_path"] = asset_path
                    el["pixel_width"] = x2 - x1
                    el["pixel_height"] = y2 - y1

    except Exception as e:
        print(f"[Asset Extraction Error]: {e}")

    return elements

def extract_structured_ir(image_path: str, job_id: str = "temp", page_num: int = 1) -> dict:
    """
    Analyzes document visual layout and extracts structured elements including page layout,
    typography, colors, bounding boxes, tables, and embedded graphics using Google Gemini Multimodal AI.
    Returns a PageLayout dictionary.
    """
    api_key = Config.GEMINI_API_KEY
    if not api_key:
        raise ValueError("GEMINI_API_KEY is not configured in environment or .env file.")

    if not os.path.exists(image_path):
        raise FileNotFoundError(f"Image not found at path: {image_path}")

    client = genai.Client(api_key=api_key)
    img = Image.open(image_path)

    prompt = """
    You are an expert Document Visual Layout & Graphic Reconstruction Engine.
    Analyze this document page with extreme precision and extract ALL visual elements, layout structure, colors, graphics, contact icons, and skill bars.

    Instructions:
    1. SIDEBAR & ACCENTS:
       - Check if the page has a sidebar (e.g. CV/Resume with a colored left or right panel).
       - If yes, set has_sidebar=True, sidebar_position='left' (or 'right'), sidebar_width_percent (e.g. 33), and sidebar_bg_color (exact hex color, e.g. #2B2B2B, #17304B, #1E3A5F).
       - Detect accent_color: The prominent accent/highlight color used for icons, bullet dots, decorative shapes, and skill bars (e.g. #E5A914, #F59E0B, #0D9488, #2563EB).
       - Detect header_bg_color: If the candidate name/title sits inside a light grey or colored banner box, detect its hex color (e.g. #F1F3F5).
       - Detect corner decorations: If there are decorative geometric corner shapes (like yellow/colored corner triangles in top-left or bottom-right), set has_corner_decorations=True and add their hex colors to corner_colors.
       - If no sidebar (e.g. standard invoice, form, letter, single-column report), set has_sidebar=False, sidebar_position='none'.

    2. ELEMENT PLACEMENT (COLUMN):
       - If sidebar is present:
         - Mark all elements inside the sidebar as column='left'.
         - Mark all elements in the main body as column='right'.
         - Mark any full-width header banner or footer as column='single'.
       - For standard documents without sidebars, set column='single'.

    3. COMPLETE TEXT & CONTACT ICONS:
       - Put the exact text content into the `text` field for EVERY heading, paragraph, job title, date, contact info, and bullet item.
       - For contact items (phone, email, website, address), set type='contact_item' or 'paragraph' and set icon_type ('phone', 'email', 'website', 'location').
       - NEVER leave `text` blank for text elements!

    4. SKILL LEVEL BARS:
       - Look closely at the SKILLS section. Extract EVERY single skill item!
       - For each skill with a horizontal progress or slider bar (e.g. Adobe Photoshop, Adobe Illustrator, HTML 5 / CSS 3, Microsoft PowerPoint), set type='skill_bar', skill_name, and estimate skill_percent (e.g. 90, 80, 75, 70).
       - DO NOT skip any skill!

    5. TYPOGRAPHY & COLORS:
       - Detect exact font_size_pt (e.g. 24 for candidate name or main title, 14 for section header, 10 for body, 9 for captions).
       - Detect is_bold and is_italic.
       - Detect exact text_color in hex (e.g. #FFFFFF for white text on dark sidebar, #000000, #334155).
       - Detect text_align ('left', 'center', 'right').

    6. IMAGES & GRAPHICAL ELEMENTS:
       - If you see a photo (profile picture), company logo, signature, stamp, or emblem:
         - Set type='image'
         - Provide exact [ymin, xmin, ymax, xmax] (0-1000 scale) bounding box covering the graphic. MANDATORY!
         - Specify image_label ('photo', 'logo', 'signature', 'icon').
         - Leave text as empty string.

    7. TABLES:
       - If there is a table or invoice grid, set type='table', populate table_rows with 2D array of cell texts, and detect table_header_bg.
    """

    models_to_try = [
        Config.GEMINI_MODEL,
        'gemini-3.5-flash-lite',
        'gemini-3.1-flash-lite',
        'gemini-3.8-flash'
    ]
    last_error = None

    for model_name in models_to_try:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[prompt, img],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=PageLayout,
                    temperature=0.1
                )
            )

            raw_text = response.text.strip() if response.text else "{}"
            page_layout_dict = json.loads(raw_text)

            raw_elements = page_layout_dict.get("elements", [])

            # Crop and preserve actual image assets (photos, logos, signatures)
            enriched_elements = crop_image_assets(
                elements=raw_elements,
                original_image_path=image_path,
                job_id=job_id,
                page_num=page_num
            )

            page_layout_dict["elements"] = enriched_elements
            return page_layout_dict

        except Exception as e:
            last_error = e
            print(f"[Gemini Error on {model_name}]: {e}")
            continue

    raise RuntimeError(f"Visual layout extraction failed across attempted models: {last_error}")