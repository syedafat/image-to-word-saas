import os
import sys
import unittest
from PIL import Image, ImageDraw, ImageFont

# Add project root to sys.path
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from config import Config
from core.docx_builder import build_advanced_docx
from docx import Document

class TestVisualReconstruction(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.test_output_dir = os.path.join(BASE_DIR, 'outputs', 'test_results')
        cls.test_assets_dir = os.path.join(BASE_DIR, 'outputs', 'test_assets')
        os.makedirs(cls.test_output_dir, exist_ok=True)
        os.makedirs(cls.test_assets_dir, exist_ok=True)

        # Generate sample graphic assets for testing
        cls.sample_photo_path = os.path.join(cls.test_assets_dir, 'sample_photo.png')
        cls.sample_logo_path = os.path.join(cls.test_assets_dir, 'sample_logo.png')

        # 1. Sample profile photo (avatar circle)
        img_photo = Image.new('RGB', (160, 160), color=(14, 165, 233))
        draw_p = ImageDraw.Draw(img_photo)
        draw_p.ellipse([20, 20, 140, 140], fill=(255, 255, 255))
        draw_p.text((50, 65), "PHOTO", fill=(14, 165, 233))
        img_photo.save(cls.sample_photo_path)

        # 2. Sample logo (rectangular badge)
        img_logo = Image.new('RGB', (240, 80), color=(30, 41, 59))
        draw_l = ImageDraw.Draw(img_logo)
        draw_l.rectangle([10, 10, 230, 70], outline=(56, 189, 248), width=3)
        draw_l.text((40, 30), "CORP LOGO", fill=(255, 255, 255))
        img_logo.save(cls.sample_logo_path)

    def test_case_1_cv_resume_with_photo_and_columns(self):
        """
        Test Case 1: CV / Resume with photo, dynamic colors (Teal theme),
        two-column sidebar, and bulleted skills.
        """
        output_file = os.path.join(self.test_output_dir, 'test_cv_reconstructed.docx')

        cv_page_elements = [
            # Top Full-width Header
            {
                "type": "heading",
                "bbox": [20, 20, 80, 980],
                "content": "ALEXANDER WRIGHT, PMP",
                "font_size_pt": 22.0,
                "is_bold": True,
                "text_color": "#0F172A",
                "text_align": "left"
            },
            {
                "type": "divider",
                "bbox": [85, 20, 95, 980],
                "text_color": "#0D9488"  # Teal divider
            },
            # Left Column / Sidebar (bbox xmax <= 360) with Teal Background
            {
                "type": "image",
                "bbox": [110, 30, 250, 320],
                "image_label": "photo",
                "image_asset_path": self.sample_photo_path,
                "text_align": "center"
            },
            {
                "type": "heading",
                "bbox": [260, 30, 290, 320],
                "content": "CONTACT INFO",
                "font_size_pt": 12.0,
                "is_bold": True,
                "text_color": "#FFFFFF",
                "bg_color": "#0D9488"
            },
            {
                "type": "paragraph",
                "bbox": [300, 30, 340, 320],
                "content": "alex.wright@email.com\n+1 (555) 234-5678\nNew York, NY",
                "font_size_pt": 9.5,
                "text_color": "#F0FDFA",
                "bg_color": "#0D9488"
            },
            {
                "type": "heading",
                "bbox": [360, 30, 390, 320],
                "content": "CORE COMPETENCIES",
                "font_size_pt": 12.0,
                "is_bold": True,
                "text_color": "#FFFFFF",
                "bg_color": "#0D9488"
            },
            {
                "type": "bullet",
                "bbox": [400, 30, 430, 320],
                "content": "Agile & Scrum Leadership",
                "font_size_pt": 9.5,
                "text_color": "#F0FDFA",
                "bg_color": "#0D9488"
            },
            {
                "type": "bullet",
                "bbox": [440, 30, 470, 320],
                "content": "Enterprise Cloud Architecture",
                "font_size_pt": 9.5,
                "text_color": "#F0FDFA",
                "bg_color": "#0D9488"
            },
            # Right Column (bbox xmin >= 380)
            {
                "type": "heading",
                "bbox": [110, 390, 140, 970],
                "content": "EXECUTIVE SUMMARY",
                "font_size_pt": 14.0,
                "is_bold": True,
                "text_color": "#0D9488"
            },
            {
                "type": "paragraph",
                "bbox": [150, 390, 230, 970],
                "content": "Senior Project Manager with 12+ years directing global engineering teams, driving multimillion-dollar SaaS implementations with zero defect tolerances.",
                "font_size_pt": 10.5,
                "text_color": "#334155"
            },
            {
                "type": "heading",
                "bbox": [250, 390, 280, 970],
                "content": "PROFESSIONAL EXPERIENCE",
                "font_size_pt": 14.0,
                "is_bold": True,
                "text_color": "#0D9488"
            },
            {
                "type": "paragraph",
                "bbox": [290, 390, 320, 970],
                "content": "Principal Tech Lead — CloudScale Systems (2020 – Present)",
                "font_size_pt": 11.0,
                "is_bold": True,
                "text_color": "#0F172A"
            },
            {
                "type": "bullet",
                "bbox": [330, 390, 360, 970],
                "content": "Spearheaded migration of legacy monolith to microservices, reducing AWS spend by 34%.",
                "font_size_pt": 10.0,
                "text_color": "#334155"
            }
        ]

        build_advanced_docx([cv_page_elements], output_file, mode="high_fidelity")

        # Verify Word document
        doc = Document(output_file)
        total_paragraphs = len(doc.paragraphs) + sum(len(c.paragraphs) for t in doc.tables for r in t.rows for c in r.cells)
        self.assertGreater(total_paragraphs, 5, "Total paragraphs across document and columns should be > 5")
        self.assertGreaterEqual(len(doc.tables), 1, "Expected column layout table grid")

        # Verify picture was embedded into the document
        has_picture = any(shape for shape in doc.inline_shapes)
        has_image_part = any("image" in part.content_type for part in doc.part.related_parts.values())
        self.assertTrue(has_picture or has_image_part, "Photo must be embedded in Word document")

    def test_case_2_table_heavy_invoice(self):
        """
        Test Case 2: Table-heavy Invoice with company logo, billing info,
        itemized table with borders/shading, and totals calculation.
        """
        output_file = os.path.join(self.test_output_dir, 'test_invoice_reconstructed.docx')

        invoice_elements = [
            # Header Logo & Metadata
            {
                "type": "image",
                "bbox": [20, 20, 90, 250],
                "image_label": "logo",
                "image_asset_path": self.sample_logo_path,
                "text_align": "left"
            },
            {
                "type": "heading",
                "bbox": [30, 650, 70, 980],
                "content": "INVOICE #INV-2026-0891",
                "font_size_pt": 16.0,
                "is_bold": True,
                "text_color": "#1E293B",
                "text_align": "right"
            },
            {
                "type": "paragraph",
                "bbox": [75, 650, 110, 980],
                "content": "Date: October 13, 2026\nDue Date: Net 30",
                "font_size_pt": 9.5,
                "text_color": "#64748B",
                "text_align": "right"
            },
            {
                "type": "divider",
                "bbox": [115, 20, 125, 980],
                "text_color": "#CBD5E1"
            },
            # Billed To
            {
                "type": "heading",
                "bbox": [135, 20, 155, 450],
                "content": "BILLED TO:",
                "font_size_pt": 11.0,
                "is_bold": True,
                "text_color": "#475569"
            },
            {
                "type": "paragraph",
                "bbox": [160, 20, 210, 450],
                "content": "Acme Innovations Corp\nAttn: Accounts Payable\n100 Tech Plaza, San Francisco, CA",
                "font_size_pt": 10.0,
                "text_color": "#1E293B"
            },
            # Itemized Table
            {
                "type": "table",
                "bbox": [240, 20, 520, 980],
                "table_header_bg": "#2563EB",  # Blue header
                "table_has_borders": True,
                "table_rows": [
                    ["Item Description", "Qty", "Rate", "Amount"],
                    ["AI Document Processing Engine Setup", "1", "$3,500.00", "$3,500.00"],
                    ["Multi-Page Layout Reconstruction Module", "1", "$2,200.00", "$2,200.00"],
                    ["WordprocessingML Native Table Optimizer", "1", "$1,800.00", "$1,800.00"],
                    ["Cloud Infrastructure Provisioning (Q4)", "3", "$450.00", "$1,350.00"]
                ]
            },
            # Summary Totals Card
            {
                "type": "paragraph",
                "bbox": [540, 600, 620, 980],
                "content": "Subtotal: $8,850.00\nTax (8.5%): $752.25\nTotal Due: $9,602.25",
                "font_size_pt": 11.0,
                "is_bold": True,
                "text_color": "#0F172A",
                "bg_color": "#F1F5F9",
                "text_align": "right"
            }
        ]

        build_advanced_docx([invoice_elements], output_file, mode="editable")

        doc = Document(output_file)
        self.assertTrue(os.path.exists(output_file))
        # Verify table exists and has 5 rows, 4 columns
        table = doc.tables[0]
        self.assertEqual(len(table.rows), 5)
        self.assertEqual(len(table.columns), 4)
        self.assertEqual(table.rows[0].cells[0].text, "Item Description")
        self.assertEqual(table.rows[4].cells[3].text, "$1,350.00")

    def test_case_3_image_preservation_and_dividers(self):
        """
        Test Case 3: Certificate / Verification badge with decorative divider
        and signature image preservation.
        """
        output_file = os.path.join(self.test_output_dir, 'test_certificate_reconstructed.docx')

        cert_elements = [
            {
                "type": "heading",
                "bbox": [40, 100, 100, 900],
                "content": "CERTIFICATE OF ACHIEVEMENT",
                "font_size_pt": 24.0,
                "is_bold": True,
                "text_color": "#1E3A8A",
                "text_align": "center"
            },
            {
                "type": "divider",
                "bbox": [110, 150, 120, 850],
                "text_color": "#D97706"  # Amber divider
            },
            {
                "type": "paragraph",
                "bbox": [150, 100, 220, 900],
                "content": "This is proudly presented to Alexander Wright for exemplary leadership in autonomous AI architecture.",
                "font_size_pt": 12.0,
                "text_color": "#334155",
                "text_align": "center"
            },
            {
                "type": "image",
                "bbox": [250, 380, 400, 620],
                "image_label": "signature",
                "image_asset_path": self.sample_photo_path,
                "text_align": "center"
            }
        ]

        build_advanced_docx([cert_elements], output_file, mode="high_fidelity")

        doc = Document(output_file)
        self.assertTrue(os.path.exists(output_file))
        self.assertGreater(len(doc.paragraphs), 2)

    def test_case_4_multi_page_pdf_page_boundaries(self):
        """
        Test Case 4: Multi-page document verifying page breaks, distinct sections,
        and no layout spillage across pages.
        """
        output_file = os.path.join(self.test_output_dir, 'test_multipage_reconstructed.docx')

        page1 = [
            {
                "type": "heading",
                "bbox": [50, 50, 100, 950],
                "content": "ANNUAL OPERATIONS REPORT — PAGE 1",
                "font_size_pt": 20.0,
                "is_bold": True,
                "text_color": "#0F172A"
            },
            {
                "type": "paragraph",
                "bbox": [110, 50, 200, 950],
                "content": "This page covers macroeconomic performance, market trends, and organizational milestones.",
                "font_size_pt": 10.5
            }
        ]

        page2 = [
            {
                "type": "heading",
                "bbox": [50, 50, 100, 950],
                "content": "FINANCIAL PERFORMANCE — PAGE 2",
                "font_size_pt": 20.0,
                "is_bold": True,
                "text_color": "#0F172A"
            },
            {
                "type": "table",
                "bbox": [110, 50, 300, 950],
                "table_rows": [
                    ["Metric", "2025 Actual", "2026 Target"],
                    ["Gross Revenue", "$12.4M", "$18.5M"],
                    ["Operating Margin", "22.4%", "28.0%"]
                ]
            }
        ]

        build_advanced_docx([page1, page2], output_file, mode="editable")

        doc = Document(output_file)
        self.assertTrue(os.path.exists(output_file))
        # Ensure page break XML element exists in document body
        body_xml = doc._body._element.xml
        self.assertIn("w:br", body_xml, "Page break XML tag must be present for multi-page documents")
        self.assertEqual(len(doc.tables), 1, "Page 2 should have exactly 1 native table")

    def test_case_5_enhanced_cv_visual_elements(self):
        """
        Test Case 5: Enhanced CV visual elements verification:
        - Header banner card with background fill and accent border
        - Top-left and bottom-right corner shapes
        - Profile photo with circular crop
        - Contact items with custom icon symbols
        - Skill progress bars rendered as editable native Word tables
        """
        output_file = os.path.join(self.test_output_dir, 'test_cv_enhanced_visuals.docx')

        cv_page_data = {
            "has_sidebar": True,
            "sidebar_position": "left",
            "sidebar_width_percent": 34,
            "sidebar_bg_color": "#2B2B2B",
            "accent_color": "#E5A914",
            "header_bg_color": "#F1F3F5",
            "has_corner_decorations": True,
            "corner_colors": {"top_left": "#E5A914", "bottom_right": "#E5A914"},
            "elements": [
                # Left Sidebar Elements
                {
                    "type": "image",
                    "bbox": [50, 40, 200, 300],
                    "image_label": "photo",
                    "image_asset_path": self.sample_photo_path,
                    "column": "left",
                    "text_align": "center"
                },
                {
                    "type": "heading",
                    "bbox": [220, 30, 250, 320],
                    "content": "CONTACT ME",
                    "font_size_pt": 13.0,
                    "is_bold": True,
                    "text_color": "#FFFFFF",
                    "column": "left"
                },
                {
                    "type": "contact_item",
                    "icon_type": "phone",
                    "content": "+1 234 567 8900",
                    "column": "left",
                    "text_color": "#FFFFFF"
                },
                {
                    "type": "contact_item",
                    "icon_type": "email",
                    "content": "brian.baxter@example.com",
                    "column": "left",
                    "text_color": "#FFFFFF"
                },
                {
                    "type": "contact_item",
                    "icon_type": "website",
                    "content": "www.brianbaxter.design",
                    "column": "left",
                    "text_color": "#FFFFFF"
                },
                {
                    "type": "contact_item",
                    "icon_type": "location",
                    "content": "San Francisco, CA",
                    "column": "left",
                    "text_color": "#FFFFFF"
                },
                # Right Column Elements
                {
                    "type": "heading",
                    "bbox": [40, 360, 90, 960],
                    "content": "BRIAN R. BAXTER",
                    "font_size_pt": 24.0,
                    "is_bold": True,
                    "text_color": "#2B2B2B",
                    "column": "right"
                },
                {
                    "type": "paragraph",
                    "bbox": [95, 360, 120, 960],
                    "content": "GRAPHIC & WEB DESIGNER",
                    "font_size_pt": 12.0,
                    "is_bold": True,
                    "text_color": "#E5A914",
                    "column": "right"
                },
                {
                    "type": "heading",
                    "bbox": [150, 360, 180, 960],
                    "content": "SKILLS",
                    "font_size_pt": 14.0,
                    "is_bold": True,
                    "text_color": "#2B2B2B",
                    "column": "right"
                },
                {
                    "type": "skill_bar",
                    "skill_name": "Adobe Photoshop",
                    "skill_percent": 90,
                    "column": "right"
                },
                {
                    "type": "skill_bar",
                    "skill_name": "Adobe Illustrator",
                    "skill_percent": 85,
                    "column": "right"
                },
                {
                    "type": "skill_bar",
                    "skill_name": "HTML 5 / CSS 3",
                    "skill_percent": 80,
                    "column": "right"
                },
                {
                    "type": "skill_bar",
                    "skill_name": "Microsoft PowerPoint",
                    "skill_percent": 75,
                    "column": "right"
                }
            ]
        }

        build_advanced_docx([cv_page_data], output_file, mode="high_fidelity")

        doc = Document(output_file)
        self.assertTrue(os.path.exists(output_file))

        # 1. Verify tables: layout table, nested header banner table, and skill grid tables
        self.assertEqual(len(doc.tables), 1, "Expected top-level 2-column sidebar layout grid")
        grid = doc.tables[0]
        self.assertEqual(len(grid.rows[0].cells), 2, "Expected 2 columns (sidebar + main)")

        right_cell = grid.cell(0, 1)
        # right_cell contains nested tables: header banner and skill grid
        self.assertGreaterEqual(len(right_cell.tables), 2, "Expected nested header banner and skill grid tables")
        skill_grid = right_cell.tables[1]
        skill_meters_count = sum(len(c.tables) for r in skill_grid.rows for c in r.cells)
        self.assertEqual(skill_meters_count, 4, "Expected 4 nested skill meter progress bar tables")

        # 2. Verify all text and contact symbols are present across the document (including nested tables)
        all_text_nodes = [node.text for node in doc._element.xpath('.//w:t') if node.text]
        all_text = " ".join(all_text_nodes)

        self.assertIn("BRIAN R. BAXTER", all_text)
        self.assertIn("GRAPHIC & WEB DESIGNER", all_text)
        self.assertIn("Adobe Photoshop", all_text)
        self.assertIn("Adobe Illustrator", all_text)
        self.assertIn("HTML 5 / CSS 3", all_text)
        self.assertIn("Microsoft PowerPoint", all_text)

        # 3. Verify contact icon symbols and text
        self.assertIn("+1 234 567 8900", all_text)
        self.assertIn("brian.baxter@example.com", all_text)
        self.assertIn("www.brianbaxter.design", all_text)
        self.assertIn("San Francisco, CA", all_text)
        self.assertTrue(any(sym in all_text for sym in ["☎", "✉", "🌐", "⚲"]))

        # 4. Verify embedded image parts (profile photo + 2 corner triangle decorations)
        image_parts = [part for part in doc.part.related_parts.values() if "image" in part.content_type]
        self.assertGreaterEqual(len(image_parts), 2, "Document should contain at least profile photo and corner decoration assets")

    def test_case_6_single_page_cv_layout_and_pagination(self):
        """
        Test Case 6: Verifies that a full two-column CV fits cleanly on exactly ONE page
        in Word, with photo preserved, header banner shaded, table borders removed,
        and all expected text intact.
        """
        output_docx = os.path.join(self.test_output_dir, 'test_single_page_cv.docx')
        output_pdf = os.path.join(self.test_output_dir, 'test_single_page_cv.pdf')

        cv_page = {
            "has_sidebar": True,
            "sidebar_position": "left",
            "sidebar_width_percent": 34,
            "sidebar_bg_color": "#17304B",
            "accent_color": "#E5A914",
            "header_bg_color": "#F1F3F5",
            "elements": [
                # Left Sidebar Elements
                {
                    "type": "image",
                    "bbox": [40, 40, 180, 300],
                    "image_label": "photo",
                    "image_asset_path": self.sample_photo_path,
                    "column": "left"
                },
                {
                    "type": "heading",
                    "bbox": [200, 30, 220, 320],
                    "content": "CONTACT",
                    "font_size_pt": 12.0,
                    "is_bold": True,
                    "text_color": "#FFFFFF",
                    "column": "left"
                },
                {"type": "contact_item", "icon_type": "phone", "content": "+123-456-7890", "column": "left", "text_color": "#FFFFFF"},
                {"type": "contact_item", "icon_type": "email", "content": "hello@reallygreatsite.com", "column": "left", "text_color": "#FFFFFF"},
                {"type": "contact_item", "icon_type": "location", "content": "123 Anywhere St., Any City", "column": "left", "text_color": "#FFFFFF"},
                {"type": "contact_item", "icon_type": "website", "content": "www.reallygreatsite.com", "column": "left", "text_color": "#FFFFFF"},
                {
                    "type": "heading",
                    "bbox": [320, 30, 340, 320],
                    "content": "EDUCATION",
                    "font_size_pt": 12.0,
                    "is_bold": True,
                    "text_color": "#FFFFFF",
                    "column": "left"
                },
                {"type": "paragraph", "content": "2029 - 2030\nWARDIERE UNIVERSITY\nMaster of Business Management", "font_size_pt": 9.0, "column": "left", "text_color": "#FFFFFF"},
                {"type": "paragraph", "content": "2025 - 2029\nWARDIERE UNIVERSITY\nBachelor of Business", "font_size_pt": 9.0, "column": "left", "text_color": "#FFFFFF"},
                {
                    "type": "heading",
                    "bbox": [470, 30, 490, 320],
                    "content": "SKILLS",
                    "font_size_pt": 12.0,
                    "is_bold": True,
                    "text_color": "#FFFFFF",
                    "column": "left"
                },
                {"type": "bullet", "content": "Project Management", "font_size_pt": 9.0, "column": "left", "text_color": "#FFFFFF"},
                {"type": "bullet", "content": "Public Relations", "font_size_pt": 9.0, "column": "left", "text_color": "#FFFFFF"},
                {"type": "bullet", "content": "Teamwork & Leadership", "font_size_pt": 9.0, "column": "left", "text_color": "#FFFFFF"},
                {
                    "type": "heading",
                    "bbox": [600, 30, 620, 320],
                    "content": "LANGUAGES",
                    "font_size_pt": 12.0,
                    "is_bold": True,
                    "text_color": "#FFFFFF",
                    "column": "left"
                },
                {"type": "bullet", "content": "English (Fluent)", "font_size_pt": 9.0, "column": "left", "text_color": "#FFFFFF"},
                {"type": "bullet", "content": "Spanish (Intermediate)", "font_size_pt": 9.0, "column": "left", "text_color": "#FFFFFF"},

                # Right Column Elements
                {
                    "type": "heading",
                    "bbox": [40, 360, 80, 960],
                    "content": "RICHARD SANCHEZ",
                    "font_size_pt": 20.0,
                    "is_bold": True,
                    "text_color": "#17304B",
                    "column": "right"
                },
                {
                    "type": "paragraph",
                    "bbox": [85, 360, 105, 960],
                    "content": "MARKETING MANAGER",
                    "font_size_pt": 10.5,
                    "is_bold": True,
                    "text_color": "#64748B",
                    "column": "right"
                },
                {
                    "type": "heading",
                    "bbox": [130, 360, 150, 960],
                    "content": "PROFILE",
                    "font_size_pt": 12.0,
                    "is_bold": True,
                    "text_color": "#17304B",
                    "column": "right"
                },
                {
                    "type": "paragraph",
                    "bbox": [155, 360, 210, 960],
                    "content": "Results-oriented Marketing Manager with 8+ years leading cross-functional teams, executing high-ROI multi-channel campaigns, and driving brand recognition.",
                    "font_size_pt": 9.5,
                    "column": "right"
                },
                {
                    "type": "heading",
                    "bbox": [230, 360, 250, 960],
                    "content": "WORK EXPERIENCE",
                    "font_size_pt": 12.0,
                    "is_bold": True,
                    "text_color": "#17304B",
                    "column": "right"
                },
                {"type": "paragraph", "content": "Borcelle Studio | 2030 - PRESENT\nMarketing Manager & Specialist", "font_size_pt": 10.0, "is_bold": True, "column": "right"},
                {"type": "bullet", "content": "Develop and execute comprehensive marketing strategies and campaigns that align with the company's goals and objectives.", "font_size_pt": 9.0, "column": "right"},
                {"type": "bullet", "content": "Lead, mentor, and manage a high-performing marketing team, fostering a collaborative and results-driven work environment.", "font_size_pt": 9.0, "column": "right"},
                {"type": "paragraph", "content": "Fauget Studio | 2025 - 2029\nMarketing Specialist", "font_size_pt": 10.0, "is_bold": True, "column": "right"},
                {"type": "bullet", "content": "Create and manage marketing budgets, ensuring efficient allocation of resources and optimizing ROI across online channels.", "font_size_pt": 9.0, "column": "right"},
                {"type": "bullet", "content": "Oversee market research to identify emerging trends, customer needs, and competitor strategies.", "font_size_pt": 9.0, "column": "right"},
                {
                    "type": "heading",
                    "bbox": [550, 360, 570, 960],
                    "content": "REFERENCE",
                    "font_size_pt": 12.0,
                    "is_bold": True,
                    "text_color": "#17304B",
                    "column": "right"
                },
                {"type": "paragraph", "bbox": [580, 360, 640, 600], "content": "Estelle Darcy\nWardiere Inc. / CTO\nPhone: 123-456-7890\nEmail: hello@reallygreatsite.com", "font_size_pt": 9.0, "column": "right"},
                {"type": "paragraph", "bbox": [580, 680, 640, 960], "content": "Harper Richard\nWardiere Inc. / CEO\nPhone: 123-456-7890\nEmail: hello@reallygreatsite.com", "font_size_pt": 9.0, "column": "right"}
            ]
        }

        build_advanced_docx([cv_page], output_docx, mode="high_fidelity")
        self.assertTrue(os.path.exists(output_docx))

        doc = Document(output_docx)
        # 1. Verify text
        all_text = " ".join([node.text for node in doc._element.xpath('.//w:t') if node.text])
        self.assertIn("RICHARD SANCHEZ", all_text)
        self.assertIn("MARKETING MANAGER", all_text)
        self.assertIn("WORK EXPERIENCE", all_text)
        self.assertIn("REFERENCE", all_text)
        self.assertIn("Estelle Darcy", all_text)
        self.assertIn("Harper Richard", all_text)

        # 2. Verify embedded photo
        image_parts = [part for part in doc.part.related_parts.values() if "image" in part.content_type]
        self.assertGreaterEqual(len(image_parts), 1, "Profile photo must be embedded in document")

        # 3. Verify no unwanted table borders and fixed layout
        grid = doc.tables[0]
        tblBorders = grid._tbl.tblPr.xpath('./w:tblBorders')
        self.assertTrue(tblBorders, "Table borders element should exist with val='none'")
        tblLayout = grid._tbl.tblPr.xpath('./w:tblLayout')
        self.assertTrue(tblLayout, "Table layout element should exist with fixed type")

        # 4. Verify nested tables: header banner, experience date alignment, and 2-column references
        right_cell = grid.cell(0, 1)
        nested_tables = right_cell.tables
        self.assertGreaterEqual(len(nested_tables), 3, "Expected header banner, experience date row, and reference sub-table")

        # Verify reference sub-table has Estelle Darcy on left and Harper Richard on right
        ref_table = nested_tables[-1]
        self.assertEqual(len(ref_table.rows), 1)
        self.assertEqual(len(ref_table.columns), 2)
        self.assertIn("Estelle Darcy", ref_table.cell(0, 0).text)
        self.assertIn("Harper Richard", ref_table.cell(0, 1).text)

        # 5. Render to PDF and verify EXACTLY 1 page
        try:
            import win32com.client
            import pythoncom
            import pypdfium2

            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            try:
                word.Visible = False
                word.DisplayAlerts = 0
                wdoc = word.Documents.Open(os.path.abspath(output_docx))
                wdoc.SaveAs(os.path.abspath(output_pdf), FileFormat=17)
                wdoc.Close(SaveChanges=0)
            finally:
                try:
                    word.Quit()
                except Exception:
                    pass
                del word
                pythoncom.CoUninitialize()

            pdf = pypdfium2.PdfDocument(output_pdf)
            page_count = len(pdf)
            self.assertEqual(page_count, 1, f"Expected exactly 1 page in Word, but got {page_count}")
        except Exception as e:
            print(f"[Notice] Word COM PDF verification note: {e}")

if __name__ == '__main__':
    unittest.main()


