def sort_and_group_layout(ocr_elements: list[dict], line_threshold=12) -> list[dict]:
    """
    OCR वर्ड्स को सेम Line Height के आधार पर सेंटेंस और पैराग्राफ में ग्रुप करता है
    और Floating Textbox के लिए आवश्यक Coordinates बनाए रखता है।
    """
    if not ocr_elements:
        return []

    # Y-coordinate के हिसाब से सॉर्ट करें
    sorted_by_y = sorted(ocr_elements, key=lambda e: e['y'])
    lines = []
    current_line = [sorted_by_y[0]]

    for element in sorted_by_y[1:]:
        if abs(element['y'] - current_line[0]['y']) <= line_threshold:
            current_line.append(element)
        else:
            lines.append(sorted(current_line, key=lambda e: e['x']))
            current_line = [element]

    if current_line:
        lines.append(sorted(current_line, key=lambda e: e['x']))

    structured_ir = []
    for line in lines:
        line_text = " ".join([e['text'] for e in line])
        avg_height = sum([e['height'] for e in line]) / len(line)
        min_x = line[0]['x']
        max_x = max([e['x'] + e['width'] for e in line])
        line_width = max_x - min_x

        structured_ir.append({
            "type": "text",
            "content": line_text,
            "line_text": line_text,
            "x": min_x,
            "y": line[0]['y'],
            "width": line_width,
            "height": int(avg_height),
            "avg_height": avg_height,
            "element_count": len(line),
            "use_floating": True,  # XML Floating Text Box का सिग्नल
            "column": "single"
        })

    return structured_ir


def detect_columns_and_group(ocr_elements: list[dict], page_width_px=2480) -> list[dict]:
    """
    1-Column या 2-Column लेआउट को डिटेक्ट करता है और एलिमेंट्स को प्रॉपर Floating Structure प्रदान करता है।
    """
    if not ocr_elements:
        return []

    mid_x = page_width_px / 2
    left_column = []
    right_column = []
    full_width_elements = []

    for el in ocr_elements:
        if el.get('width', 0) > (page_width_px * 0.5) or abs(el['x'] - mid_x) < 80:
            full_width_elements.append(el)
        elif el['x'] < mid_x:
            left_column.append(el)
        else:
            right_column.append(el)

    # अगर 2-Column स्ट्रक्चर पाया जाता है
    if len(left_column) > 5 and len(right_column) > 5:
        sorted_left = sort_and_group_layout(left_column)
        sorted_right = sort_and_group_layout(right_column)
        sorted_full = sort_and_group_layout(full_width_elements)

        final_ir = []
        for line in sorted_full:
            line['column'] = 'single'
            final_ir.append(line)

        for line in sorted_left:
            line['column'] = 'left'
            final_ir.append(line)

        for line in sorted_right:
            line['column'] = 'right'
            final_ir.append(line)

        return sorted(final_ir, key=lambda item: item['y'])

    # Standard Single Column Document Logic
    sorted_elements = sort_and_group_layout(ocr_elements)
    for line in sorted_elements:
        line['column'] = 'single'

    return sorted(sorted_elements, key=lambda item: item['y'])


def merge_tables_and_text(text_elements: list[dict], tables: list[dict]) -> list[dict]:
    """
    टेबल्स और टेक्स्ट ब्लॉक्स को सही सीक्वेंस में मिलाकर फाइनल लेआउट तैयार करता है।
    """
    combined_ir = []

    for text_item in text_elements:
        is_inside_table = False
        for table in tables:
            if (table['x'] <= text_item['x'] <= table['x'] + table['width'] and
                table['y'] <= text_item['y'] <= table['y'] + table['height']):
                is_inside_table = True
                break

        if not is_inside_table:
            text_item["type"] = "text"
            text_item["use_floating"] = True
            combined_ir.append(text_item)

    for tbl in tables:
        tbl["type"] = "table"
        combined_ir.append(tbl)

    return sorted(combined_ir, key=lambda item: item['y'])