import cv2
import numpy as np

def detect_tables_and_cells(image_path: str) -> list[dict]:
    """
    OpenCV का उपयोग करके इमेज में मौजूद Tables और उनके Bounding Boxes को खोजता है।
    """
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    
    # 1. Binarization (Otsu Thresholding)
    _, thresh = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY_INV)
    
    # 2. Horizontal and Vertical Line Kernels
    kernel_len = np.array(img).shape[1] // 30
    ver_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, kernel_len))
    hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_len, 1))
    
    # 3. Detect Vertical & Horizontal Lines
    image_1 = cv2.erode(thresh, ver_kernel, iterations=3)
    vertical_lines = cv2.dilate(image_1, ver_kernel, iterations=3)
    
    image_2 = cv2.erode(thresh, hor_kernel, iterations=3)
    horizontal_lines = cv2.dilate(image_2, hor_kernel, iterations=3)
    
    # 4. Table Grid Structure Combined
    table_grid = cv2.addWeighted(vertical_lines, 0.5, horizontal_lines, 0.5, 0.0)
    _, table_grid = cv2.threshold(table_grid, 128, 255, cv2.THRESH_BINARY)
    
    # 5. Find Table Contours
    contours, _ = cv2.findContours(table_grid, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    tables = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        # केवल बड़े Bounding Boxes को Table मानें (छोटी Noise रेखाओं को छोड़ें)
        if w > 100 and h > 50:
            tables.append({"type": "table", "x": x, "y": y, "width": w, "height": h})
            
    return tables