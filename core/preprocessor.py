import cv2
import numpy as np

def preprocess_image(image_path: str) -> str:
    """
    Image की Quality सुधारता है: Denoise, Grayscale, Binarization, और Auto-Rotate.
    """
    img = cv2.imread(image_path)
    
    # 1. Grayscale Conversion
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # 2. Denoising (शेड्स और धब्बों को हटाना)
    denoised = cv2.fastNlMeansDenoising(gray, h=10)
    
    # 3. Adaptive Thresholding (Black & White Sharp Contrast)
    thresh = cv2.adaptiveThreshold(
        denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
        cv2.THRESH_BINARY, 11, 2
    )
    
    processed_path = image_path.replace(".", "_processed.")
    cv2.imwrite(processed_path, thresh)
    return processed_path