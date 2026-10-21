import os
from pdf2image import convert_from_path

def convert_pdf_to_images(pdf_path: str, output_folder: str, job_id: str = None) -> list[str]:
    """
    Converts each page of a PDF file into PNG images and returns their file paths.
    Uses unique job_id prefix to prevent concurrency collisions.
    """
    poppler_path = os.environ.get("POPPLER_PATH")
    kwargs = {}
    if poppler_path and os.path.exists(poppler_path):
        kwargs["poppler_path"] = poppler_path

    images = convert_from_path(pdf_path, **kwargs)
    image_paths = []

    prefix = f"{job_id}_" if job_id else ""
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    for i, image in enumerate(images):
        img_filename = f"{prefix}{base_name}_page_{i+1}.png"
        img_path = os.path.join(output_folder, img_filename)
        image.save(img_path, 'PNG')
        image_paths.append(img_path)

    return image_paths