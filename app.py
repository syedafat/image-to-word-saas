import os
import uuid
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename
from apscheduler.schedulers.background import BackgroundScheduler

from config import Config
from core.job_manager import job_manager
from core.ocr_engine import extract_structured_ir
from core.docx_builder import build_advanced_docx
from core.pdf_converter import convert_pdf_to_images
from core.cleaner import cleanup_old_files

app = Flask(__name__)
app.config.from_object(Config)

# Ensure upload, output, and asset directories exist
os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)
os.makedirs(Config.OUTPUT_FOLDER, exist_ok=True)
os.makedirs(Config.ASSET_FOLDER, exist_ok=True)

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS

# Background Cleanup (Hourly cleanup every 15 minutes)
scheduler = BackgroundScheduler()
scheduler.add_job(
    func=lambda: (
        cleanup_old_files(Config.UPLOAD_FOLDER, max_age_seconds=Config.FILE_MAX_AGE_SECONDS),
        cleanup_old_files(Config.OUTPUT_FOLDER, max_age_seconds=Config.FILE_MAX_AGE_SECONDS),
        cleanup_old_files(Config.ASSET_FOLDER, max_age_seconds=Config.FILE_MAX_AGE_SECONDS),
        job_manager.cleanup_expired_jobs(max_age_seconds=Config.FILE_MAX_AGE_SECONDS)
    ),
    trigger="interval",
    minutes=Config.CLEANUP_INTERVAL_MINUTES
)
scheduler.start()

def run_conversion_job(job_id: str):
    """
    Background worker task to visually reconstruct document layouts, extract tables,
    crop embedded images/logos, and generate Word (.docx) documents.
    """
    job = job_manager.get_job(job_id)
    if not job:
        return

    uploaded_path = job["file_path"]
    original_filename = job["original_filename"]
    mode = job.get("mode", "editable")
    base_name = os.path.splitext(original_filename)[0]
    safe_base_name = secure_filename(base_name) or "document"
    ext = original_filename.rsplit('.', 1)[1].lower()

    image_paths = []
    docx_filename = f"DocuShift_{safe_base_name}_{job_id[:8]}.docx"
    docx_path = os.path.join(Config.OUTPUT_FOLDER, docx_filename)

    try:
        job_manager.update_job(
            job_id,
            status="processing",
            progress=10,
            message="Analyzing document visual structure and styling..."
        )

        # 1. Page Splitting / Preparation
        if ext == 'pdf':
            job_manager.update_job(
                job_id,
                progress=15,
                message="Converting PDF pages to high-resolution images..."
            )
            image_paths = convert_pdf_to_images(uploaded_path, Config.UPLOAD_FOLDER, job_id=job_id)
        else:
            image_paths = [uploaded_path]

        total_pages = len(image_paths)
        job_manager.update_job(
            job_id,
            total_pages=total_pages,
            progress=20,
            message=f"Detected {total_pages} page(s). Starting AI visual layout & image extraction..."
        )

        # 2. Multi-Page Visual Layout & Asset Extraction
        all_pages_ir = []
        for idx, img_path in enumerate(image_paths):
            current_page = idx + 1
            page_progress = 20 + int((idx / total_pages) * 60)
            job_manager.update_job(
                job_id,
                current_page=current_page,
                progress=page_progress,
                message=f"Extracting layout, tables & graphics from page {current_page} of {total_pages}..."
            )

            page_ir = extract_structured_ir(
                image_path=img_path,
                job_id=job_id,
                page_num=current_page
            )
            all_pages_ir.append(page_ir)

        # 3. Native DOCX Generation with Visual Reconstruction
        job_manager.update_job(
            job_id,
            progress=85,
            message=f"Reconstructing Word document in {mode.replace('_', '-').title()} mode..."
        )
        build_advanced_docx(all_pages_ir, docx_path, mode=mode)

        # 4. Cleanup intermediate page images
        job_manager.update_job(
            job_id,
            progress=95,
            message="Finalizing document formatting and cleaning temporary files..."
        )
        for path in image_paths:
            if path != uploaded_path and os.path.exists(path):
                try:
                    os.remove(path)
                except Exception as e:
                    print(f"Warning: could not delete temp image {path}: {e}")

        # 5. Mark Completed
        job_manager.update_job(
            job_id,
            status="completed",
            progress=100,
            message="Document visual reconstruction completed successfully!",
            download_url=f"/download/{job_id}",
            docx_path=docx_path,
            docx_filename=docx_filename
        )

    except Exception as e:
        print(f"[Conversion Error for Job {job_id}]: {e}")
        job_manager.update_job(
            job_id,
            status="failed",
            progress=100,
            error=str(e),
            message=f"Conversion failed: {str(e)}"
        )

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        return start_conversion()
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def start_conversion():
    file = request.files.get('file') or request.files.get('document')

    if not file:
        return jsonify({'success': False, 'error': 'No file uploaded'}), 400

    if file.filename == '':
        return jsonify({'success': False, 'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({
            'success': False,
            'error': f'Unsupported file type. Allowed: {", ".join(Config.ALLOWED_EXTENSIONS)}'
        }), 400

    # Read conversion mode ('editable' or 'high_fidelity')
    mode = request.form.get('mode', Config.DEFAULT_MODE).strip().lower()
    if mode not in ['editable', 'high_fidelity']:
        mode = 'editable'

    job_id = uuid.uuid4().hex
    safe_name = secure_filename(file.filename) or "upload"
    unique_upload_name = f"{job_id}_{safe_name}"
    uploaded_path = os.path.join(Config.UPLOAD_FOLDER, unique_upload_name)

    file.save(uploaded_path)

    # Initialize job in manager
    job_manager.create_job(
        job_id=job_id,
        original_filename=file.filename,
        file_path=uploaded_path,
        mode=mode
    )

    # Dispatch to background worker
    job_manager.submit_task(run_conversion_job, job_id)

    return jsonify({
        'success': True,
        'job_id': job_id,
        'mode': mode,
        'message': 'File uploaded and queued for visual document reconstruction.'
    }), 202

@app.route('/status/<job_id>', methods=['GET'])
def get_job_status(job_id):
    job = job_manager.get_job(job_id)
    if not job:
        return jsonify({'success': False, 'error': 'Job not found'}), 404

    return jsonify({
        'success': True,
        'job_id': job['job_id'],
        'status': job['status'],
        'mode': job.get('mode', 'editable'),
        'progress': job['progress'],
        'message': job['message'],
        'total_pages': job.get('total_pages', 1),
        'current_page': job.get('current_page', 0),
        'download_url': job.get('download_url'),
        'error': job.get('error')
    })

@app.route('/download/<job_id_or_filename>', methods=['GET'])
def download_file(job_id_or_filename):
    # 1. Check if parameter matches an active job_id
    job = job_manager.get_job(job_id_or_filename)
    if job and job.get('docx_path') and os.path.exists(job['docx_path']):
        return send_file(
            job['docx_path'],
            as_attachment=True,
            download_name=job.get('docx_filename', 'converted_document.docx'),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    # 2. Check direct filename in OUTPUT_FOLDER for backward compatibility
    file_path = os.path.join(Config.OUTPUT_FOLDER, secure_filename(job_id_or_filename))
    if os.path.exists(file_path):
        return send_from_directory(Config.OUTPUT_FOLDER, os.path.basename(file_path), as_attachment=True)

    return jsonify({'success': False, 'error': 'Requested file not found or still processing.'}), 404

if __name__ == '__main__':
    try:
        app.run(debug=True, port=5000)
    finally:
        scheduler.shutdown()