import os
import time
import threading
from concurrent.futures import ThreadPoolExecutor
from config import Config

class JobManager:
    """
    Thread-safe background job manager for asynchronous visual document reconstruction.
    Tracks status, progress percentage, detailed status messages, mode, and output links.
    """
    def __init__(self, max_workers: int = 4):
        self._lock = threading.Lock()
        self._jobs: dict[str, dict] = {}
        self._executor = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="DocuShiftWorker")

    def create_job(self, job_id: str, original_filename: str, file_path: str, mode: str = "editable") -> dict:
        job_data = {
            "job_id": job_id,
            "status": "pending",           # pending, processing, completed, failed
            "progress": 0,                  # 0 to 100
            "mode": mode,                   # 'editable' or 'high_fidelity'
            "message": "Upload complete. Job queued for visual reconstruction...",
            "original_filename": original_filename,
            "file_path": file_path,
            "total_pages": 1,
            "current_page": 0,
            "download_url": None,
            "docx_path": None,
            "docx_filename": None,
            "error": None,
            "created_at": time.time(),
            "updated_at": time.time(),
        }
        with self._lock:
            self._jobs[job_id] = job_data
        return job_data

    def get_job(self, job_id: str) -> dict | None:
        with self._lock:
            job = self._jobs.get(job_id)
            return dict(job) if job else None

    def update_job(self, job_id: str, **kwargs):
        with self._lock:
            if job_id in self._jobs:
                self._jobs[job_id].update(kwargs)
                self._jobs[job_id]["updated_at"] = time.time()

    def submit_task(self, fn, *args, **kwargs):
        return self._executor.submit(fn, *args, **kwargs)

    def cleanup_expired_jobs(self, max_age_seconds: int = 3600):
        """
        Removes in-memory job records older than max_age_seconds.
        """
        now = time.time()
        with self._lock:
            expired_ids = [
                jid for jid, j in self._jobs.items() 
                if now - j.get("created_at", now) > max_age_seconds
            ]
            for jid in expired_ids:
                del self._jobs[jid]

# Global singleton instance
job_manager = JobManager(max_workers=Config.MAX_WORKERS)
