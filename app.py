import asyncio
import json
import shutil
import subprocess
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")
executor = ThreadPoolExecutor(max_workers=2)

UPLOAD_DIR = Path("/app/uploads")
OUTPUT_DIR = Path("/app/output")
JOBS_FILE = Path("/app/output/jobs.json")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

_lock = threading.Lock()
ALLOWED = {".stl", ".3mf", ".obj", ".amf", ".igs", ".iges"}


def _load_jobs() -> dict:
    if JOBS_FILE.exists():
        try:
            with open(JOBS_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_jobs(jobs: dict) -> None:
    with open(JOBS_FILE, "w") as f:
        json.dump(jobs, f)


jobs: dict = _load_jobs()


def _convert(job_id: str, input_path: Path, output_path: Path) -> None:
    try:
        result = subprocess.run(
            ["python", "/app/2STEP-Converter.py", str(input_path), str(output_path)],
            capture_output=True,
            text=True,
            timeout=300,
            stdin=subprocess.DEVNULL,
        )
        # Script exits non-zero due to EOFError on "Press Enter to exit" prompt —
        # use output file existence as the success signal instead.
        if output_path.exists():
            jobs[job_id]["status"] = "done"
        else:
            jobs[job_id]["status"] = "error"
            jobs[job_id]["error"] = (result.stderr or result.stdout or "Conversion failed").strip()
    except subprocess.TimeoutExpired:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"] = "Timed out after 300s"
    except Exception as exc:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"] = str(exc)
    finally:
        with _lock:
            _save_jobs(jobs)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html")


@app.get("/jobs")
async def list_jobs():
    sorted_jobs = sorted(jobs.items(), key=lambda x: x[1].get("created_at", 0), reverse=True)
    return [
        {
            "job_id": jid,
            "filename": j["filename"],
            "status": j["status"],
            "error": j["error"],
        }
        for jid, j in sorted_jobs
    ]


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    suffix = Path(file.filename).suffix.lower()
    if suffix not in ALLOWED:
        raise HTTPException(400, f"Unsupported format: {suffix}")

    job_id = str(uuid.uuid4())
    input_path = UPLOAD_DIR / f"{job_id}{suffix}"
    output_path = OUTPUT_DIR / f"{job_id}.stp"

    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    with _lock:
        jobs[job_id] = {
            "status": "processing",
            "filename": file.filename,
            "output": str(output_path),
            "error": None,
            "created_at": time.time(),
        }
        _save_jobs(jobs)

    loop = asyncio.get_running_loop()
    loop.run_in_executor(executor, _convert, job_id, input_path, output_path)
    return {"job_id": job_id}


@app.get("/status/{job_id}")
async def status(job_id: str):
    if job_id not in jobs:
        raise HTTPException(404, "Not found")
    j = jobs[job_id]
    return {"status": j["status"], "error": j["error"]}


@app.get("/download/{job_id}")
async def download(job_id: str):
    if job_id not in jobs:
        raise HTTPException(404, "Not found")
    j = jobs[job_id]
    if j["status"] != "done":
        raise HTTPException(400, "Not ready")
    path = Path(j["output"])
    if not path.exists():
        raise HTTPException(404, "File missing")
    stem = Path(j["filename"]).stem
    return FileResponse(str(path), filename=f"{stem}.stp", media_type="application/octet-stream")
