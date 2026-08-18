"""
Daily Presentation Slides Builder — backend API.

Endpoints
  GET  /health
  POST /generate                 synchronous build (unchanged, kept for small builds)
  POST /generate-async           202 {"job_id", "status"} — builds in the background
  GET  /jobs/{job_id}            job record with honest, source-backed progress
  GET  /jobs/{job_id}/download   streams the finished .pptx (409 until done)
  DELETE /jobs/{job_id}          cancels/drops a job

Run locally:   uvicorn server:app --reload --port 8000
"""
import os, glob, shutil, tempfile, threading, time, uuid, subprocess
from types import SimpleNamespace
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import build_module

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
DECKS = "/tmp/decks"
JOB_TTL_SECONDS = 60 * 60

app = FastAPI(title="Daily Presentation Slides Builder")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=False,
                   allow_methods=["*"], allow_headers=["*"])

# ---------------------------------------------------------------- job store
JOBS = {}                       # job_id -> dict
JOBS_LOCK = threading.Lock()
BUILD_SLOT = threading.Semaphore(1)   # only one build at a time (512 MB instance)
QUEUE = []                      # job ids waiting for the slot


@app.on_event("startup")
def _startup():
    # Anything left in /tmp/decks belongs to a previous instance and is unreachable.
    shutil.rmtree(DECKS, ignore_errors=True)
    os.makedirs(DECKS, exist_ok=True)


def _save(upload: UploadFile, folder: str) -> str:
    path = os.path.join(folder, os.path.basename(upload.filename or "file"))
    with open(path, "wb") as f:
        shutil.copyfileobj(upload.file, f)
    return path


def _toint(x):
    try: return int(str(x).strip())
    except Exception: return None


def _pdf_pages(path: str) -> int:
    """Page count of a PDF (0 if unknown) — used for an honest page total."""
    if not path or not path.lower().endswith(".pdf"):
        return 0
    try:
        out = subprocess.run(["pdfinfo", path], capture_output=True, text=True, timeout=60).stdout
        for line in out.splitlines():
            if line.lower().startswith("pages:"):
                return int(line.split(":", 1)[1].strip())
    except Exception:
        pass
    return 0


def _filename(title, lesson_from, lesson_to):
    fname = ((title or "Daily_Slides").strip().replace(" ", "_") or "Daily_Slides")
    if lesson_from or lesson_to:
        fname += f"_Lessons_{lesson_from or '1'}-{lesson_to or 'end'}"
    return fname + ".pptx"


@app.get("/health")
def health():
    with JOBS_LOCK:
        running = sum(1 for j in JOBS.values() if j["status"] == "running")
        queued = sum(1 for j in JOBS.values() if j["status"] == "queued")
    return {"ok": True, "service": "daily-slides-builder", "running": running, "queued": queued}


# ---------------------------------------------------------------- sync build
@app.post("/generate")
async def generate(
    mathtalks: Optional[UploadFile] = File(None),
    dares: Optional[UploadFile] = File(None),
    answerguides: Optional[UploadFile] = File(None),
    sorts: Optional[UploadFile] = File(None),
    mathtalks_pdf: Optional[UploadFile] = File(None),
    dares_pdf: Optional[UploadFile] = File(None),
    answerguides_pdf: Optional[UploadFile] = File(None),
    sorts_pdf: Optional[UploadFile] = File(None),
    tasks_pdf: Optional[UploadFile] = File(None),
    games_pdf: Optional[UploadFile] = File(None),
    title: Optional[str] = Form(None),
    grade: Optional[str] = Form(None),
    module: Optional[str] = Form(None),
    sorts_lessons: Optional[str] = Form(None),
    game_links: Optional[str] = Form(None),
    lesson_from: Optional[str] = Form(None),
    lesson_to: Optional[str] = Form(None),
):
    work = tempfile.mkdtemp()
    try:
        paths = _collect(work, mathtalks, dares, answerguides, sorts,
                         mathtalks_pdf, dares_pdf, answerguides_pdf, sorts_pdf,
                         tasks_pdf, games_pdf)
        out = os.path.join(work, "module.pptx")
        with BUILD_SLOT:
            build_module.build(_args(paths, out, title, grade, module,
                                     sorts_lessons, game_links, lesson_from, lesson_to))
        if not os.path.exists(out):
            raise HTTPException(500, "Deck was not produced.")
        return FileResponse(out, filename=_filename(title, lesson_from, lesson_to),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


def _collect(work, mathtalks, dares, answerguides, sorts,
             mathtalks_pdf, dares_pdf, answerguides_pdf, sorts_pdf,
             tasks_pdf, games_pdf):
    def pick(pptx, pdf, label, required):
        up = pptx or pdf
        if up is None:
            if required:
                raise HTTPException(422, f"{label}: attach a .pptx or a .pdf.")
            return None
        return _save(up, work)
    return {
        "mathtalks": pick(mathtalks, mathtalks_pdf, "Math Talks", True),
        "dares": pick(dares, dares_pdf, "DAREs", True),
        "answerguides": pick(answerguides, answerguides_pdf, "DARE Answer Guides", False),
        "sorts": pick(sorts, sorts_pdf, "Sorts", False),
        "tasks": pick(None, tasks_pdf, "Math Tasks", False),
        "games": pick(None, games_pdf, "Games", False),
    }


def _args(paths, out, title, grade, module, sorts_lessons, game_links, lesson_from, lesson_to):
    return SimpleNamespace(
        mathtalks=paths["mathtalks"], sorts=paths["sorts"], dares=paths["dares"],
        answerguides=paths["answerguides"], tasks=paths["tasks"], games=paths["games"],
        game_links=(game_links or None), sorts_lessons=(sorts_lessons or None),
        out=out, title=(title or None), topics=None,
        grade=(grade or None), module=(module or None),
        lesson_from=(lesson_from or None), lesson_to=(lesson_to or None),
        assets=ASSETS)


# --------------------------------------------------------------- async build
@app.post("/generate-async", status_code=202)
async def generate_async(
    background: BackgroundTasks,
    mathtalks: Optional[UploadFile] = File(None),
    dares: Optional[UploadFile] = File(None),
    answerguides: Optional[UploadFile] = File(None),
    sorts: Optional[UploadFile] = File(None),
    mathtalks_pdf: Optional[UploadFile] = File(None),
    dares_pdf: Optional[UploadFile] = File(None),
    answerguides_pdf: Optional[UploadFile] = File(None),
    sorts_pdf: Optional[UploadFile] = File(None),
    tasks_pdf: Optional[UploadFile] = File(None),
    games_pdf: Optional[UploadFile] = File(None),
    title: Optional[str] = Form(None),
    grade: Optional[str] = Form(None),
    module: Optional[str] = Form(None),
    sorts_lessons: Optional[str] = Form(None),
    game_links: Optional[str] = Form(None),
    lesson_from: Optional[str] = Form(None),
    lesson_to: Optional[str] = Form(None),
):
    job_id = str(uuid.uuid4())
    work = os.path.join(DECKS, job_id)
    os.makedirs(work, exist_ok=True)
    paths = _collect(work, mathtalks, dares, answerguides, sorts,
                     mathtalks_pdf, dares_pdf, answerguides_pdf, sorts_pdf,
                     tasks_pdf, games_pdf)
    total_pages = sum(_pdf_pages(p) for p in paths.values() if p)

    job = {
        "job_id": job_id, "status": "queued", "stage": "queued",
        "current": 0, "total": total_pages,
        "lessons_done": 0, "lessons_total": None, "slides": None,
        "started_at": time.time(), "finished_at": None,
        "error": None, "warnings": [], "queue_position": None,
        "filename": _filename(title, lesson_from, lesson_to),
        "work": work, "deck": None, "cancelled": False,
    }
    with JOBS_LOCK:
        _sweep_locked()
        JOBS[job_id] = job
        QUEUE.append(job_id)
        job["queue_position"] = len(QUEUE)

    args = _args(paths, os.path.join(work, "module.pptx"), title, grade, module,
                 sorts_lessons, game_links, lesson_from, lesson_to)
    threading.Thread(target=_run_job, args=(job_id, args), daemon=True).start()
    return {"job_id": job_id, "status": "queued", "total": total_pages}


def _run_job(job_id, args):
    job = JOBS.get(job_id)
    if job is None:
        return
    BUILD_SLOT.acquire()
    try:
        with JOBS_LOCK:
            if job_id in QUEUE:
                QUEUE.remove(job_id)
            for i, qid in enumerate(QUEUE):
                q = JOBS.get(qid)
                if q: q["queue_position"] = i + 1
        if job.get("cancelled"):
            job["status"] = "error"; job["error"] = "cancelled"
            return

        job["status"] = "running"
        job["stage"] = "starting"
        job["queue_position"] = None
        job["started_at"] = time.time()

        def report(**kw):
            if job.get("cancelled"):
                raise RuntimeError("cancelled")
            if "pages_rendered" in kw:
                job["current"] = job.get("current", 0) + int(kw["pages_rendered"])
                if job["total"] and job["current"] > job["total"]:
                    job["total"] = job["current"]
            for k in ("stage", "lessons_total", "lessons_done", "slides"):
                if k in kw and kw[k] is not None:
                    job[k] = kw[k]

        build_module.PROGRESS = report
        try:
            build_module.build(args)
        finally:
            build_module.PROGRESS = None

        out = args.out
        if not os.path.exists(out):
            raise RuntimeError("Deck was not produced.")
        job["deck"] = out
        job["bytes"] = os.path.getsize(out)
        job["status"] = "done"
        job["stage"] = "done"
        job["finished_at"] = time.time()
    except Exception as e:
        job["status"] = "error"
        job["stage"] = "failed"
        job["error"] = "cancelled" if job.get("cancelled") else str(e)
        job["finished_at"] = time.time()
    finally:
        BUILD_SLOT.release()


def _public(job):
    elapsed = int((job.get("finished_at") or time.time()) - job["started_at"])
    return {
        "job_id": job["job_id"], "status": job["status"], "stage": job["stage"],
        "current": job.get("current", 0), "total": job.get("total") or 0,
        "lessons_done": job.get("lessons_done", 0), "lessons_total": job.get("lessons_total"),
        "slides": job.get("slides"), "bytes": job.get("bytes"),
        "elapsed_seconds": elapsed, "error": job.get("error"),
        "warnings": job.get("warnings", []), "queue_position": job.get("queue_position"),
        "filename": job.get("filename"),
    }


def _sweep_locked():
    now = time.time()
    for jid, j in list(JOBS.items()):
        fin = j.get("finished_at")
        if fin and now - fin > JOB_TTL_SECONDS:
            shutil.rmtree(j.get("work") or "", ignore_errors=True)
            JOBS.pop(jid, None)


@app.get("/jobs/{job_id}")
def job_status(job_id: str):
    with JOBS_LOCK:
        _sweep_locked()
        job = JOBS.get(job_id)
    if not job:
        return JSONResponse(status_code=404, content={"error": "job_not_found"})
    return _public(job)


@app.get("/jobs/{job_id}/download")
def job_download(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        return JSONResponse(status_code=404, content={"error": "job_not_found"})
    if job["status"] != "done" or not job.get("deck"):
        return JSONResponse(status_code=409, content={"error": "not_ready", "status": job["status"]})
    return FileResponse(job["deck"], filename=job.get("filename") or "Daily_Slides.pptx",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")


@app.delete("/jobs/{job_id}")
def job_cancel(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        return JSONResponse(status_code=404, content={"error": "job_not_found"})
    job["cancelled"] = True
    if job["status"] in ("queued", "running"):
        job["status"] = "error"
        job["stage"] = "cancelled"
        job["error"] = "cancelled"
        job["finished_at"] = time.time()
    with JOBS_LOCK:
        if job_id in QUEUE:
            QUEUE.remove(job_id)
    # Keep the record so a later poll reports "cancelled" instead of a bare 404.
    return {"ok": True, "cancelled": job_id}
