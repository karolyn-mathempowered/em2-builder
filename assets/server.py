"""
EM2 Module Builder — backend API.
Wraps build_module.py behind a single endpoint that a web app (e.g. a Lovable
frontend) can call. Accepts the 4 source .pptx files and returns the finished deck.

Run locally:   uvicorn server:app --reload --port 8000
Deploy:        see Dockerfile (installs LibreOffice + poppler + python deps)
"""
import os, tempfile
from types import SimpleNamespace
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import build_module

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")

app = FastAPI(title="EM2 Module Builder")
# Allow the Lovable app (and local dev) to call this API from the browser.
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=False,
                   allow_methods=["*"], allow_headers=["*"])

def _save(upload: UploadFile, folder: str) -> str:
    path = os.path.join(folder, os.path.basename(upload.filename or "file.pptx"))
    with open(path, "wb") as f:
        f.write(upload.file.read())
    return path

def _toint(x):
    try: return int(str(x).strip())
    except Exception: return None

@app.get("/health")
def health():
    return {"ok": True, "service": "em2-module-builder"}

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
):
    work = tempfile.mkdtemp()
    try:
        def pick(pptx, pdf, label, required):
            up = pptx or pdf
            if up is None:
                if required:
                    raise HTTPException(422, f"{label}: attach a .pptx or a .pdf.")
                return None
            return _save(up, work)

        mt = pick(mathtalks, mathtalks_pdf, "Math Talks", True)
        dr = pick(dares, dares_pdf, "DAREs", True)
        ag = pick(answerguides, answerguides_pdf, "DARE Answer Guides", False) or dr
        so = pick(sorts, sorts_pdf, "Sorts", False)
        out = os.path.join(work, "module.pptx")
        ns = SimpleNamespace(mathtalks=mt, sorts=so, dares=dr, answerguides=ag,
                             out=out, title=(title or None), topics=None,
                             grade=_toint(grade), module=_toint(module), assets=ASSETS)
        build_module.build(ns)
        if not os.path.exists(out):
            raise HTTPException(500, "Deck was not produced.")
        fname = ((title or "EM2_Module").strip().replace(" ", "_") or "EM2_Module") + ".pptx"
        return FileResponse(out, filename=fname,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
