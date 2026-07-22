"""
EM2 Module Builder — backend API.
Wraps build_module.py behind a single endpoint that a web app (e.g. a Lovable
frontend) can call. Accepts the 4 source .pptx files and returns the finished deck.

Run locally:   uvicorn server:app --reload --port 8000
Deploy:        see Dockerfile (installs LibreOffice + poppler + python deps)
"""
import os, tempfile
from types import SimpleNamespace
from typing import Optional, List
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import build_module
import build_kit

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
    mathtalks: UploadFile = File(...),
    dares: UploadFile = File(...),
    answerguides: UploadFile = File(...),
    sorts: Optional[UploadFile] = File(None),
    # Optional pre-rendered PDFs. When supplied, the server skips LibreOffice for
    # that file and rasterizes the PDF instead (much lower memory use).
    mathtalks_pdf: Optional[UploadFile] = File(None),
    dares_pdf: Optional[UploadFile] = File(None),
    answerguides_pdf: Optional[UploadFile] = File(None),
    sorts_pdf: Optional[UploadFile] = File(None),
    title: Optional[str] = Form(None),
    grade: Optional[str] = Form(None),
    module: Optional[str] = Form(None),
):
    work = tempfile.mkdtemp()
    try:
        mt = _save(mathtalks, work)
        dr = _save(dares, work)
        ag = _save(answerguides, work)
        so = _save(sorts, work) if sorts is not None else None
        mt_pdf = _save(mathtalks_pdf, work) if mathtalks_pdf is not None else None
        dr_pdf = _save(dares_pdf, work) if dares_pdf is not None else None
        ag_pdf = _save(answerguides_pdf, work) if answerguides_pdf is not None else None
        so_pdf = _save(sorts_pdf, work) if sorts_pdf is not None else None
        out = os.path.join(work, "module.pptx")
        ns = SimpleNamespace(mathtalks=mt, sorts=so, dares=dr, answerguides=ag,
                             mathtalks_pdf=mt_pdf, sorts_pdf=so_pdf,
                             dares_pdf=dr_pdf, answerguides_pdf=ag_pdf,
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


@app.get("/debug/fetch")
def debug_fetch(id: str):
    import requests
    r = requests.get(
        f"https://drive.google.com/uc?export=download&id={id}",
        allow_redirects=True,
    )
    return {
        "status": r.status_code,
        "ctype": r.headers.get("content-type"),
        "bytes": len(r.content),
        "head": r.content[:8].decode("latin-1"),
    }


class KitItem(BaseModel):
    name: Optional[str] = None
    standard: Optional[str] = None
    url: str


class KitRequest(BaseModel):
    title: Optional[str] = None
    grade: Optional[str] = None
    resource_type: str = "dare"
    items: List[KitItem]


@app.post("/generate-from-kit")
def generate_from_kit(req: KitRequest):
    work = tempfile.mkdtemp()
    try:
        out = os.path.join(work, "kit.pptx")
        build_kit.build_from_kit(
            {
                "title": req.title,
                "grade": req.grade,
                "resource_type": req.resource_type,
                "items": [i.dict() for i in req.items],
            },
            out,
            ASSETS,
        )
        if not os.path.exists(out):
            raise HTTPException(500, "Deck was not produced.")
        fname = ((req.title or "EM2_Kit").strip().replace(" ", "_") or "EM2_Kit") + ".pptx"
        return FileResponse(
            out,
            filename=fname,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
