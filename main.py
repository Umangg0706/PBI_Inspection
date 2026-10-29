"""ASGI entry point for hosting (Render): uvicorn main:app --host 0.0.0.0 --port $PORT

Wraps the logic in pbix_inspector.py; the local drag-and-drop mode there is unchanged.
"""
import json
import os
import tempfile
import uuid
from collections import OrderedDict
from urllib.parse import unquote

from fastapi import FastAPI, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, JSONResponse, Response

import pbix_inspector as pi

app = FastAPI(title="PBIX Inspector")

MAX_UPLOAD = int(os.environ.get("MAX_UPLOAD_MB", "200")) * 1024 * 1024
MAX_KEPT = int(os.environ.get("MAX_KEPT_RESULTS", "20"))  # bound server memory
RESULTS = OrderedDict()


@app.get("/", response_class=HTMLResponse)
def index():
    return pi.PAGE


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


@app.post("/analyze")
async def analyze(request: Request):
    tmp = None
    try:
        fname = unquote(request.headers.get("X-Filename", "model.pbix"))
        fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(fname)[1] or ".pbix")
        size = 0
        with os.fdopen(fd, "wb") as f:
            async for chunk in request.stream():
                size += len(chunk)
                if size > MAX_UPLOAD:
                    return JSONResponse({"error": f"File too large (limit {MAX_UPLOAD >> 20} MB)"})
                f.write(chunk)
        base = os.path.splitext(os.path.basename(fname))[0]
        data, warnings = await run_in_threadpool(pi.inspect, tmp, base)
        fid = uuid.uuid4().hex
        RESULTS[fid] = {"data": data, "report": base, "warnings": warnings}
        while len(RESULTS) > MAX_KEPT:
            RESULTS.popitem(last=False)
        sheets = {n: {"headers": pi.SHEETS[n], "total": len(data[n]),
                      "rows": [[pi.clean(v) for v in r] for r in data[n][:pi.PREVIEW_ROWS]]} for n in pi.SHEETS}
        nvis = len({(v["page"], v["vid"]) for v in data["_visuals"]
                    if v["vtype"] not in ("Page Filters", "Report Filters")})
        stats = [["Tables", len(data["Tables"]), "Tables"], ["Columns", len(data["Columns"]), "Columns"],
                 ["Measures", len(data["Measures"]), "Measures"],
                 ["Relationships", len(data["Relationships"]), "Relationships"],
                 ["M Queries", len(data["M Queries"]), "M Queries"],
                 ["Visuals", nvis, "Visual Design"], ["Lineage rows", len(data["Lineage"]), "Lineage"],
                 ["Impact rows", len(data["Impact Analysis"]), "Impact Analysis"]]
        return Response(json.dumps({"id": fid, "sheets": sheets, "warnings": warnings, "stats": stats},
                                   default=str), media_type="application/json")
    except Exception as e:
        return JSONResponse({"error": f"{type(e).__name__}: {e}"})
    finally:
        if tmp and os.path.exists(tmp):
            os.remove(tmp)


@app.get("/download/{fid}")
def download(fid: str, sheets: str = None, summary: str = "1"):
    item = RESULTS.get(fid)
    if not item:
        return Response("not found (results expire when the server restarts)", status_code=404)
    chosen = [x for x in sheets.split(",") if x in pi.SHEETS] if sheets is not None else None
    body = pi.build_workbook(item["data"], chosen, item["report"], item["warnings"], summary != "0")
    return Response(body, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{item["report"]}_inspection.xlsx"'})
