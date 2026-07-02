# -*- coding: utf-8 -*-
"""
FastAPI wrapper cho docx_exporter — nhận LaTeX (khối \\begin{ex}...\\end{ex}),
trả file .docx tải về.

Hai endpoint:
- POST /convert         : trả thẳng file .docx (dùng khi đề KHÔNG có TikZ — nhanh).
- POST /convert-stream  : trả dòng NDJSON tiến trình render từng hình rồi file base64
                          (dùng khi đề CÓ TikZ/tabular để hiện thanh tiến trình).
"""

import os
import json
import queue
import base64
import tempfile
import threading
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel

from docx_exporter import export_questions_to_docx, omml_available

app = FastAPI(title="tex2word", version="1.1.0")

# Cho phép mọi origin (bản free không kiểm soát người dùng). Có thể siết lại sau.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


class Meta(BaseModel):
    school: Optional[str] = ""
    title: Optional[str] = ""
    subject: Optional[str] = ""
    duration: Optional[str] = ""
    class_name: Optional[str] = ""
    made: Optional[str] = ""


class ConvertRequest(BaseModel):
    text: str
    meta: Optional[Meta] = None
    filename: Optional[str] = "de_thi.docx"
    # "equation" (mặc định — OMML, sửa được trong Word) hoặc "latex" (giữ $..$ cho MathType).
    mode: Optional[str] = "equation"


def _prepare(req: ConvertRequest):
    """Kiểm tra đầu vào chung cho cả 2 endpoint. Trả (text, meta, mode, fname)."""
    text = (req.text or "").strip()
    if not text:
        raise HTTPException(status_code=400, detail="Thiếu nội dung LaTeX.")
    if "\\begin{ex}" not in text:
        raise HTTPException(
            status_code=400,
            detail="Không thấy khối \\begin{ex}...\\end{ex} trong nội dung.",
        )
    meta = req.meta.model_dump() if req.meta else {}
    mode = (req.mode or "equation").lower()
    if mode not in ("equation", "latex"):
        mode = "equation"
    fname = (req.filename or "de_thi.docx").strip() or "de_thi.docx"
    if not fname.lower().endswith(".docx"):
        fname += ".docx"
    return text, meta, mode, fname


@app.get("/")
def health():
    return {
        "app": "tex2word",
        "status": "ok",
        "omml_available": omml_available(),
    }


@app.post("/convert")
def convert(req: ConvertRequest):
    text, meta, mode, fname = _prepare(req)

    fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        export_questions_to_docx(text, tmp_path, meta=meta, math_mode=mode)
    except Exception as e:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise HTTPException(status_code=500, detail=f"Lỗi khi xuất Word: {e}")

    return FileResponse(tmp_path, media_type=DOCX_MIME, filename=fname)


@app.post("/convert-stream")
def convert_stream(req: ConvertRequest):
    """Chạy export trong 1 luồng, đẩy tiến trình render hình qua hàng đợi, và stream
    ra client dạng NDJSON. Dòng cuối chứa file .docx đã mã hoá base64."""
    text, meta, mode, fname = _prepare(req)

    def generate():
        q: "queue.Queue" = queue.Queue()
        result: dict = {}

        def progress_cb(done, total, message):
            q.put({"type": "progress", "done": done, "total": total, "message": message})

        def worker():
            fd, tmp_path = tempfile.mkstemp(suffix=".docx")
            os.close(fd)
            try:
                export_questions_to_docx(
                    text, tmp_path, meta=meta, progress_cb=progress_cb, math_mode=mode
                )
                with open(tmp_path, "rb") as f:
                    result["data"] = base64.b64encode(f.read()).decode("ascii")
            except Exception as e:  # noqa: BLE001 - báo lỗi về client qua stream
                result["error"] = str(e)
            finally:
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
                q.put(None)  # sentinel: worker xong

        t = threading.Thread(target=worker, daemon=True)
        t.start()

        # Bơm các sự kiện tiến trình ra client cho tới khi worker báo xong.
        while True:
            item = q.get()
            if item is None:
                break
            yield json.dumps(item, ensure_ascii=False) + "\n"
        t.join()

        if "error" in result:
            yield json.dumps(
                {"type": "error", "detail": f"Lỗi khi xuất Word: {result['error']}"},
                ensure_ascii=False,
            ) + "\n"
        else:
            yield json.dumps(
                {"type": "done", "filename": fname, "data": result["data"]},
                ensure_ascii=False,
            ) + "\n"

    # X-Accel-Buffering: no để tránh proxy gom buffer, đảm bảo tiến trình về ngay.
    return StreamingResponse(
        generate(),
        media_type="application/x-ndjson",
        headers={"X-Accel-Buffering": "no", "Cache-Control": "no-cache"},
    )
