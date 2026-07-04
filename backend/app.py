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

from docx_exporter import (
    export_questions_to_docx,
    export_questions_to_html,
    omml_available,
    compile_tikz_b64,
)

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
    # True: in tiêu đề + các đề mục "Phần I/II/III/IV"; False: chỉ in câu hỏi.
    show_header: Optional[bool] = True
    # True: đầy đủ (đánh dấu đáp án + đáp số TLN + Lời giải); False: chỉ có đề bài.
    show_answers: Optional[bool] = True


class TikzRequest(BaseModel):
    source: str
    transparent: Optional[bool] = True


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


@app.post("/tikz-png")
def tikz_png(req: TikzRequest):
    """Biên dịch 1 đoạn mã TikZ thành PNG. Trả {image_base64} (chuỗi base64 của PNG).
    Chấp nhận cả mã có/không bọc \\begin{tikzpicture} — server TikZ tự lo (mode=auto)."""
    src = (req.source or "").strip()
    if not src:
        raise HTTPException(status_code=400, detail="Thiếu mã TikZ.")

    # Dùng chung helper có THỬ LẠI (server Render hay ngủ nên request đầu dễ hỏng).
    b64, reason, is_compile_error = compile_tikz_b64(src, transparent=bool(req.transparent))
    if b64:
        return {"image_base64": b64}

    if is_compile_error:
        raise HTTPException(status_code=422, detail=f"Biên dịch TikZ lỗi:\n{reason}")
    raise HTTPException(
        status_code=502,
        detail=f"Không gọi được server vẽ TikZ (đã thử lại vài lần, có thể đang khởi động, thử lại sau ~30s): {reason}",
    )


@app.post("/convert")
def convert(req: ConvertRequest):
    text, meta, mode, fname = _prepare(req)

    fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        export_questions_to_docx(
            text, tmp_path, meta=meta, math_mode=mode,
            show_header=bool(req.show_header), show_answers=bool(req.show_answers),
        )
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
                    text, tmp_path, meta=meta, progress_cb=progress_cb,
                    math_mode=mode, show_header=bool(req.show_header),
                    show_answers=bool(req.show_answers),
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


@app.post("/convert-html")
def convert_html(req: ConvertRequest):
    """Trả HTML (một tài liệu hoàn chỉnh có MathJax) để client in ra PDF. Dùng khi đề
    KHÔNG có TikZ/tabular (nhanh, không cần thanh tiến trình)."""
    text, meta, _mode, fname = _prepare(req)
    pdf_name = fname[:-5] + ".pdf"
    try:
        html = export_questions_to_html(
            text, meta=meta, show_header=bool(req.show_header),
            show_answers=bool(req.show_answers), pdf_name=pdf_name,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi tạo HTML: {e}")
    return {"html": html}


@app.post("/convert-html-stream")
def convert_html_stream(req: ConvertRequest):
    """Như /convert-stream nhưng dòng cuối chứa chuỗi HTML (field 'html') thay vì file
    .docx — dùng khi đề CÓ hình để hiện thanh tiến trình render TikZ."""
    text, meta, _mode, fname = _prepare(req)
    pdf_name = fname[:-5] + ".pdf"

    def generate():
        q: "queue.Queue" = queue.Queue()
        result: dict = {}

        def progress_cb(done, total, message):
            q.put({"type": "progress", "done": done, "total": total, "message": message})

        def worker():
            try:
                result["html"] = export_questions_to_html(
                    text, meta=meta, progress_cb=progress_cb,
                    show_header=bool(req.show_header), show_answers=bool(req.show_answers),
                    pdf_name=pdf_name,
                )
            except Exception as e:  # noqa: BLE001 - báo lỗi về client qua stream
                result["error"] = str(e)
            finally:
                q.put(None)  # sentinel: worker xong

        t = threading.Thread(target=worker, daemon=True)
        t.start()

        while True:
            item = q.get()
            if item is None:
                break
            yield json.dumps(item, ensure_ascii=False) + "\n"
        t.join()

        if "error" in result:
            yield json.dumps(
                {"type": "error", "detail": f"Lỗi khi tạo HTML: {result['error']}"},
                ensure_ascii=False,
            ) + "\n"
        else:
            yield json.dumps(
                {"type": "done", "html": result["html"]},
                ensure_ascii=False,
            ) + "\n"

    return StreamingResponse(
        generate(),
        media_type="application/x-ndjson",
        headers={"X-Accel-Buffering": "no", "Cache-Control": "no-cache"},
    )
