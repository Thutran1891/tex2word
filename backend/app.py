# -*- coding: utf-8 -*-
"""
FastAPI wrapper cho docx_exporter — nhận LaTeX (khối \\begin{ex}...\\end{ex}),
trả file .docx tải về.
"""

import os
import tempfile
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from docx_exporter import export_questions_to_docx, omml_available

app = FastAPI(title="tex2word", version="1.0.0")

# Cho phép mọi origin (bản free không kiểm soát người dùng). Có thể siết lại sau.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


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


@app.get("/")
def health():
    return {
        "app": "tex2word",
        "status": "ok",
        "omml_available": omml_available(),
    }


@app.post("/convert")
def convert(req: ConvertRequest):
    text = (req.text or "").strip()
    if not text:
        raise HTTPException(status_code=400, detail="Thiếu nội dung LaTeX.")
    if "\\begin{ex}" not in text:
        raise HTTPException(
            status_code=400,
            detail="Không thấy khối \\begin{ex}...\\end{ex} trong nội dung.",
        )

    meta = (req.meta.model_dump() if req.meta else {})

    # Ghi ra file tạm rồi trả về; xoá sau khi gửi (background).
    fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    mode = (req.mode or "equation").lower()
    if mode not in ("equation", "latex"):
        mode = "equation"
    try:
        export_questions_to_docx(text, tmp_path, meta=meta, math_mode=mode)
    except Exception as e:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise HTTPException(status_code=500, detail=f"Lỗi khi xuất Word: {e}")

    fname = (req.filename or "de_thi.docx").strip() or "de_thi.docx"
    if not fname.lower().endswith(".docx"):
        fname += ".docx"

    return FileResponse(
        tmp_path,
        media_type=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        filename=fname,
    )
