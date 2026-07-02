# tex2word — Chuyển LaTeX thành Word

Web app **miễn phí**: dán các khối câu hỏi `\begin{ex}...\end{ex}` → tải về file `.docx`
có công thức Toán (OMML — sửa được trong Word) + hình TikZ/tabular.

## Bản chạy trực tuyến

- **Trang web (frontend)**: https://tex2word.vercel.app
- **API (backend)**: https://tex2word-api.onrender.com

Hai chế độ xuất:
- **Equation** — công thức chuyển sang Equation của Word (chỉnh sửa trực tiếp).
- **LaTeX** — giữ nguyên `$..$` / `$$..$$` để dán qua MathType (file có hậu tố `-latex`).

Đề có hình **TikZ/tabular** sẽ hiện modal tiến trình render từng hình; đề không có hình tải thẳng.

## Cấu trúc

```
tex2word/
  backend/          FastAPI + docx_exporter (Python)
    app.py
    docx_exporter.py     (bản sao từ D:\TOAN10 — đã sửa để chạy Linux)
    MML2OMML.XSL         (XSLT Microsoft cho MathML → OMML)
    requirements.txt
    render.yaml          (cấu hình deploy Render)
  frontend/         Vite + React 19 + TypeScript + Tailwind (CDN)
    src/App.tsx
    package.json
    .env.example
```

## Chạy cục bộ

**Backend** (yêu cầu Python 3.11+):

```bash
cd backend
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

Mở http://127.0.0.1:8000/ — phải thấy `omml_available: true`.

**Frontend**:

```bash
cd frontend
npm install
npm run dev
```

Mở http://localhost:5175.

## Deploy

**Backend → Render** (free tier):

1. Đẩy thư mục `backend/` lên GitHub.
2. Render → New → Web Service → connect repo → chọn `render.yaml` sẽ tự đọc.
3. Sau khi deploy: mở `https://<tên-service>.onrender.com/` để chắc chắn OMML available.

**Frontend → Vercel**:

1. Đẩy thư mục `frontend/` lên GitHub.
2. Vercel → import repo → framework: Vite.
3. Thêm biến môi trường `VITE_API_URL` = URL của backend trên Render.

## Ghi chú kỹ thuật

- `MML2OMML.XSL` là file XSLT của Microsoft Office (copy từ `C:\Program Files\Microsoft
  Office\root\Office16\`). Đặt cạnh `docx_exporter.py` — hàm `_find_mml2omml_xsl` tự tìm.
- TikZ/tabular render qua API sẵn: `compile-tikz-code.onrender.com` — cần Internet.
- Free tier Render **cold start** ~30 giây khi lâu không dùng. Có thể ping định kỳ để
  giữ ấm hoặc nâng gói nếu cần.
