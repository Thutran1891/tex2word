import { useState } from "react";

const API_URL = import.meta.env.VITE_API_URL || "http://localhost:8000";

const SAMPLE = String.raw`\begin{ex}
Cho tập hợp $A = \{1;2;3\}$ và $B = \{2;3;4\}$. Tập $A \cap B$ bằng:
\choice
{$\{1\}$}
{\True $\{2;3\}$}
{$\{1;2;3;4\}$}
{$\varnothing$}
\loigiai{Ta có $A \cap B = \{2;3\}$.}
\end{ex}`;

type Meta = {
  school: string;
  title: string;
  subject: string;
  duration: string;
  class_name: string;
  made: string;
};

const emptyMeta: Meta = {
  school: "TRƯỜNG THPT CÂY DƯƠNG",
  title: "ĐỀ ĐÁNH GIÁ THƯỜNG XUYÊN",
  subject: "Môn: Toán 10",
  duration: "20 phút",
  class_name: "10A",
  made: "1001",
};

export default function App() {
  const [text, setText] = useState<string>(SAMPLE);
  const [meta, setMeta] = useState<Meta>(emptyMeta);
  const [filename, setFilename] = useState<string>("de_thi.docx");
  const [busy, setBusy] = useState<"" | "equation" | "latex">("");
  const [err, setErr] = useState<string>("");

  function updateMeta<K extends keyof Meta>(key: K, value: Meta[K]) {
    setMeta((m) => ({ ...m, [key]: value }));
  }

  async function handleConvert(mode: "equation" | "latex") {
    setErr("");
    if (!text.trim()) {
      setErr("Cô ơi, chưa nhập nội dung LaTeX.");
      return;
    }
    if (!text.includes("\\begin{ex}")) {
      setErr("Không thấy khối \\begin{ex}...\\end{ex} trong nội dung.");
      return;
    }
    // Thêm hậu tố -latex vào tên file để phân biệt khi tải cả 2 bản.
    const outName = (() => {
      const base = filename.endsWith(".docx") ? filename.slice(0, -5) : filename;
      const suffix = mode === "latex" ? "-latex" : "";
      return `${base}${suffix}.docx`;
    })();
    setBusy(mode);
    try {
      const resp = await fetch(`${API_URL}/convert`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text, meta, filename: outName, mode }),
      });
      if (!resp.ok) {
        let msg = `Lỗi ${resp.status}`;
        try {
          const j = await resp.json();
          if (j?.detail) msg = j.detail;
        } catch {
          /* keep default */
        }
        throw new Error(msg);
      }
      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = outName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy("");
    }
  }

  const countEx = (text.match(/\\begin\{ex\}/g) || []).length;

  return (
    <div className="min-h-screen">
      <header className="bg-gradient-to-r from-indigo-600 to-violet-600 text-white">
        <div className="max-w-6xl mx-auto px-6 py-5">
          <div className="flex flex-wrap items-start justify-between gap-3">
            <div>
              <h1 className="text-2xl font-bold tracking-tight">
                Chuyển LaTeX thành Word
              </h1>
              <p className="text-indigo-100 text-sm mt-1">
                Dán các khối <code className="bg-white/20 px-1 rounded">\begin&#123;ex&#125;…\end&#123;ex&#125;</code>
                {" "}vào khung bên dưới, bấm nút để tải về file <b>.docx</b> (công thức Toán vẫn sửa được trong Word).
              </p>
            </div>
            <div className="text-right text-indigo-100 text-xs leading-relaxed shrink-0">
              <div className="font-semibold text-white">Trần Thị Kim Thu</div>
              <div>TRƯỜNG THPT CÂY DƯƠNG</div>
              <div>0397 58 43 58</div>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-6 grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Cột trái: form meta */}
        <aside className="lg:col-span-1 space-y-4">
          <section className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
            <h2 className="font-semibold text-slate-700 mb-3">Thông tin đề</h2>
            <div className="space-y-3">
              <Field label="Trường" value={meta.school}
                onChange={(v) => updateMeta("school", v)}
                placeholder="TRƯỜNG THPT CÂY DƯƠNG" />
              <Field label="Tiêu đề" value={meta.title}
                onChange={(v) => updateMeta("title", v)} />
              <Field label="Môn" value={meta.subject}
                onChange={(v) => updateMeta("subject", v)} />
              <div className="grid grid-cols-2 gap-3">
                <Field label="Thời gian" value={meta.duration}
                  onChange={(v) => updateMeta("duration", v)} />
                <Field label="Lớp" value={meta.class_name}
                  onChange={(v) => updateMeta("class_name", v)} />
              </div>
              <Field label="Mã đề" value={meta.made}
                onChange={(v) => updateMeta("made", v)} />
            </div>
          </section>

          <section className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
            <h2 className="font-semibold text-slate-700 mb-3">Tên file tải về</h2>
            <Field label="" value={filename}
              onChange={setFilename} placeholder="de_thi.docx" />
          </section>

          <section className="bg-amber-50 border border-amber-200 rounded-xl p-4 text-sm text-amber-900">
            <p className="font-semibold mb-1">Mẹo dùng</p>
            <ul className="list-disc pl-5 space-y-1">
              <li>Copy câu hỏi <code>\begin&#123;ex&#125;</code> từ ngân hàng đề của cô.</li>
              <li>Hình <b>TikZ</b> và <b>tabular</b> tự vẽ qua server.</li>
              <li>Công thức Toán <b>chỉnh sửa được</b> trong Word.</li>
            </ul>
          </section>
        </aside>

        {/* Cột phải: textarea + nút */}
        <section className="lg:col-span-2 space-y-4">
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
            <div className="flex items-center justify-between mb-3">
              <h2 className="font-semibold text-slate-700">
                Nội dung LaTeX
                {countEx > 0 && (
                  <span className="ml-2 text-xs bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded-full">
                    {countEx} khối ex
                  </span>
                )}
              </h2>
              <div className="flex gap-2">
                <button
                  onClick={() => setText(SAMPLE)}
                  className="text-sm text-slate-600 hover:text-indigo-600"
                  type="button"
                >
                  Chèn mẫu
                </button>
                <button
                  onClick={() => setText("")}
                  className="text-sm text-slate-600 hover:text-rose-600"
                  type="button"
                >
                  Xoá
                </button>
              </div>
            </div>
            <textarea
              value={text}
              onChange={(e) => setText(e.target.value)}
              spellCheck={false}
              className="w-full h-[420px] font-mono text-sm p-3 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-400 focus:border-indigo-400"
              placeholder="Dán các khối \begin{ex}...\end{ex} vào đây..."
            />
          </div>

          {err && (
            <div className="bg-rose-50 border border-rose-200 rounded-xl px-4 py-3 text-rose-800 text-sm">
              {err}
            </div>
          )}

          <div className="flex flex-wrap items-center gap-3">
            <button
              onClick={() => handleConvert("equation")}
              disabled={busy !== ""}
              className="inline-flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-400 text-white font-semibold px-6 py-3 rounded-xl shadow-sm transition"
              type="button"
              title="Công thức Toán chuyển sang Equation của Word — chỉnh sửa được trực tiếp trong Word."
            >
              {busy === "equation" ? (
                <>
                  <Spinner /> Đang chuyển…
                </>
              ) : (
                <>📝 Chuyển Word - Equation &amp; Tải về</>
              )}
            </button>
            <button
              onClick={() => handleConvert("latex")}
              disabled={busy !== ""}
              className="inline-flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-400 text-white font-semibold px-6 py-3 rounded-xl shadow-sm transition"
              type="button"
              title="Giữ nguyên $..$/$$..$$ trong Word — thuận tiện chuyển tiếp qua MathType trên máy."
            >
              {busy === "latex" ? (
                <>
                  <Spinner /> Đang chuyển…
                </>
              ) : (
                <>🧮 Chuyển Word - LaTeX &amp; Tải về</>
              )}
            </button>
            <span className="text-sm text-slate-500">
              Máy chủ: <code>{API_URL}</code>
            </span>
          </div>
        </section>
      </main>

      <footer className="max-w-6xl mx-auto px-6 py-8 text-center text-slate-400 text-sm">
        Miễn phí. Công thức &amp; hình vẽ render qua server ngoài (cần Internet).
      </footer>
    </div>
  );
}

function Field({
  label,
  value,
  onChange,
  placeholder,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  placeholder?: string;
}) {
  return (
    <label className="block">
      {label && (
        <span className="block text-xs font-medium text-slate-500 mb-1">{label}</span>
      )}
      <input
        type="text"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-400 focus:border-indigo-400"
      />
    </label>
  );
}

function Spinner() {
  return (
    <svg
      className="animate-spin h-4 w-4"
      xmlns="http://www.w3.org/2000/svg"
      fill="none"
      viewBox="0 0 24 24"
    >
      <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" className="opacity-25" />
      <path
        fill="currentColor"
        className="opacity-75"
        d="M4 12a8 8 0 018-8v4a4 4 0 00-4 4H4z"
      />
    </svg>
  );
}
