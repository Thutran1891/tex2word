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

const TIKZ_SAMPLE = String.raw`\begin{tikzpicture}
\draw[->] (-2.2,0)--(2.2,0) node[right]{$x$};
\draw[->] (0,-2.2)--(0,2.2) node[above]{$y$};
\draw[blue,thick] (0,0) circle (1.5);
\node at (0,-0.3) {$O$};
\end{tikzpicture}`;

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

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

// ------- Tiện ích dùng chung -------
function triggerDownload(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function base64ToBlob(b64: string, mime: string): Blob {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type: mime });
}

type TabKey = "tex2doc" | "tikz2png";

export default function App() {
  const [tab, setTab] = useState<TabKey>("tex2doc");

  return (
    <div className="min-h-screen">
      <header className="bg-gradient-to-r from-indigo-600 to-violet-600 text-white">
        <div className="max-w-6xl mx-auto px-6 py-5">
          <div className="flex flex-wrap items-start justify-between gap-3">
            <div>
              <h1 className="text-2xl font-bold tracking-tight">
                Công cụ convert LaTeX to Word &amp; TikZ to PNG
              </h1>
              <p className="text-indigo-100 text-sm mt-1">
                Chuyển câu hỏi LaTeX sang Word &amp; vẽ hình TikZ thành ảnh.
              </p>
            </div>
            <div className="text-right text-indigo-100 text-xs leading-relaxed shrink-0">
              <div className="font-semibold text-white">Trần Thị Kim Thu</div>
              <div>TRƯỜNG THPT CÂY DƯƠNG</div>
              <div>0397 58 43 58</div>
            </div>
          </div>

          {/* Thanh tab */}
          <div className="mt-4 flex gap-1">
            <TabButton
              active={tab === "tex2doc"}
              onClick={() => setTab("tex2doc")}
              label="📝 Tex2Doc"
            />
            <TabButton
              active={tab === "tikz2png"}
              onClick={() => setTab("tikz2png")}
              label="🖼️ Tikz2PNG"
            />
          </div>
        </div>
      </header>

      {tab === "tex2doc" ? <Tex2DocTab /> : <Tikz2PngTab />}

      <footer className="max-w-6xl mx-auto px-6 py-8 text-center text-slate-400 text-sm">
        Miễn phí. Công thức &amp; hình vẽ render qua server ngoài (cần Internet).
      </footer>
    </div>
  );
}

function TabButton({
  active,
  onClick,
  label,
}: {
  active: boolean;
  onClick: () => void;
  label: string;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={
        "px-5 py-2 rounded-t-lg text-sm font-semibold transition " +
        (active
          ? "bg-slate-50 text-indigo-700 shadow-sm"
          : "bg-white/10 text-white hover:bg-white/20")
      }
    >
      {label}
    </button>
  );
}

// ================= TAB 1: Tex2Doc =================
function Tex2DocTab() {
  const [text, setText] = useState<string>(SAMPLE);
  const [meta, setMeta] = useState<Meta>(emptyMeta);
  const [filename, setFilename] = useState<string>("de_thi.docx");
  const [busy, setBusy] = useState<"" | "equation" | "latex">("");
  const [err, setErr] = useState<string>("");
  const [tikz, setTikz] = useState<{ done: number; total: number } | null>(null);
  // Có in tiêu đề (header) + các đề mục "Phần I/II/III/IV" hay không.
  const [showHeader, setShowHeader] = useState<boolean>(true);

  function updateMeta<K extends keyof Meta>(key: K, value: Meta[K]) {
    setMeta((m) => ({ ...m, [key]: value }));
  }

  function countImages(src: string): number {
    const t = (src.match(/\\begin\s*\{tikzpicture\}/g) || []).length;
    const tab = (src.match(/\\begin\s*\{tabular\}/g) || []).length;
    return t + tab;
  }

  async function streamConvert(
    mode: "equation" | "latex",
    outName: string,
    imgCount: number
  ) {
    setTikz({ done: 0, total: imgCount });
    const resp = await fetch(`${API_URL}/convert-stream`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text, meta, filename: outName, mode, show_header: showHeader }),
    });
    if (!resp.ok || !resp.body) {
      let msg = `Lỗi ${resp.status}`;
      try {
        const j = await resp.json();
        if (j?.detail) msg = j.detail;
      } catch {
        /* keep default */
      }
      throw new Error(msg);
    }
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buf = "";
    let downloaded = false;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      let nl: number;
      while ((nl = buf.indexOf("\n")) >= 0) {
        const line = buf.slice(0, nl).trim();
        buf = buf.slice(nl + 1);
        if (!line) continue;
        const msg = JSON.parse(line) as {
          type: string;
          done?: number;
          total?: number;
          data?: string;
          detail?: string;
        };
        if (msg.type === "progress") {
          setTikz({ done: msg.done ?? 0, total: msg.total ?? imgCount });
        } else if (msg.type === "done" && msg.data) {
          setTikz((p) => (p ? { done: p.total, total: p.total } : p));
          triggerDownload(base64ToBlob(msg.data, DOCX_MIME), outName);
          downloaded = true;
        } else if (msg.type === "error") {
          throw new Error(msg.detail || "Lỗi khi xuất Word.");
        }
      }
    }
    if (!downloaded) throw new Error("Máy chủ không trả về file. Cô thử lại nhé.");
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
    const outName = (() => {
      const base = filename.endsWith(".docx") ? filename.slice(0, -5) : filename;
      const suffix = mode === "latex" ? "-latex" : "";
      return `${base}${suffix}.docx`;
    })();
    const imgCount = countImages(text);
    setBusy(mode);
    try {
      if (imgCount === 0) {
        const resp = await fetch(`${API_URL}/convert`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text, meta, filename: outName, mode, show_header: showHeader }),
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
        triggerDownload(await resp.blob(), outName);
      } else {
        await streamConvert(mode, outName, imgCount);
      }
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy("");
      setTikz(null);
    }
  }

  const countEx = (text.match(/\\begin\{ex\}/g) || []).length;

  return (
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

        <section className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
          <h2 className="font-semibold text-slate-700 mb-3">Tiêu đề &amp; đề mục</h2>
          <div className="space-y-2">
            <label className="flex items-center gap-2 text-sm text-slate-700 cursor-pointer">
              <input
                type="radio"
                name="show-header"
                checked={showHeader}
                onChange={() => setShowHeader(true)}
                className="h-4 w-4"
              />
              Có hiện tiêu đề và các đề mục
            </label>
            <label className="flex items-center gap-2 text-sm text-slate-700 cursor-pointer">
              <input
                type="radio"
                name="show-header"
                checked={!showHeader}
                onChange={() => setShowHeader(false)}
                className="h-4 w-4"
              />
              Không hiện tiêu đề và các đề mục
            </label>
          </div>
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
              <button onClick={() => setText(SAMPLE)}
                className="text-sm text-slate-600 hover:text-indigo-600" type="button">
                Chèn mẫu
              </button>
              <button onClick={() => setText("")}
                className="text-sm text-slate-600 hover:text-rose-600" type="button">
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
            {busy === "equation" ? (<><Spinner /> Đang chuyển…</>) : (<>📝 Chuyển Word - Equation &amp; Tải về</>)}
          </button>
          <button
            onClick={() => handleConvert("latex")}
            disabled={busy !== ""}
            className="inline-flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-400 text-white font-semibold px-6 py-3 rounded-xl shadow-sm transition"
            type="button"
            title="Giữ nguyên $..$/$$..$$ trong Word — thuận tiện chuyển tiếp qua MathType trên máy."
          >
            {busy === "latex" ? (<><Spinner /> Đang chuyển…</>) : (<>🧮 Chuyển Word - LaTeX &amp; Tải về</>)}
          </button>
          <span className="text-sm text-slate-500">
            Máy chủ: <code>{API_URL}</code>
          </span>
        </div>
      </section>

      {tikz && <TikzModal done={tikz.done} total={tikz.total} />}
    </main>
  );
}

// ================= TAB 2: Tikz2PNG =================
function Tikz2PngTab() {
  const [code, setCode] = useState<string>(TIKZ_SAMPLE);
  const [transparent, setTransparent] = useState<boolean>(true);
  const [busy, setBusy] = useState<boolean>(false);
  const [err, setErr] = useState<string>("");
  const [pngB64, setPngB64] = useState<string>("");
  const [copied, setCopied] = useState<boolean>(false);

  async function handleRender() {
    setErr("");
    setCopied(false);
    if (!code.trim()) {
      setErr("Cô ơi, chưa nhập mã TikZ.");
      return;
    }
    setBusy(true);
    setPngB64("");
    try {
      const resp = await fetch(`${API_URL}/tikz-png`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ source: code, transparent }),
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
      const j = (await resp.json()) as { image_base64: string };
      setPngB64(j.image_base64);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  async function handleCopy() {
    if (!pngB64) return;
    try {
      const blob = base64ToBlob(pngB64, "image/png");
      await navigator.clipboard.write([
        new ClipboardItem({ "image/png": blob }),
      ]);
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    } catch (e) {
      setErr(
        "Trình duyệt không cho copy ảnh trực tiếp. Cô dùng nút 'Tải PNG về máy' rồi chèn vào Word nhé. (" +
          (e instanceof Error ? e.message : String(e)) +
          ")"
      );
    }
  }

  function handleDownload() {
    if (!pngB64) return;
    triggerDownload(base64ToBlob(pngB64, "image/png"), "hinh_tikz.png");
  }

  return (
    <main className="max-w-6xl mx-auto px-6 py-6 grid grid-cols-1 lg:grid-cols-2 gap-6">
      {/* Cột trái: nhập mã */}
      <section className="space-y-4">
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
          <div className="flex items-center justify-between mb-3">
            <h2 className="font-semibold text-slate-700">Mã TikZ</h2>
            <div className="flex gap-2">
              <button onClick={() => setCode(TIKZ_SAMPLE)}
                className="text-sm text-slate-600 hover:text-indigo-600" type="button">
                Chèn mẫu
              </button>
              <button onClick={() => setCode("")}
                className="text-sm text-slate-600 hover:text-rose-600" type="button">
                Xoá
              </button>
            </div>
          </div>
          <textarea
            value={code}
            onChange={(e) => setCode(e.target.value)}
            spellCheck={false}
            className="w-full h-[360px] font-mono text-sm p-3 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-400 focus:border-indigo-400"
            placeholder="Dán mã \begin{tikzpicture}...\end{tikzpicture} vào đây..."
          />
        </div>

        <label className="flex items-center gap-2 text-sm text-slate-600 select-none">
          <input
            type="checkbox"
            checked={transparent}
            onChange={(e) => setTransparent(e.target.checked)}
            className="h-4 w-4"
          />
          Nền trong suốt (bỏ tick nếu muốn nền trắng)
        </label>

        {err && (
          <div className="bg-rose-50 border border-rose-200 rounded-xl px-4 py-3 text-rose-800 text-sm whitespace-pre-wrap">
            {err}
          </div>
        )}

        <button
          onClick={handleRender}
          disabled={busy}
          className="inline-flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-400 text-white font-semibold px-6 py-3 rounded-xl shadow-sm transition"
          type="button"
        >
          {busy ? (<><Spinner /> Đang vẽ…</>) : (<>🖼️ Convert TikZ To PNG</>)}
        </button>
      </section>

      {/* Cột phải: kết quả */}
      <section className="space-y-4">
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 min-h-[360px] flex flex-col">
          <h2 className="font-semibold text-slate-700 mb-3">Kết quả PNG</h2>
          <div
            className="flex-1 flex items-center justify-center rounded-lg border border-dashed border-slate-300 bg-[repeating-conic-gradient(#f1f5f9_0_25%,#ffffff_0_50%)] bg-[length:20px_20px] p-3"
          >
            {pngB64 ? (
              <img
                src={`data:image/png;base64,${pngB64}`}
                alt="Hình TikZ"
                className="max-h-[320px] max-w-full object-contain"
              />
            ) : (
              <span className="text-slate-400 text-sm">
                {busy ? "Đang biên dịch TikZ…" : "Ảnh sẽ hiện ở đây sau khi bấm Convert."}
              </span>
            )}
          </div>
        </div>

        {pngB64 && (
          <div className="flex flex-wrap items-center gap-3">
            <button
              onClick={handleCopy}
              className="inline-flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-5 py-3 rounded-xl shadow-sm transition"
              type="button"
              title="Copy ảnh vào bộ nhớ tạm để dán thẳng (Ctrl+V) vào Word."
            >
              {copied ? "✅ Đã copy! Dán vào Word (Ctrl+V)" : "📋 Copy PNG (dán vào Word)"}
            </button>
            <button
              onClick={handleDownload}
              className="inline-flex items-center gap-2 bg-slate-600 hover:bg-slate-700 text-white font-semibold px-5 py-3 rounded-xl shadow-sm transition"
              type="button"
            >
              ⬇️ Tải PNG về máy
            </button>
          </div>
        )}

        <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 text-sm text-amber-900">
          <p className="font-semibold mb-1">Mẹo dùng</p>
          <ul className="list-disc pl-5 space-y-1">
            <li>Dán mã <code>\begin&#123;tikzpicture&#125;</code> → bấm <b>Convert</b>.</li>
            <li>Bấm <b>Copy PNG</b> rồi vào Word nhấn <b>Ctrl+V</b> là có ảnh.</li>
            <li>Ảnh nền trong suốt hợp để chèn lên nền màu; cần nền trắng thì bỏ tick.</li>
          </ul>
        </div>
      </section>
    </main>
  );
}

// ================= Thành phần dùng chung =================
function TikzModal({ done, total }: { done: number; total: number }) {
  const percent = total > 0 ? Math.round((done / total) * 100) : 0;
  const R = 52;
  const C = 2 * Math.PI * R;
  const offset = C - (percent / 100) * C;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl px-10 py-8 w-[min(90vw,420px)] text-center">
        <div className="relative mx-auto h-32 w-32">
          <svg className="h-32 w-32 -rotate-90" viewBox="0 0 120 120">
            <circle cx="60" cy="60" r={R} fill="none" stroke="#e2e8f0" strokeWidth="10" />
            <circle
              cx="60" cy="60" r={R} fill="none"
              stroke="#6366f1" strokeWidth="10" strokeLinecap="round"
              strokeDasharray={C} strokeDashoffset={offset}
              style={{ transition: "stroke-dashoffset 0.4s ease" }}
            />
          </svg>
          <div className="absolute inset-0 flex items-center justify-center">
            <span className="text-2xl font-bold text-indigo-600">{percent}%</span>
          </div>
        </div>
        <h3 className="mt-5 text-xl font-bold tracking-wide text-slate-800">
          ĐANG XỬ LÝ TIKZ
        </h3>
        <p className="mt-1 text-indigo-600 font-medium">
          Đã xong: {done} / {total} hình vẽ
        </p>
        <div className="mt-4 h-2.5 w-full rounded-full bg-slate-100 overflow-hidden">
          <div className="h-full rounded-full bg-indigo-500"
            style={{ width: `${percent}%`, transition: "width 0.4s ease" }} />
        </div>
        <p className="mt-4 text-xs uppercase tracking-wider text-slate-400 leading-relaxed">
          Hệ thống đang biên dịch tuần tự để đảm bảo độ ổn định
        </p>
      </div>
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
    <svg className="animate-spin h-4 w-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" className="opacity-25" />
      <path fill="currentColor" className="opacity-75" d="M4 12a8 8 0 018-8v4a4 4 0 00-4 4H4z" />
    </svg>
  );
}
