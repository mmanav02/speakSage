import * as React from "react";
import axios from "axios";

/* ── helper: strip ``` fences & check braces ───────────────── */
const extractExecutableCode = (
  raw: string,
): { ok: boolean; code: string; err?: string } => {
  const match   = raw.match(/```(?:typescript|javascript)?\s*([\s\S]*?)```/i);
  const snippet = (match ? match[1] : raw).trim();
  const safe    = snippet.replace(/`/g, "\\`");

  const opens  = (safe.match(/[({]/g) || []).length;
  const closes = (safe.match(/[)}]/g) || []).length;
  if (opens !== closes) {
    return { ok: false, code: "", err: "Brace / paren mismatch – reply looks truncated." };
  }
  return { ok: true, code: safe };
};

export const Taskpane: React.FC = () => {
  const [apiKey,  setApiKey]  = React.useState("");
  const [prompt,  setPrompt]  = React.useState("");
  const [images,  setImages]  = React.useState<File[]>([]);
  const [output,  setOutput]  = React.useState("");
  const [loading, setLoading] = React.useState(false);

  /* helper – base-64 encode images */
  const encodeImages = (files: File[]) =>
    Promise.all(
      files.map(
        f =>
          new Promise<string>((res, rej) => {
            const r = new FileReader();
            r.onerror = () => rej(r.error);
            r.onload  = () => res(r.result as string);
            r.readAsDataURL(f);
          }),
      ),
    );

  /* main handler */
  const callAnthropic = async () => {
    setLoading(true);
    setOutput("");

    try {
      const systemPrompt = `You are an assistant that writes Office.js code for Excel.
Return ONLY executable JavaScript (no Markdown, no comments).

:: STRICT RULES ::
1.  When writing a 2-D array 'data', compute:
      const rows = data.length;
      const cols = data[0].length;
    and assert every row.length === cols (throw if not).

2.  Before assigning .values/.formulas:
      const rng = sheet.getRangeByIndexes(r0, c0, rows, cols);
      rng.values = data;
    Never assign if array dims ≠ range dims.

3.  ALWAYS wrap work in:
      await Excel.run(async (context) => { … });
4.  End with 'await context.sync();' and nothing after it.`;

      const imageData = await encodeImages(images);

      /* call local proxy → Anthropic */
      const { data } = await axios.post("https://127.0.0.1:5050/anthropic", {
        apiKey,
        prompt,
        systemPrompt,
        images: imageData,
      });

      const raw = data?.content?.[0]?.text ?? "";
      setOutput(raw);

      const { ok, code, err } = extractExecutableCode(raw);
      if (!ok) {
        setOutput("❌ " + err);
        return;
      }

      /* run code in Excel */
      let fn: Function;
      try {
        fn = new Function(
          `"use strict";\nreturn (async (Excel) => {\n${code}\n})(arguments[0]);`,
        );
      } catch (syntaxErr: any) {
        setOutput("❌ Syntax error: " + syntaxErr.message);
        return;
      }

      try {
        await fn(Excel as any);
      } catch (runtimeErr: any) {
        setOutput(
          "❌ Runtime error: " + runtimeErr.message +
          "\n\nLast script:\n" + code +
          "\n\nExcel stack:\n" + (runtimeErr.stack || ""),
        );
      }
    } catch (err: any) {
      setOutput("❌ " + (err.response?.data?.error || err.message));
    } finally {
      setLoading(false);
    }
  };

  /* JSX */
  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <Header />

        <Field
          label="Anthropic API Key"
          value={apiKey}
          onChange={setApiKey}
          type="password"
          placeholder="sk-ant-…"
        />

        <Field label="Prompt">
          <textarea
            style={styles.textareaPrompt}
            placeholder='e.g. "Format like the images"'
            value={prompt}
            onChange={e => setPrompt(e.target.value)}
          />
        </Field>

        <Field label="Images">
          <label style={styles.uploadBtn}>
            Upload 📎
            <input
              type="file"
              accept="image/*"
              multiple
              onChange={e => setImages(Array.from(e.target.files ?? []))}
              hidden
            />
          </label>
          {images.length > 0 && (
            <span style={styles.uploadInfo}>
              {images.length} image{images.length > 1 ? "s" : ""}
            </span>
          )}
        </Field>

        {images.length > 0 && (
          <div style={styles.thumbStrip}>
            {images.map((f, i) => (
              <img key={i} src={URL.createObjectURL(f)} alt={f.name} style={styles.thumb} />
            ))}
          </div>
        )}

        <button
          style={styles.runBtn}
          onClick={callAnthropic}
          disabled={loading || !apiKey.trim() || !prompt.trim()}
        >
          {loading ? "Running…" : "Run"}
        </button>

        <Field label="Claude response (JS)">
          <textarea readOnly rows={6} value={output} style={styles.textarea} />
        </Field>
      </div>
    </div>
  );
};

/* ── tiny sub-components ─────────────────────────── */

const Header = () => (
  <div style={styles.header}>
    <span style={styles.logo}>🧠</span>
    <h1 style={styles.title}>speakSage</h1>
  </div>
);

const Field: React.FC<{
  label: string;
  children?: React.ReactNode;
  value?: string;
  onChange?: (s: string) => void;
  type?: string;
  placeholder?: string;
}> = ({ label, children, value, onChange, type = "text", placeholder }) => (
  <div style={styles.section}>
    <label style={styles.label}>{label}</label>
    {children ?? (
      <input
        type={type}
        value={value}
        onChange={e => onChange?.(e.target.value)}
        placeholder={placeholder}
        style={styles.input}
      />
    )}
  </div>
);

/* ── inline styles object ─────────────────────────── */

const styles: { [k: string]: React.CSSProperties } = {
  container: { padding: "1rem", background: "#f3f2f1", height: "100%" },
  card: {
    background: "#fff",
    borderRadius: 10,
    padding: 20,
    boxShadow: "0 0 8px rgba(0,0,0,.1)",
    fontFamily: "Segoe UI, sans-serif",
  },
  header: { display: "flex", alignItems: "center", marginBottom: 16 },
  logo: { fontSize: 24, marginRight: 8 },
  title: { margin: 0, fontSize: "1.5rem", color: "#0078d4" },

  section: { marginBottom: 16 },
  label: { fontWeight: 600, marginBottom: 4, display: "block" },

  input: {
    width: "100%",
    padding: 10,
    fontSize: 14,
    border: "1px solid #ccc",
    borderRadius: 6,
    boxSizing: "border-box",
    resize: "vertical",
  },
  textareaPrompt: {
    width: "100%",
    minHeight: 120,
    padding: 10,
    fontSize: 14,
    border: "1px solid #ccc",
    borderRadius: 6,
    boxSizing: "border-box",
    resize: "vertical",
  },
  textarea: {
    width: "100%",
    padding: 10,
    fontSize: 14,
    background: "#f9f9f9",
    border: "1px solid #ccc",
    borderRadius: 6,
    boxSizing: "border-box",
    resize: "vertical",
  },

  uploadBtn: {
    padding: "10px 14px",
    background: "#0078d4",
    color: "#fff",
    fontWeight: 600,
    borderRadius: 6,
    cursor: "pointer",
    fontSize: 12,
    whiteSpace: "nowrap",
  },
  uploadInfo: { marginLeft: 8, fontSize: 12 },

  thumbStrip: { display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" },
  thumb: {
    width: 48,
    height: 48,
    objectFit: "cover",
    borderRadius: 4,
    border: "1px solid #ccc",
  },

  runBtn: {
    width: "100%",
    padding: 10,
    marginBottom: 16,
    background: "#28a745",
    color: "#fff",
    fontWeight: 600,
    border: "none",
    borderRadius: 6,
    cursor: "pointer",
  },
};
