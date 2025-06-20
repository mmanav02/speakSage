import * as React from "react";
import axios from "axios";

/* ── helper: strip ``` fences & check braces ───────────────────────── */
const extractExecutableCode = (
  raw: string,
): { ok: boolean; code: string; err?: string } => {
  const match   = raw.match(/```(?:typescript|javascript)?\s*([\s\S]*?)```/i);
  const snippet = (match ? match[1] : raw).trim();
  const safe    = snippet.replace(/`/g, "\\`");
  if ((safe.match(/[({]/g) || []).length !== (safe.match(/[)}]/g) || []).length) {
    return { ok: false, code: "", err: "Brace / paren mismatch – reply looks truncated." };
  }
  return { ok: true, code: safe };
};

/* ── helper: capture live sheet values + numberFormat ──────────────── */
const captureSheetJSON = async () => {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const rng   = sheet.getUsedRange();
    rng.load(["values", "numberFormat"]);
    await ctx.sync();
    return JSON.stringify({
      values:       rng.values,
      numberFormat: rng.numberFormat,
    });
  });
};

export const Taskpane: React.FC = () => {
  const [apiKey,  setApiKey]  = React.useState("");
  const [prompt,  setPrompt]  = React.useState("");
  const [images,  setImages]  = React.useState<File[]>([]);
  const [output,  setOutput]  = React.useState("");
  const [loading, setLoading] = React.useState(false);

  /* helper – base-64 encode images */
  const encodeImages = (files: File[]) =>
    Promise.all(files.map(f => new Promise<string>((res, rej) => {
      const r = new FileReader();
      r.onerror = () => rej(r.error);
      r.onload  = () => res(r.result as string);
      r.readAsDataURL(f);
    })));

  /* ── main handler ────────────────────────────────────────────────── */
  const callAnthropic = async () => {
    setLoading(true);
    setOutput("");

    try {
      /* 1️⃣  snapshot sheet BEFORE we talk to Claude */
      const sheetJSON = await captureSheetJSON();

      /* 2️⃣ build system prompt that embeds that snapshot */
      const systemPrompt = `You are an assistant that writes Office.js code for Excel.

        Below is the CURRENT sheet state in JSON (two top-level keys: "values" and "numberFormat").
        <<<SHEET_STATE>>>
        ${sheetJSON}
        <<<END>>>

        • Preserve existing data unless explicitly told to overwrite.
        • Match or extend current formatting (consult "numberFormat").
        • Return ONLY executable JavaScript (no Markdown).

        :: STRICT RULES ::
        1. Compute rows/cols for every 2-D array, ensure rectangular.
        2. Use getRangeByIndexes before assigning .values / .formulas.
        3. Wrap all work in: await Excel.run(async (context) => { … });
        4. End with 'await context.sync();' and nothing after it.`;

      /* 3️⃣  encode images (if any) */
      const imageData = await encodeImages(images);

      /* 4️⃣  call local proxy → Claude */
      const { data } = await axios.post("https://127.0.0.1:5050/anthropic", {
        apiKey,
        prompt,
        systemPrompt,
        images: imageData,
        sheet: sheetJSON,          // forwarded, useful for logging/debug
      });

      /* 5️⃣  show Claude's raw reply & extract executable code */
      const raw = data?.content?.[0]?.text ?? "";
      setOutput(raw);

      const { ok, code, err } = extractExecutableCode(raw);
      if (!ok) {
        setOutput("❌ " + err);
        return;
      }

      /* 6️⃣  run code inside Excel */
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

  /* ── JSX UI (unchanged) ─────────────────────────────────────────── */
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
            placeholder='e.g. "Align new table with existing styles"'
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

/* ── tiny sub-components & inline styles (unchanged) ──────────────── */

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
    width: "100%", padding: 10, fontSize: 14, border: "1px solid #ccc",
    borderRadius: 6, boxSizing: "border-box", resize: "vertical",
  },
  textareaPrompt: {
    width: "100%", minHeight: 120, padding: 10, fontSize: 14,
    border: "1px solid #ccc", borderRadius: 6, boxSizing: "border-box",
    resize: "vertical",
  },
  textarea: {
    width: "100%", padding: 10, fontSize: 14, background: "#f9f9f9",
    border: "1px solid #ccc", borderRadius: 6, boxSizing: "border-box",
    resize: "vertical",
  },
  uploadBtn: {
    padding: "10px 14px", background: "#0078d4", color: "#fff",
    fontWeight: 600, borderRadius: 6, cursor: "pointer", fontSize: 12,
  },
  uploadInfo: { marginLeft: 8, fontSize: 12 },
  thumbStrip: { display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" },
  thumb: {
    width: 48, height: 48, objectFit: "cover", borderRadius: 4,
    border: "1px solid #ccc",
  },
  runBtn: {
    width: "100%", padding: 10, marginBottom: 16, background: "#28a745",
    color: "#fff", fontWeight: 600, border: "none", borderRadius: 6,
    cursor: "pointer",
  },
};
