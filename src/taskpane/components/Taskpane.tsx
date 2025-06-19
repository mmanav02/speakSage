import * as React from "react";
import { useAnthropic } from "../hooks/callLLM";   // keep the hook file

export const Taskpane: React.FC = () => {
  const [apiKey, setApiKey] = React.useState("");
  const [prompt, setPrompt] = React.useState("");
  const [files,  setFiles]  = React.useState<File[]>([]);

  const { loading, output, call, encodeImages } = useAnthropic();

  const run = async () => {
    const images = await encodeImages(files);
    await call({ apiKey, prompt, images });
  };

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
              onChange={e => setFiles(Array.from(e.target.files ?? []))}
              hidden
            />
          </label>
          {files.length > 0 && (
            <span style={styles.uploadInfo}>
              {files.length} image{files.length > 1 ? "s" : ""}
            </span>
          )}
        </Field>

        {files.length > 0 && (
          <div style={styles.thumbStrip}>
            {files.map((f, i) => (
              <img
                key={i}
                src={URL.createObjectURL(f)}
                alt={f.name}
                style={styles.thumb}
              />
            ))}
          </div>
        )}

        <button
          style={styles.runBtn}
          onClick={run}
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

/* ── presentational helpers ──────────────────────────── */

const Header = () => (
  <div style={styles.header}>
    <span style={styles.logo}>🧠</span>
    <h1 style={styles.title}>SheetSage</h1>
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

/* ── inline style object (same values as the CSS) ────────────────────────── */

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
  runBtnHover: { background: "#218838" }, // optional hover handling
};
