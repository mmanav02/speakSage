import * as React from "react";
import axios from "axios";


export const Taskpane: React.FC = () => {
  const [apiKey, setApiKey] = React.useState("");
  const [prompt, setPrompt] = React.useState("");
  const [output, setOutput] = React.useState("");
  const [loading, setLoading] = React.useState(false);

  const callAnthropic = async () => {
    setLoading(true);
    setOutput("");

    try {
      const systemPrompt = `You are an assistant that edits Excel spreadsheets using the Office.js API. Always return only the executable JavaScript code wrapped in:
await Excel.run(async (context) => { ... });`;
      console.log(apiKey);
      // call your local proxy
      const response = await axios.post("https://127.0.0.1:5050/anthropic", {
        apiKey,
        prompt,
        systemPrompt,
      });

      const reply = response.data?.content?.[0]?.text ?? "";
      setOutput(reply);                      // show raw Claude text

      // ② Strip ``` fences if any
      const clean = reply.replace(/```[a-z]*|```/g, "").trim();

      // ③ Wrap inside an async IIFE so leading “await” is valid JS
      const wrapped = new Function(
        "Excel",
        `"use strict";
         return (async () => {
           ${clean}
         })();`
      );

      // ④ Execute
      await wrapped(Excel as any);
    } catch (error: any) {
      console.error("Execution error:", error);
      setOutput(`❌ Error: ${error?.response?.data?.error || error.message}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <div style={styles.header}>
          <span style={styles.logo}>🧠</span>
          <h1 style={styles.title}>speakExcel</h1>
        </div>

        <div style={styles.section}>
          <label style={styles.label}>Anthropic API Key</label>
          <input
            type="password"
            value={apiKey}
            onChange={(e) => setApiKey(e.target.value)}
            placeholder="sk-ant-..."
            style={styles.input}
          />
        </div>

        <div style={styles.section}>
          <label style={styles.label}>Prompt</label>
          <input
            type="text"
            value={prompt}
            onChange={(e) => setPrompt(e.target.value)}
            placeholder='e.g. "Bold the first row and center it"'
            style={styles.input}
          />
        </div>

        <button
          onClick={callAnthropic}
          style={styles.button}
          disabled={loading || !apiKey.trim() || !prompt.trim()}
        >
          {loading ? "Running..." : "Run"}
        </button>

        <div style={styles.section}>
          <label style={styles.label}>Claude’s Response (editable JS)</label>
          <textarea
            value={output}
            readOnly
            rows={6}
            style={styles.textarea}
            placeholder="Claude's generated Excel.run() script will appear here..."
          />
        </div>
      </div>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  container: {
    padding: "1rem",
    backgroundColor: "#f3f2f1",
    height: "100%",
    boxSizing: "border-box",
  },
  card: {
    backgroundColor: "#ffffff",
    borderRadius: "10px",
    boxShadow: "0 0 8px rgba(0,0,0,0.1)",
    padding: "20px",
    fontFamily: "Segoe UI, sans-serif",
  },
  header: {
    display: "flex",
    alignItems: "center",
    marginBottom: "1rem",
  },
  logo: {
    fontSize: "2rem",
    marginRight: "10px",
  },
  title: {
    fontSize: "1.5rem",
    margin: 0,
    color: "#0078d4",
  },
  section: {
    marginBottom: "1rem",
  },
  label: {
    fontWeight: 600,
    marginBottom: "0.2rem",
    display: "block",
    color: "#323130",
  },
  input: {
    width: "100%",
    padding: "10px",
    fontSize: "14px",
    border: "1px solid #ccc",
    borderRadius: "6px",
    boxSizing: "border-box",
  },
  textarea: {
    width: "100%",
    padding: "10px",
    fontSize: "14px",
    backgroundColor: "#f9f9f9",
    border: "1px solid #ccc",
    borderRadius: "6px",
    resize: "vertical",
  },
  button: {
    width: "100%",
    padding: "10px",
    backgroundColor: "#0078d4",
    color: "white",
    fontSize: "14px",
    fontWeight: 600,
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    marginBottom: "1rem",
  },
};
