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

    const systemPrompt = `
You are an assistant that helps edit Excel spreadsheets using the Office.js API.
Always respond with a complete and correct JavaScript snippet that uses:
  await Excel.run(async (context) => { ... });
Do NOT explain anything. Do NOT include Markdown or backticks.
Only return executable JavaScript code.`;

    try {
      const response = await axios.post(
        "https://api.anthropic.com/v1/messages",
        {
          model: "claude-3-opus-20240229",
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: prompt }
          ],
          max_tokens: 400
        },
        {
          headers: {
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
          }
        }
      );

      const responseText = response.data?.content?.[0]?.text ?? "";
      setOutput(responseText);

      const cleanCode = responseText.replace(/```[a-z]*|```/g, "").trim();

      const dynamicFn = new Function("Excel", `"use strict"; return (${cleanCode})`);
      await Excel.run(dynamicFn as (context: Excel.RequestContext) => Promise<unknown>);
    } catch (error: any) {
      console.error("Execution error:", error);
      setOutput(`❌ Error: ${error.message}`);
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
          disabled={loading || !apiKey.trim()}
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
