import axios from "axios";
import { useState } from "react";

export interface AnthropicArgs {
  apiKey: string;
  prompt: string;
  images: string[];          // base-64 data URLs
}

export const useAnthropic = () => {
  const [loading, setLoading] = useState(false);
  const [output,  setOutput]  = useState("");

  const encodeImages = (files: File[]) =>
    Promise.all(files.map(f => new Promise<string>((res, rej) => {
      const r = new FileReader();
      r.onerror = () => rej(r.error);
      r.onload  = () => res(r.result as string);
      r.readAsDataURL(f);
    })));

  const call = async ({ apiKey, prompt, images }: AnthropicArgs) => {
    setLoading(true);
    setOutput("");
    try {
      const { data } = await axios.post("https://127.0.0.1:5050/anthropic", {
        apiKey,
        prompt,
        images,
        systemPrompt: SYSTEM_PROMPT,
      });
      setOutput(data?.content?.[0]?.text ?? "");
    } catch (err: any) {
      setOutput("❌ " + (err.response?.data?.error || err.message));
    } finally {
      setLoading(false);
    }
  };

  return { loading, output, call, encodeImages };
};

/* --- static --- */
const SYSTEM_PROMPT = `You are an assistant that writes Office.js code for Excel.
Return ONLY executable JavaScript (no Markdown, no comments).

:: STRICT RULES ::
1.  When writing a 2-D array 'data', compute:
      const rows = data.length;
      const cols = data[0].length;
    and assert that every row.length === cols. Throw if not.

2.  Before assigning .values or .formulas:
      const rng = sheet.getRangeByIndexes(r0, c0, rows, cols);
      rng.values = data;
    Never assign if array dims ≠ range dims.

3.  ALWAYS wrap work in:
      await Excel.run(async (context) => { … });
4.  End with 'await context.sync();' and nothing after it.`;

