// proxy-server.js  (run:  node proxy-server.js)
const express   = require("express");
const cors      = require("cors");
const axios     = require("axios");
const https     = require("https");
const devCerts  = require("office-addin-dev-certs");

/* ── helper: scrub Claude reply → safe JS ── */
function extractExecutableCode(raw) {
  const match   = raw.match(/```(?:typescript|javascript)?\s*([\s\S]*?)```/i);
  const snippet = (match ? match[1] : raw).trim();
  const safe    = snippet.replace(/`/g, "\\`");

  const opens  = (safe.match(/[({]/g) || []).length;
  const closes = (safe.match(/[)}]/g) || []).length;
  if (opens !== closes) {
    return { ok: false, code: "", err: "Brace / paren mismatch – reply seems truncated." };
  }
  return { ok: true, code: safe };
}

(async () => {
  const { key, cert } = await devCerts.getHttpsServerOptions();

  const app = express()
    .use(cors({ origin: "*" }))
    .use(express.json({ limit: "25mb" }));

  app.post("/anthropic", async (req, res) => {
    const { apiKey, prompt, systemPrompt = "", images = [] } = req.body;
    if (!apiKey || !prompt) {
      return res.status(400).json({ error: "apiKey and prompt are required" });
    }

    const imageBlocks = (Array.isArray(images) ? images : []).map((url) => {
      const match = url.match(/^data:(.+?);base64,(.*)$/);
      if (!match) throw new Error("Malformed image data URL.");
      const [, media_type, data] = match;
      return { type: "image", source: { type: "base64", media_type, data } };
    });

    const body = {
      model: "claude-opus-4-20250514",
      system: systemPrompt,
      max_tokens: 8196,
      messages: [
        { role: "user", content: [...imageBlocks, { type: "text", text: prompt }] },
      ],
    };

      console.log("──────── Claude Prompt Sent────────");

    try {
      const r = await axios.post(
        "https://api.anthropic.com/v1/messages",
        body,
        {
          headers: {
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
          },
          timeout: 60_000,
        },
      );

      console.log("──────── Claude Response Received────────");

      /* show what will be run */
      const raw = r.data?.content?.[0]?.text ?? "";
      const { ok, code, err } = extractExecutableCode(raw);
      if (!ok) {
        console.warn("⚠️  " + err);
      } else {
        console.log(
          "──────── Sanitised code sent to Excel ────────\n" +
            "\n──────────────────────────────────────────────",
        );
      }

      /* attach for client if wanted */
      r.data.sanitisedCode = ok ? code : null;
      r.data.sanitiseError = ok ? null : err;
      res.json(r.data);
    } catch (err) {
      const status  = err.response?.status || 500;
      const details = err.response?.data   || err.message;
      console.error("❌ Anthropic error:", details);
      res.status(status).json({ error: details });
    }
  });

  const PORT = 5050;
  https.createServer({ key, cert }, app).listen(PORT, () =>
    console.log(`🔐 Proxy live at https://127.0.0.1:${PORT}/anthropic`),
  );
})();
