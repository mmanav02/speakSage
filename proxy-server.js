// proxy-server.js
//
// Secure local proxy for Anthropic ⇄ speakExcel
//
// ①  npm install express axios cors office-addin-dev-certs
// ②  node proxy-server.js
// ③  Front-end POSTs to  https://localhost:5050/anthropic
// -------------------------------------------------------

const express = require("express");
const axios   = require("axios");
const cors    = require("cors");
const fs      = require("fs");
const https   = require("https");
const devCerts = require("office-addin-dev-certs"); // generates trusted localhost certs

(async () => {
  // ────────────────────────────── 1.  Generate / read HTTPS cert/key
  const httpsOptions = await devCerts.getHttpsServerOptions();
  const { key, cert } = httpsOptions;   // already trusted if you ran `office-addin-dev-certs install`

  // ────────────────────────────── 2.  Express setup
  const app = express();
  app.use(cors({ origin: "*" }));       // relax CORS for local dev
  app.use(express.json());

  // ────────────────────────────── 3.  Proxy endpoint
  app.post("/anthropic", async (req, res) => {
    const { apiKey, prompt, systemPrompt = "" } = req.body;
    console.log("🔍 Incoming:", { prompt });

    if (!apiKey || !prompt) {
      return res.status(400).json({ error: "apiKey and prompt are required" });
    }

    try {
      const anthroRes = await axios.post(
        "https://api.anthropic.com/v1/messages",
        {
          model: "claude-opus-4-20250514",          // ← change if you have Opus/Haiku
          system: systemPrompt,
          messages: [{ role: "user", content: prompt }],
          max_tokens: 400
        },
        {
          headers: {
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
          },
          timeout: 30_000
        }
      );
      console.log(anthroRes.data.content[0].text);
      res.json(anthroRes.data);
    } catch (err) {
      const status = err.response?.status || 500;
      console.error("❌ Anthropic error:", err.response?.data || err.message);
      res.status(status).json({ error: err.response?.data || err.message });
    }
  });

  // ────────────────────────────── 4.  Start HTTPS server
  const PORT = 5050;
  https.createServer({ key, cert }, app).listen(PORT, () =>
    console.log(`🔐 Proxy live at https://localhost:${PORT}/anthropic`)
  );
})();
