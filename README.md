
# speakSage

Interact with Microsoft Excel using natural‑language prompts powered by Anthropic Claude.

---

## Features

| Natural‑language prompt | What speakSage does |
|------------------------|----------------------|
| **“Bold the first row.”** | Executes `Excel.run` to bold row 1 |
| **“Colour cells A1:D1 yellow and centre the text.”** | Claude returns Office JS code → add‑in runs it |
| **Any prompt** | Claude returns only JavaScript wrapped in `Excel.run(...)`; the add‑in executes it |

---

## Local development stack

| Layer | Tech |
|-------|------|
| UI | React + TypeScript task‑pane (Webpack dev‑server) |
| Excel runtime | Office JS (Excel API 1.13+) |
| Claude proxy | Express HTTPS server on `https://127.0.0.1:5050` |
| Dev server | HTTPS Webpack on `https://localhost:3000` |

---

## Prerequisites

* Node ≥ 18
* npm ≥ 9
* Excel Desktop (Microsoft 365) **or** Excel Online
* Anthropic API key (`sk-ant‑…`)

---

## Quick start

```bash
git clone https://github.com/mmanav02/speakSage.git
cd speakSage
npm install
```

### 1 — Trust localhost certificate

```bash
npx office-addin-dev-certs install
```

### 2 — Start dev server + sideload

```bash
npm run dev-server       # HTTPS on 3000 & sideloads into Excel (desktop)
```

### 3 — Start Claude proxy

```bash
node proxy-server.js     # HTTPS on 5050
```

### 4 — Excel Online (optional)

Upload **manifest.xml** via *Home ▸ Add‑ins ▸ Upload My Add‑in*.

---

## Using speakSage

1. Open task pane → paste API key  
2. Enter prompt → **Run**  
3. Claude’s code appears → add‑in executes it

---

## Security notes

* Proxy keeps your key local; task‑pane never hits Anthropic directly  
* For production, deploy proxy to Vercel / Render with env vars  
* Add verb whitelist in `ExecuteCommands.ts` for extra safety

---

## License

MIT
