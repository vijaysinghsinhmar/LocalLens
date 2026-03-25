# LocalLens

**Blazing-fast local file search** — search across hundreds of CSV, XLSX, and TXT files in milliseconds. Everything runs in your browser using parallel Web Workers. Nothing leaves your machine.

## Features

- ⚡ **Sub-second search** across hundreds of files using parallel Web Workers
- 🔒 **100% local** — no uploads, no servers, no data leaves your machine
- 🧵 **Multi-core** — automatically uses all available CPU cores
- 🔍 **Fuzzy + exact match** modes
- 📊 **File type filters** — toggle XLSX, XLS, CSV, TXT independently
- 📈 **Sort by** relevance score, filename, or row number
- 🔦 **Highlighted matches** in results
- 📋 **Row detail expand** — click any result to see the full row
- 💾 **Export to CSV** — download all matched rows

## Supported Formats

| Format | Engine |
|--------|--------|
| `.xlsx` / `.xls` | SheetJS (XLSX) |
| `.csv` | PapaParse |
| `.txt` | Native line-by-line |

## Deploy to Railway

```bash
# 1. Push to GitHub
git init && git add . && git commit -m "init"
git remote add origin <your-repo>
git push -u origin main

# 2. On Railway
# New Project → Deploy from GitHub → select repo
# Railway auto-detects Node.js and runs npm run start
```

The `PORT` environment variable is automatically set by Railway and respected by the app.

## Local Development

```bash
npm install
npm run dev
```
