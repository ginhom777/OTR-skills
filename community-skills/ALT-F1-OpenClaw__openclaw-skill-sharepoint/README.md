# openclaw-skill-sharepoint

OpenClaw skill for Microsoft SharePoint — secure file operations and Office document intelligence via Microsoft Graph API with certificate-based authentication (`Sites.Selected`).

[![ClawHub](https://img.shields.io/badge/ClawHub-sharepoint--by--altf1be-orange)](https://clawhub.ai/skills/sharepoint-by-altf1be)
[![Security](https://img.shields.io/badge/Security_Scan-Benign-green)](https://clawhub.ai/skills/sharepoint-by-altf1be)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Install via ClawHub:**
```bash
clawhub install sharepoint-by-altf1be
```

## Features

### SharePoint file operations
- **list** — browse files and folders in a SharePoint library
- **read** — download and extract text from Office documents (`.docx`, `.xlsx`, `.pptx`, `.pdf`, `.txt`, `.md`)
- **upload** — upload files to SharePoint
- **search** — find files by name
- **mkdir** — create folders
- **delete** — remove files (disabled by default, requires explicit flag)

### Office document intelligence (AI-powered via OpenClaw)
- **Meeting notes** — clean up raw notes into professional Word documents
- **Summarize** — business or technical summaries of any document
- **Action items** — extract action items into Excel trackers
- **Presentations** — generate PowerPoint slides from notes or reports

## Security model

This skill follows an **extreme security** approach:

- **Certificate-based auth only** — no client secrets
- **`Sites.Selected`** — app can only access one specific SharePoint site
- **Dedicated app registration** — single-tenant, single-purpose
- **Isolated site** — target library is in a dedicated SharePoint site
- **Code-level controls** — path validation, size limits, operation allowlists
- **No tokens in logs** — all sensitive data is redacted

## Supported file formats

| Format | Read | Write | Library |
|--------|------|-------|---------|
| `.docx` (Word) | ✅ | ✅ | `mammoth` / `docx` |
| `.xlsx` (Excel) | ✅ | ✅ | `exceljs` |
| `.pptx` (PowerPoint) | ✅ | ✅ | `pptxgenjs` |
| `.pdf` | ✅ | — | `pdf-parse` |
| `.txt` / `.md` | ✅ | ✅ | native |

## Prerequisites

- Node.js 18+
- Microsoft Entra app registration with:
  - `Sites.Selected` application permission (admin-consented)
  - Certificate-based authentication (no client secret)
  - Site-level grant (`read` or `write`) on target SharePoint site
- Certificate (`.pfx` or `.key` + `.crt`) accessible to OpenClaw

## Setup

See [references/setup-guide.md](references/setup-guide.md) for the full step-by-step secure setup guide.

## Quick start

```bash
# Install dependencies
cd openclaw-skill-sharepoint
npm install

# Configure (edit .env or set environment variables)
cp .env.example .env
# Edit .env with your tenant ID, app client ID, cert path, site ID

# Test connection
node scripts/sharepoint.mjs list
```

## Usage with OpenClaw

Once installed as a skill, you can use natural language:

> "List files in SharePoint"

> "Read the meeting notes from today"

> "Upload this report to SharePoint"

> "Summarize the quarterly review from a business perspective"

> "Extract action items from last week's meeting into Excel"

## License

MIT — see [LICENSE](LICENSE).

## Author

**Abdelkrim BOUJRAF**
[ALT-F1 SRL](https://www.alt-f1.be) — Brussels, Belgium

## Contributing

Contributions welcome! Please open an issue or PR.
