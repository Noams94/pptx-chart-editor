---
title: PPTX Chart Editor
emoji: 📊
colorFrom: blue
colorTo: purple
sdk: docker
app_port: 8501
pinned: false
short_description: Edit PowerPoint charts with live preview
---

# PPTX Chart Editor / עורך גרפים

Edit charts in PowerPoint presentations with a live visual preview.

עריכת גרפים במצגות PowerPoint עם תצוגה מקדימה חיה.

---

## Features / תכונות

- **Interactive Getting Started wizard** — 3-step onboarding: Upload → Edit All Charts → Edit Slide
- **Split-screen editor** — live chart preview (left) + data editor (right) when editing slides
- **Excel import/export** — bulk edit all charts via a single `.xlsx` file with one sheet per chart
- **Batch add rows** — add a new category row to all charts at once
- **CSV import/export** — per-chart data exchange
- **Before/after comparison** — toggle chart and slide comparisons side by side
- **Series visibility** — show/hide individual data series in charts
- **Interactive Plotly charts** — hover details, unified tooltips, horizontal legend
- **Auto-save** — automatic file download after saving changes
- **Percentage format preservation** — display and edit percentages naturally (67 for 67%)
- **Sidebar thumbnail navigation** — slide list with thumbnails and filtering
- **Undo support** — revert last edit per chart
- **Hebrew / English** interface with full RTL support and Assistant font
- **Modern UI** — gradient sidebar, rounded buttons, styled tabs, progress indicators
- **Responsive design** — works on desktop and tablet

---

## Usage Options / אפשרויות שימוש

### Option 1: Use Online (Recommended)

> **[Open App](https://pptx-chart-editor.streamlit.app)**

No installation needed. Works on any device with a browser.

---

### Option 2: Self-Host with Docker

```bash
git clone https://github.com/Noams94/pptx-chart-editor.git
cd pptx-chart-editor
docker compose up
```

Open [http://localhost:8501](http://localhost:8501) in your browser.

---

### Option 3: Local Development

#### 1. Install LibreOffice

| OS | Command |
|----|---------|
| macOS | `brew install --cask libreoffice` |
| Ubuntu/Debian | `sudo apt install libreoffice poppler-utils` |
| Fedora | `sudo dnf install libreoffice poppler-utils` |
| Windows | [Download from libreoffice.org](https://www.libreoffice.org/download/) |

#### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

#### 3. Run

```bash
streamlit run app.py
```

---

## Deploying to Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo and select `app.py` as the entry point
4. The `packages.txt` file automatically installs LibreOffice on the server

---

## Workflow

### Step 1: Upload
Upload a `.pptx` file via drag-and-drop or file browser.

### Step 2: Edit All Charts
Export all charts to Excel, edit in your spreadsheet tool, and import back. Add rows to all charts in batch.

### Step 3: Edit Slide
Select a slide from the grid, then edit individual charts with:
- Data editor with dynamic rows
- Live Plotly chart preview
- Series visibility controls
- CSV import/export
- Before/after comparison
- One-click save and auto-download

---

## Project Structure

```
pptx-chart-editor/
├── app.py                  # Main Streamlit app (wizard, split-screen, all features)
├── core/
│   ├── data_extractor.py   # Extract chart data from PPTX into DataFrames
│   ├── data_writer.py      # Write edited data back to PPTX (preserves formats)
│   └── slide_renderer.py   # Render slides via LibreOffice (PPTX → PDF → JPEG)
├── ui/
│   ├── rtl_support.py      # i18n (Hebrew/English, 150+ keys) + RTL CSS
│   └── chart_preview.py    # Plotly chart rendering (bar, line, pie, scatter, area)
├── .streamlit/
│   └── config.toml         # Streamlit theme and server config
├── .claude/
│   └── launch.json         # Dev server config for Claude Code
├── .devcontainer/
│   └── devcontainer.json   # Dev container config
├── requirements.txt        # Python dependencies
├── packages.txt            # System packages (Streamlit Cloud)
├── Dockerfile              # Docker image definition
└── docker-compose.yml      # Docker Compose config
```

---

## Tech Stack

- **[Streamlit](https://streamlit.io)** — Web framework
- **[python-pptx](https://python-pptx.readthedocs.io)** — PPTX manipulation
- **[Plotly](https://plotly.com/python/)** — Interactive chart previews
- **[pandas](https://pandas.pydata.org)** — Data manipulation
- **[openpyxl](https://openpyxl.readthedocs.io)** — Excel read/write
- **[LibreOffice](https://www.libreoffice.org)** — Headless slide rendering
- **[Pillow](https://pillow.readthedocs.io)** — Image processing
