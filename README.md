# PPTX Chart Editor / עורך גרפים

Edit charts in PowerPoint presentations with a live visual preview.

עריכת גרפים במצגות PowerPoint עם תצוגה מקדימה חיה.

---

## Features / תכונות

- Upload `.pptx` and edit chart data in-browser
- Live slide preview (before/after comparison)
- Batch add rows to all charts at once
- CSV import/export
- Hebrew / English interface

---

## Usage Options / אפשרויות שימוש

### Option 1: Use Online (Recommended)

> **[Open App](https://pptx-chart-editor.streamlit.app)**

No installation needed. Works on any device with a browser.

---

### Option 2: Self-Host with Docker

```bash
git clone https://github.com/YOUR_USER/pptx-chart-editor.git
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

## Project Structure

```
pptx-chart-editor/
├── app.py                 # Main Streamlit app
├── core/
│   ├── data_extractor.py  # Extract chart data from PPTX
│   ├── data_writer.py     # Write edited data back to PPTX
│   └── slide_renderer.py  # Render slides via LibreOffice
├── ui/
│   └── rtl_support.py     # i18n (Hebrew/English) + RTL CSS
├── requirements.txt       # Python dependencies
├── packages.txt           # System packages (Streamlit Cloud)
├── Dockerfile             # Docker image definition
└── docker-compose.yml     # Docker Compose config
```
