"""PPTX Chart Editor - Streamlit App with Enhanced UX/UI

Split-screen tool for editing PowerPoint chart data with live slide preview.
Features: modern design, progress indicators, split-screen layout,
thumbnail navigation, before/after comparison, CSV/Excel import/export,
batch row addition, Hebrew/English language switching.
"""

import base64
from collections import defaultdict
import io
from io import BytesIO
import re

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from pptx import Presentation

from core.data_extractor import extract_all_charts, is_percentage_format
from core.data_writer import update_chart_data, update_multiple_charts
from core.slide_renderer import render_slides
from ui.chart_preview import render_chart_plotly
from ui.rtl_support import t, inject_rtl_css

# --- Language Selector (must be before page config uses translated title) ---
if "lang" not in st.session_state:
    st.session_state.lang = "en"

# Page config
st.set_page_config(
    page_title=t("page_title"),
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_rtl_css()

# --- Enhanced Custom CSS for Modern UI ---
st.markdown("""
    <style>
    /* Main container padding and spacing */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 5rem;
        max-width: 1400px;
    }

    /* Typography improvements */
    h1, h2, h3 {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
        font-weight: 700;
        letter-spacing: -0.5px;
    }

    /* Sidebar styling - modern design */
    section[data-testid="stSidebar"] {
        background: linear-gradient(135deg, #f5f7fa 0%, #f0f3f7 100%);
        border-right: 1px solid #e0e6ed;
    }

    section[data-testid="stSidebar"] h2 {
        color: #1a202c;
        font-size: 1.1rem;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 700;
    }

    /* Slide button styling - improved visual hierarchy */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
        border: 1px solid transparent;
    }

    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%);
        color: white;
    }

    .stButton > button[kind="secondary"] {
        background: white;
        border: 1px solid #e0e6ed;
        color: #4b5563;
    }

    /* Data editor styling */
    .stDataEditor {
        border: 1px solid #e0e6ed;
        border-radius: 8px;
        overflow: hidden;
    }

    /* Tabs styling - modern look */
    [data-baseweb="tab-list"] {
        border-bottom: 2px solid #e0e6ed;
        gap: 2rem;
    }

    [data-baseweb="tab"] {
        font-weight: 600;
        color: #64748b;
        padding: 0.75rem 0;
    }

    [data-baseweb="tab"][aria-selected="true"] {
        color: #2563eb;
        border-bottom: 3px solid #2563eb;
    }

    /* Info/Warning/Success messages - better styling */
    .stAlert {
        border-radius: 8px;
        border-left: 4px solid;
        padding: 1rem;
    }

    /* Divider styling */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(to right, transparent, #e0e6ed, transparent);
        margin: 1.5rem 0;
    }

    /* Section headers with accent line */
    .section-header {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        margin: 1.5rem 0 1rem 0;
        padding-bottom: 0.75rem;
        border-bottom: 2px solid #2563eb;
    }

    .section-header h3 {
        margin: 0;
        color: #1a202c;
        font-size: 1.1rem;
    }

    /* Progress indicator styling */
    .progress-indicator {
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 1rem 0 1.5rem 0;
        padding: 1rem 1.5rem;
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
        border-radius: 12px;
        border: 1px solid #bae6fd;
        gap: 0;
    }

    .progress-step {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 0.4rem;
        flex: 0 0 auto;
        padding: 0 1.2rem;
        position: relative;
    }

    .progress-connector {
        width: 40px;
        height: 2px;
        background: #cbd5e1;
        flex-shrink: 0;
        margin-top: -12px;
    }

    .progress-connector.completed {
        background: #10b981;
    }

    .progress-step-number {
        width: 36px;
        height: 36px;
        border-radius: 50%;
        background: #cbd5e1;
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 0.9rem;
        transition: all 0.3s ease;
    }

    .progress-step.active .progress-step-number {
        background: #2563eb;
        box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.2);
    }

    .progress-step.completed .progress-step-number {
        background: #10b981;
    }

    .progress-step-label {
        font-size: 0.8rem;
        color: #94a3b8;
        font-weight: 600;
        text-align: center;
        white-space: nowrap;
    }

    .progress-step.active .progress-step-label {
        color: #2563eb;
        font-weight: 700;
    }

    .progress-step.completed .progress-step-label {
        color: #10b981;
    }

    /* Wizard content card */
    .wizard-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 0.75rem 1rem;
        margin: 0.5rem 0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        text-align: center;
        min-height: 126px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }

    .wizard-card h2 {
        color: #1e293b;
        margin-bottom: 0.25rem;
        font-size: 1.1rem;
    }

    .wizard-card p {
        color: #64748b;
        font-size: 0.95rem;
        line-height: 1.4;
        max-width: 600px;
        margin: 0 auto;
    }

    .wizard-icon {
        font-size: 1.5rem;
        margin-bottom: 0.25rem;
        display: block;
    }

    /* Chart container styling */
    .chart-container {
        background: white;
        border: 1px solid #e0e6ed;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }

    /* Footer */
    .fixed-footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background: white;
        border-top: 1px solid #e0e6ed;
        padding: 12px 0;
        text-align: center;
        color: #64748b;
        font-size: 0.85rem;
        z-index: 999;
        direction: ltr;
    }
    .fixed-footer a {
        color: #2563eb;
        text-decoration: none;
        font-weight: 600;
        transition: color 0.2s ease;
    }
    .fixed-footer a:hover {
        color: #1e40af;
        text-decoration: underline;
    }
    .stApp > .main { padding-bottom: 60px; }

    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Responsive adjustments */
    @media (max-width: 768px) {
        .main .block-container {
            padding-top: 1rem;
        }
        section[data-testid="stSidebar"] {
            width: 100% !important;
        }
    }
    </style>

    <div class="fixed-footer">
        &copy; All Rights Reserved &middot; Dr. Noam Keshet &middot;
        <a href="https://noamkeshet.com" target="_blank">noamkeshet.com</a> &middot;
        <a href="mailto:keshet.noam@gmail.com">keshet.noam@gmail.com</a>
    </div>
""", unsafe_allow_html=True)


# --- Helper Functions ---

def get_chart_df(chart_info):
    """Get current DataFrame for a chart (edited version if exists, otherwise original)."""
    key = chart_info.key
    if key in st.session_state.edited_data:
        return st.session_state.edited_data[key].copy()
    return chart_info.dataframe.copy()


def _apply_and_rerender(updated_bytes: bytes):
    """Save updated PPTX bytes, re-render slides, and invalidate cache."""
    st.session_state.pptx_bytes = updated_bytes
    st.session_state.slide_images = render_slides(updated_bytes)
    st.session_state.charts_cache = None


def _schedule_auto_download():
    """Mark that an auto-download should happen on the next render."""
    if st.session_state.get("auto_save", True):
        st.session_state.pending_auto_download = True


def _sanitize_sheet_name(slide_index: int, shape_name: str) -> str:
    """Create an Excel-safe sheet name: Slide{n}_{shape_name}, max 31 chars."""
    prefix = f"Slide{slide_index + 1}_"
    clean_name = re.sub(r'[\[\]:*?/\\]', '', shape_name)
    max_name_len = 31 - len(prefix)
    return prefix + clean_name[:max_name_len]


def _build_sheet_name_map(charts_list) -> dict:
    """Build a deterministic mapping of chart -> unique sheet name, handling collisions."""
    name_map = {}
    seen = set()
    for ci in charts_list:
        sheet = _sanitize_sheet_name(ci.slide_index, ci.shape_name)
        base = sheet
        counter = 1
        while sheet in seen:
            sheet = base[:29] + f"_{counter}"
            counter += 1
        seen.add(sheet)
        name_map[ci.key] = sheet
    return name_map


def _commit_update(updated_bytes: bytes):
    """Save updated PPTX, re-render, trigger auto-download, and rerun."""
    _apply_and_rerender(updated_bytes)
    _schedule_auto_download()
    st.rerun()


def _build_chart_labels(charts_list):
    """Build unique display labels for a list of charts."""
    labels = []
    seen = {}
    for ci in charts_list:
        label = f"{ci.chart_title or ci.shape_name} ({ci.chart_type_name})"
        seen[label] = seen.get(label, 0) + 1
        if seen[label] > 1:
            label = f"{label} #{seen[label]}"
        labels.append(label)
    return labels


def _render_chart_checkboxes(key_prefix: str, charts_list, labels: list) -> list:
    """Render chart selection with Select All / Clear All toggles and individual checkboxes."""
    col_sa, col_ca = st.columns(2)
    with col_sa:
        if st.button(t("select_all_charts"), key=f"{key_prefix}_sel_all_btn", use_container_width=True):
            for i in range(len(charts_list)):
                st.session_state[f"{key_prefix}_cb_{i}"] = True
            st.rerun()
    with col_ca:
        if st.button(t("clear_all_charts"), key=f"{key_prefix}_clr_all_btn", use_container_width=True):
            for i in range(len(charts_list)):
                st.session_state[f"{key_prefix}_cb_{i}"] = False
            st.rerun()

    selected = []
    num_cols = min(2, len(charts_list))
    cols = st.columns(num_cols)
    for i, (ci, label) in enumerate(zip(charts_list, labels)):
        with cols[i % num_cols]:
            if st.checkbox(label, value=st.session_state.get(f"{key_prefix}_cb_{i}", True), key=f"{key_prefix}_cb_{i}"):
                selected.append(ci)
    st.caption(t("batch_selected_count", selected=len(selected), total=len(charts_list)))
    return selected


def show_progress_indicator(current_step: int, total_steps: int = 3):
    """Display a visual progress indicator showing the current workflow step."""
    steps = [
        t("step_upload"),
        t("step_select"),
        t("step_edit"),
    ]

    parts = []
    for i, step in enumerate(steps[:total_steps], 1):
        if i < current_step:
            state_class = "completed"
        elif i == current_step:
            state_class = "active"
        else:
            state_class = ""

        if i > 1:
            conn_class = "completed" if i <= current_step else ""
            parts.append(f'<div class="progress-connector {conn_class}"></div>')

        check = "✓" if i < current_step else str(i)
        parts.append(
            f'<div class="progress-step {state_class}">'
            f'<div class="progress-step-number">{check}</div>'
            f'<div class="progress-step-label">{step}</div>'
            f'</div>'
        )

    html = '<div class="progress-indicator">' + ''.join(parts) + '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_user_guide():
    """Render the bilingual user guide content."""
    st.markdown(f"### {t('guide_overview_title')}")
    st.markdown(t("guide_overview_body"))

    st.markdown(f"### {t('guide_start_title')}")
    st.markdown(f"1. {t('guide_start_1')}")
    st.markdown(f"2. {t('guide_start_2')}")
    st.markdown(f"3. {t('guide_start_3')}")
    st.markdown(f"4. {t('guide_start_4')}")
    st.markdown(f"5. {t('guide_start_5')}")

    st.markdown(f"### {t('guide_edit_title')}")
    st.markdown(t("guide_edit_body"))

    st.markdown(f"### {t('guide_excel_title')}")
    st.markdown(t("guide_excel_body"))

    st.markdown(f"### {t('guide_csv_title')}")
    st.markdown(t("guide_csv_body"))

    st.markdown(f"### {t('guide_batch_title')}")
    st.markdown(t("guide_batch_body"))

    st.markdown(f"### {t('guide_batch_col_title')}")
    st.markdown(t("guide_batch_col_body"))

    st.markdown(f"### {t('guide_visibility_title')}")
    st.markdown(t("guide_visibility_body"))

    st.markdown(f"### {t('guide_tips_title')}")
    st.markdown(t("guide_tips_body"))


def show_interactive_wizard():
    """Display the interactive Getting Started wizard with clickable steps.

    Step 0 = Welcome landing (no progress active)
    Step 1 = Upload explanation (step 1 active)
    Step 2 = Select explanation (step 2 active)
    Step 3 = Edit explanation (step 3 active)
    Step 4 = Download explanation (step 4 active)
    """
    if "wizard_step" not in st.session_state:
        st.session_state.wizard_step = 0

    step = st.session_state.wizard_step

    # Welcome title + description on landing (step 0)
    if step == 0:
        st.markdown(f"### 👋 {t('wizard_welcome_title')}")
        st.markdown(f"*{t('wizard_welcome_desc')}*")

    # Render progress bar (step 0 means none active)
    show_progress_indicator(step)

    # --- Step content ---
    if step == 0:
        # Landing — just Next button
        _, col_next, _ = st.columns([2, 1, 2])
        with col_next:
            if st.button(t("wizard_next"), type="primary", use_container_width=True, key="wiz_next_0"):
                st.session_state.wizard_step = 1
                st.rerun()

    elif step == 1:
        # Show already uploaded file if exists
        if "file_name" in st.session_state:
            st.success(f"✅ {t('wizard_upload_done', name=st.session_state.file_name)}")

        # File uploader
        wizard_file = st.file_uploader(
                t("upload_label"),
                type=["pptx"],
                help=t("upload_help"),
                key="wizard_main_uploader",
                label_visibility="collapsed",
            )
        if wizard_file:
            # Only initialize on first upload (avoid rerun loop)
            if st.session_state.get("file_name") != wizard_file.name:
                st.session_state.pptx_bytes = wizard_file.getvalue()
                st.session_state.file_name = wizard_file.name
                st.session_state.slide_images = None
                st.session_state.original_slide_images = None
                st.session_state.edited_data = {}
                st.session_state.selected_slide = None
                st.session_state.show_chart_comparison = False
                st.session_state.show_slide_comparison = False
                st.session_state.original_charts = None
                st.session_state.charts_cache = None
                st.session_state.series_visibility = {}
                st.session_state.undo_stack = []
                st.rerun()
            st.success(f"✅ {t('wizard_upload_done', name=wizard_file.name)}")

        col_back, _, col_next = st.columns([1, 2, 1])
        with col_back:
            if st.button(t("wizard_back"), use_container_width=True, key="wiz_back_1"):
                st.session_state.wizard_step = 0
                st.rerun()
        with col_next:
            if st.button(t("wizard_next"), type="primary", use_container_width=True, key="wiz_next_1"):
                st.session_state.wizard_step = 2
                st.rerun()

    elif step == 2:
        # Edit all charts (Excel)
        st.markdown(
            '<div class="wizard-card">'
            '<span class="wizard-icon">📊</span>'
            f'<h2>{t("wizard_select_title")}</h2>'
            f'<p>{t("wizard_select_desc")}</p>'
            '</div>',
            unsafe_allow_html=True,
        )

        # Edit single slide
        st.markdown(
            '<div class="wizard-card">'
            '<span class="wizard-icon">🖼️</span>'
            f'<h2>{t("wizard_edit_title")}</h2>'
            f'<p>{t("wizard_edit_desc")}</p>'
            '</div>',
            unsafe_allow_html=True,
        )

        col_back, _, col_start = st.columns([1, 2, 1])
        with col_back:
            if st.button(t("wizard_back"), use_container_width=True, key="wiz_back_2"):
                st.session_state.wizard_step = 1
                st.rerun()
        with col_start:
            if st.button(t("wizard_start_editing"), type="primary", use_container_width=True, key="wiz_start"):
                st.session_state.wizard_step = 0
                st.rerun()


# --- Header with Language Selector ---
col_title, col_lang = st.columns([4, 1])
with col_title:
    st.title(t("page_title"))
    st.markdown(f"*{t('app_subtitle')}*")
with col_lang:
    lang_options = {"עברית 🇮🇱": "he", "English 🇺🇸": "en"}
    current_label = "עברית 🇮🇱" if st.session_state.lang == "he" else "English 🇺🇸"
    selected_lang_label = st.selectbox(
        "Language / שפה",
        options=list(lang_options.keys()),
        index=list(lang_options.keys()).index(current_label),
        label_visibility="collapsed",
    )
    new_lang = lang_options[selected_lang_label]
    if new_lang != st.session_state.lang:
        st.session_state.lang = new_lang
        st.session_state.charts_cache = None
        st.rerun()

st.divider()

# --- Sidebar with Enhanced Design ---
with st.sidebar:
    # Show "Quick Start" title only on landing screen (no file uploaded yet)
    if "pptx_bytes" not in st.session_state:
        st.markdown(f"### ⚡ {t('quick_start')}")
        with st.expander(f"📖 {t('tab_guide')}"):
            render_user_guide()

    st.header(f"📁 {t('upload_section_title')}")
    uploaded_file = st.file_uploader(
        t("upload_label"),
        type=["pptx"],
        help=t("upload_help"),
        label_visibility="collapsed"
    )
    if "pptx_bytes" in st.session_state:
        st.warning(t("sidebar_upload_override_warning"))

    # Show sidebar controls if file is loaded (from sidebar OR wizard uploader)
    if uploaded_file or "pptx_bytes" in st.session_state:
        st.session_state.setdefault("auto_save", True)
        auto_save_enabled = st.checkbox(
            t("auto_save_label"),
            value=st.session_state.auto_save,
            help=t("auto_save_info"),
        )

        # User Guide expander
        with st.expander(f"📖 {t('tab_guide')}"):
            render_user_guide()

        # Only show SLIDES section in step 3
        if st.session_state.get("show_step3", False):
            st.divider()
            st.header(f"🖼️ {t('slides')}")

            # Slide filter
            slide_filter = st.text_input(
                t("filter_slides"),
                placeholder=t("filter_slides"),
                label_visibility="collapsed",
                help=t("filter_help"),
            )
        else:
            slide_filter = ""
    else:
        slide_filter = ""

# Check if we have a file from either the sidebar uploader or the wizard uploader
has_file = uploaded_file is not None or "pptx_bytes" in st.session_state

if not has_file:
    # No file from any source — show wizard
    show_interactive_wizard()
    st.stop()

# --- Initialize Session State ---
# If the sidebar uploader has a (new) file, use it
if uploaded_file is not None and st.session_state.get("file_name") != uploaded_file.name:
    st.session_state.pptx_bytes = uploaded_file.getvalue()
    st.session_state.file_name = uploaded_file.name
    st.session_state.slide_images = None
    st.session_state.original_slide_images = None
    st.session_state.edited_data = {}
    st.session_state.selected_slide = None
    st.session_state.show_chart_comparison = False
    st.session_state.show_slide_comparison = False
    st.session_state.original_charts = None
    st.session_state.charts_cache = None
    st.session_state.series_visibility = {}
    st.session_state.undo_stack = []
# If no sidebar file but we have pptx_bytes from wizard, that's already initialized

# --- Auto-download trigger (fires after rerun following an update) ---
if st.session_state.pop("pending_auto_download", False):
    b64 = base64.b64encode(st.session_state.pptx_bytes).decode()
    components.html(
        f"""<script>
        const link = document.createElement('a');
        link.href = 'data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}';
        link.download = 'updated_{st.session_state.file_name}';
        link.click();
        </script>""",
        height=0,
    )
    st.toast(t("auto_saved_msg"), icon="✅")

# --- Extract Charts (cached in session state) ---
if st.session_state.get("charts_cache") is None:
    with st.spinner(t("rendering")):
        st.session_state.charts_cache = extract_all_charts(st.session_state.pptx_bytes)
        if st.session_state.original_charts is None:
            st.session_state.original_charts = {
                c.key: c.dataframe.copy()
                for c in st.session_state.charts_cache
            }

charts = st.session_state.charts_cache

if not charts:
    st.warning(t("no_charts"))
    st.stop()

# --- Group charts by slide ---
charts_by_slide = defaultdict(list)
for c in charts:
    charts_by_slide[c.slide_index].append(c)

# --- Sidebar: Slide List with Thumbnails (only in step 3) ---
slide_images = st.session_state.slide_images or []

if st.session_state.get("show_step3", False):
    with st.sidebar:
        st.caption(t("slide_count_info", count=len(charts_by_slide)))

        for slide_idx in sorted(charts_by_slide):
            if slide_filter:
                slide_num_str = str(slide_idx + 1)
                chart_names = " ".join(c.chart_title or c.shape_name for c in charts_by_slide[slide_idx]).lower()
                if slide_filter.lower() not in slide_num_str and slide_filter.lower() not in chart_names:
                    continue

            is_selected = st.session_state.selected_slide == slide_idx
            btn_label = t("slide_n_charts", n=slide_idx + 1, count=len(charts_by_slide[slide_idx]))

            if st.button(
                btn_label,
                key=f"slide_btn_{slide_idx}",
                use_container_width=True,
                type="primary" if is_selected else "secondary"
            ):
                st.session_state.selected_slide = slide_idx
                st.rerun()

            # Show thumbnail if available
            if slide_idx < len(slide_images):
                st.image(slide_images[slide_idx], use_container_width=True)
            st.divider()


# ==================== MAIN CONTENT AREA ====================

if not st.session_state.get("show_step3", False):
    # --- Step 2: Edit all charts via Excel ---
    show_progress_indicator(2)

    st.success(f"✅ {t('wizard_upload_done', name=st.session_state.file_name)}")

    # Overview metrics
    total_slide_count = len(Presentation(BytesIO(st.session_state.pptx_bytes)).slides)
    overview_cols = st.columns(3)
    with overview_cols[0]:
        st.metric(t("total_slides"), f"{len(charts_by_slide)} / {total_slide_count}")
    with overview_cols[1]:
        st.metric(t("total_charts"), len(charts))
    with overview_cols[2]:
        st.metric(t("chart_types"), len(set(c.chart_type_name for c in charts)))

    st.divider()

    # --- STEP 2: EXCEL IMPORT/EXPORT (ALL CHARTS) ---
    st.subheader(f"📊 {t('tab_excel')}")

    col_export_xl, col_import_xl = st.columns(2, gap="large")

    # --- Export ---
    with col_export_xl:
        st.markdown(f"**{t('excel_export_title')}**")
        st.caption(t("excel_export_caption", count=len(charts)))

        sheet_name_map = _build_sheet_name_map(charts)

        edit_fingerprint = str(sorted(st.session_state.edited_data.keys()))
        if (st.session_state.get("xl_export_fingerprint") != edit_fingerprint
                or "xl_export_bytes" not in st.session_state):
            xl_buffer = io.BytesIO()
            with pd.ExcelWriter(xl_buffer, engine="openpyxl") as writer:
                for chart_info in charts:
                    sheet = sheet_name_map[chart_info.key]
                    get_chart_df(chart_info).to_excel(writer, sheet_name=sheet, index=False)
            st.session_state.xl_export_bytes = xl_buffer.getvalue()
            st.session_state.xl_export_fingerprint = edit_fingerprint

        base_name = st.session_state.file_name.replace(".pptx", "")
        st.download_button(
            label=t("excel_export_button"),
            data=st.session_state.xl_export_bytes,
            file_name=f"charts_{base_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # --- Import ---
    with col_import_xl:
        st.markdown(f"**{t('excel_import_title')}**")
        st.caption(t("excel_import_caption"))

        xl_file = st.file_uploader(
            t("excel_import_upload_label"),
            type=["xlsx"],
            key="excel_import_all",
        )

    # --- Import results (full width, below both columns) ---
    if xl_file is not None:
        st.divider()
        try:
            xl_cache_key = (xl_file.name, xl_file.size)
            if st.session_state.get("xl_import_cache_key") != xl_cache_key:
                xls = pd.ExcelFile(xl_file, engine="openpyxl")

                sheet_to_chart = {v: k for k, v in sheet_name_map.items()}
                charts_by_key = {ci.key: ci for ci in charts}

                changed = []
                unchanged = 0
                skipped = []
                for sheet_name in xls.sheet_names:
                    chart_key = sheet_to_chart.get(sheet_name)
                    if chart_key and chart_key in charts_by_key:
                        ci = charts_by_key[chart_key]
                        imported_df = pd.read_excel(xls, sheet_name=sheet_name)
                        expected_cols = len(ci.dataframe.columns)
                        if len(imported_df.columns) != expected_cols:
                            skipped.append(t("excel_column_mismatch_warning",
                                             sheet=sheet_name, expected=expected_cols,
                                             found=len(imported_df.columns)))
                        else:
                            imported_df.columns = ci.dataframe.columns
                            current_df = get_chart_df(ci)
                            if not imported_df.equals(current_df):
                                changed.append((ci, imported_df))
                            else:
                                unchanged += 1
                    else:
                        skipped.append(t("excel_sheet_no_match", sheet=sheet_name))

                st.session_state.xl_import_cache_key = xl_cache_key
                st.session_state.xl_import_results = (changed, unchanged, skipped)

            changed, unchanged, skipped = st.session_state.xl_import_results

            if changed:
                st.success(t("excel_changes_found", changed=len(changed), total=len(changed) + unchanged))
                if unchanged:
                    st.caption(t("excel_unchanged", count=unchanged))

                for ci, df in changed:
                    st.markdown(f"**{t('slide_num')} {ci.slide_index + 1} — {ci.shape_name}**")
                    st.dataframe(df, use_container_width=True)

                if skipped:
                    with st.expander(f"⚠️ {len(skipped)}", expanded=False):
                        for msg in skipped:
                            st.warning(msg)

                if st.button(t("excel_apply_button"), type="primary", use_container_width=True):
                    with st.spinner(t("excel_apply_spinner", count=len(changed))):
                        updates = []
                        for ci, df in changed:
                            st.session_state.edited_data[ci.key] = df
                            updates.append((
                                ci.slide_index,
                                ci.shape_name,
                                df,
                                ci.is_xy,
                                ci.series_formats,
                                None,
                                ci.shape_id,
                            ))
                        updated_bytes = update_multiple_charts(
                            st.session_state.pptx_bytes, updates,
                        )
                        st.success(t("excel_apply_success", count=len(changed)))
                        _commit_update(updated_bytes)
            elif unchanged > 0:
                st.info(t("excel_no_changes"))
            elif xls.sheet_names:
                st.error(t("excel_no_matches"))
                for msg in skipped:
                    st.warning(msg)
        except Exception as e:
            st.error(t("file_read_error", e=e))

    # --- BATCH ADD ROW ---
    st.divider()
    st.subheader(f"➕ {t('tab_batch')}")
    st.caption(t("batch_caption"))

    batch_labels = _build_chart_labels(charts)
    selected_row_charts = _render_chart_checkboxes("batch_row_s2", charts, batch_labels)

    new_category = st.text_input(t("new_category_label"), key="batch_category")

    if new_category:
        if not selected_row_charts:
            st.info(t("batch_no_charts_selected"))
        else:
            st.markdown(f"**{t('batch_preview', name=new_category, count=len(selected_row_charts))}**")

            if st.button(t("batch_button"), type="primary", use_container_width=True, key="batch_btn_step2"):
                with st.spinner(t("batch_spinner")):
                    try:
                        updates = []
                        for chart_info in selected_row_charts:
                            df = get_chart_df(chart_info)
                            new_row = {df.columns[0]: new_category}
                            for col in df.columns[1:]:
                                new_row[col] = None
                            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                            st.session_state.edited_data[chart_info.key] = df
                            updates.append((chart_info.slide_index, chart_info.shape_name, df, chart_info.is_xy, chart_info.series_formats, None, chart_info.shape_id))
                        updated_bytes = update_multiple_charts(st.session_state.pptx_bytes, updates)
                        st.success(t("batch_success", name=new_category, count=len(selected_row_charts)))
                        _commit_update(updated_bytes)
                    except Exception as e:
                        st.error(t("error_generic", e=e))

    # --- BATCH ADD COLUMN ---
    st.divider()
    st.subheader(f"➕ {t('tab_batch_col')}")
    st.caption(t("batch_col_caption"))

    selected_col_charts = _render_chart_checkboxes("batch_col_s2", charts, batch_labels)

    new_series = st.text_input(t("new_series_label"), key="batch_series_step2")

    if new_series:
        if not selected_col_charts:
            st.info(t("batch_no_charts_selected"))
        else:
            # Warn if column already exists in some charts
            conflicts = [ci for ci in selected_col_charts if new_series in get_chart_df(ci).columns]
            if conflicts:
                st.warning(t("batch_col_exists_warning", name=new_series, count=len(conflicts)))

            st.markdown(f"**{t('batch_col_preview', name=new_series, count=len(selected_col_charts))}**")

            if st.button(t("batch_col_button"), type="primary", use_container_width=True, key="batch_col_btn_step2"):
                with st.spinner(t("batch_col_spinner")):
                    try:
                        updates = []
                        for chart_info in selected_col_charts:
                            df = get_chart_df(chart_info)
                            df[new_series] = None
                            st.session_state.edited_data[chart_info.key] = df

                            # Inherit format from last existing series
                            updated_formats = dict(chart_info.series_formats)
                            last_format = list(updated_formats.values())[-1] if updated_formats else "General"
                            updated_formats[new_series] = last_format

                            # New column visible by default
                            updated_visibility = dict(chart_info.series_visibility)
                            updated_visibility[new_series] = True

                            updates.append((
                                chart_info.slide_index,
                                chart_info.shape_name,
                                df,
                                chart_info.is_xy,
                                updated_formats,
                                updated_visibility,
                                chart_info.shape_id,
                            ))
                        updated_bytes = update_multiple_charts(st.session_state.pptx_bytes, updates)
                        st.success(t("batch_col_success", name=new_series, count=len(selected_col_charts)))
                        _commit_update(updated_bytes)
                    except Exception as e:
                        st.error(t("error_generic", e=e))

    # --- Navigation buttons ---
    st.divider()
    col_back, _, col_next = st.columns([1, 2, 1])
    with col_back:
        if st.button(t("wizard_back"), use_container_width=True, key="step2_back"):
            # Go back to step 1 — keep the uploaded file
            st.session_state.wizard_step = 1
            st.rerun()
    with col_next:
        if st.button(t("wizard_next"), type="primary", use_container_width=True, key="step2_next"):
            st.session_state.show_step3 = True
            st.rerun()

    # If user hasn't clicked Next to step 3, stop here
    if not st.session_state.get("show_step3", False):
        st.stop()


# ==================== STEP 3: SELECTED SLIDE VIEW - SPLIT SCREEN ====================

show_progress_indicator(3)

# Show slide selection UI if none selected yet
if st.session_state.selected_slide is None:
    # Render preview button above the slides grid
    if st.session_state.slide_images is None:
        render_col, _ = st.columns([1, 3])
        with render_col:
            if st.button(t("render_preview_btn"), type="primary", use_container_width=True, key="step3_render"):
                with st.spinner(t("rendering")):
                    try:
                        st.session_state.slide_images = render_slides(st.session_state.pptx_bytes)
                        st.session_state.original_slide_images = list(st.session_state.slide_images)
                        st.rerun()
                    except RuntimeError as e:
                        st.error(str(e))
        st.caption(t("render_hint"))

    # Show slide selection grid
    st.subheader(f"🖼️ {t('slides')}")
    st.caption(t("no_slide_selected"))

    num_cols = min(4, len(charts_by_slide))
    slide_cols = st.columns(num_cols, gap="medium")
    for col_i, s_idx in enumerate(sorted(charts_by_slide)):
        with slide_cols[col_i % num_cols]:
            btn_label = t("slide_n_charts", n=s_idx + 1, count=len(charts_by_slide[s_idx]))
            if st.button(btn_label, key=f"main_slide_btn_{s_idx}", use_container_width=True, type="secondary"):
                st.session_state.selected_slide = s_idx
                st.rerun()
            # Show thumbnail if available
            s_images = st.session_state.slide_images or []
            if s_idx < len(s_images):
                st.image(s_images[s_idx], use_container_width=True)

    st.divider()
    col_back3, _, _ = st.columns([1, 2, 1])
    with col_back3:
        if st.button(t("wizard_back"), use_container_width=True, key="step3_back_noselection"):
            st.session_state.show_step3 = False
            st.rerun()
    st.stop()

slide_idx = st.session_state.selected_slide
slide_charts = charts_by_slide[slide_idx]

# Back button to return to step 2
col_back_s3, _ = st.columns([1, 5])
with col_back_s3:
    if st.button(f"← {t('wizard_back')}", use_container_width=True, key="step3_back_editing"):
        st.session_state.selected_slide = None
        st.session_state.show_step3 = False
        st.rerun()

# Layout: Left (Preview) | Right (Editor)
col_preview, col_editor = st.columns([1, 1], gap="large")

# --- LEFT COLUMN: Preview ---
with col_preview:
    st.markdown(f"### 👁️ {t('preview_section_title')}")

    # Slide image preview
    if st.session_state.slide_images is None:
        if st.button(t("render_preview_btn"), type="primary", use_container_width=True):
            with st.spinner(t("rendering")):
                try:
                    st.session_state.slide_images = render_slides(st.session_state.pptx_bytes)
                    st.session_state.original_slide_images = list(st.session_state.slide_images)
                    st.rerun()
                except RuntimeError as e:
                    st.error(str(e))
        st.caption(t("render_hint"))
    else:
        if slide_idx < len(st.session_state.slide_images):
            st.image(
                st.session_state.slide_images[slide_idx],
                use_container_width=True,
                caption=f"{t('slide_num')} {slide_idx + 1}",
            )
        if st.button(t("update_preview"), use_container_width=True):
            with st.spinner(t("rendering")):
                try:
                    st.session_state.slide_images = render_slides(st.session_state.pptx_bytes)
                    st.rerun()
                except Exception as e:
                    st.error(f"{t('error_render')}: {e}")

    st.divider()

    # Chart Selection
    st.markdown(f"#### {t('select_chart')}")
    chart_labels = []
    seen_labels = {}
    for i, c in enumerate(slide_charts):
        label = f"{c.chart_title or c.shape_name} ({c.chart_type_name})"
        seen_labels[label] = seen_labels.get(label, 0) + 1
        if seen_labels[label] > 1:
            label = f"{label} #{seen_labels[label]}"
        chart_labels.append(label)
    chart_options = {label: i for i, label in enumerate(chart_labels)}

    if not chart_options:
        st.info(t("no_charts_in_slide"))
        st.stop()

    selected_label = st.selectbox(
        t("select_chart"),
        options=list(chart_options.keys()),
        label_visibility="collapsed",
    )
    selected_idx = chart_options[selected_label]
    selected_chart = slide_charts[selected_idx]

    chart_key = selected_chart.key
    current_df = get_chart_df(selected_chart)
    current_vis = st.session_state.get("series_visibility", {}).get(
        chart_key, selected_chart.series_visibility
    )

    # Plotly Chart Preview with Before/After
    col_chart_title, col_chart_toggle = st.columns([3, 1])
    with col_chart_title:
        st.markdown(f"**{t('chart_preview')}**")
    with col_chart_toggle:
        show_chart_comparison = st.checkbox(
            t("chart_comparison_toggle"),
            value=st.session_state.show_chart_comparison,
            key="chart_comp_toggle",
        )
        st.session_state.show_chart_comparison = show_chart_comparison

    if show_chart_comparison:
        chart_col_before, chart_col_after = st.columns(2, gap="medium")
        with chart_col_before:
            st.caption(t("before"))
            original_df = st.session_state.original_charts.get(chart_key)
            if original_df is not None:
                fig_before = render_chart_plotly(
                    original_df, selected_chart.chart_type,
                    selected_chart.series_visibility, selected_chart.series_formats,
                )
                st.plotly_chart(fig_before, use_container_width=True, key=f"plotly_before_{chart_key}")
        with chart_col_after:
            st.caption(t("after"))
            fig_after = render_chart_plotly(
                current_df, selected_chart.chart_type,
                current_vis, selected_chart.series_formats,
            )
            st.plotly_chart(fig_after, use_container_width=True, key=f"plotly_after_{chart_key}")
    else:
        fig_current = render_chart_plotly(
            current_df, selected_chart.chart_type,
            current_vis, selected_chart.series_formats,
        )
        st.plotly_chart(fig_current, use_container_width=True, key=f"plotly_current_{chart_key}")

    # Full Slide Before/After comparison
    col_slide_title, col_slide_toggle = st.columns([3, 1])
    with col_slide_title:
        st.markdown(f"**{t('full_slide_preview')}**")
    with col_slide_toggle:
        show_slide_comparison = st.checkbox(
            t("slide_comparison_toggle"),
            value=st.session_state.show_slide_comparison,
            key="slide_comp_toggle",
        )
        st.session_state.show_slide_comparison = show_slide_comparison

    if show_slide_comparison:
        slide_col_before, slide_col_after = st.columns(2, gap="medium")
        with slide_col_before:
            st.caption(t("before"))
            original_images = st.session_state.original_slide_images or []
            if original_images and slide_idx < len(original_images):
                st.image(original_images[slide_idx], use_container_width=True)
        with slide_col_after:
            st.caption(t("after"))
            if slide_images and slide_idx < len(slide_images):
                st.image(slide_images[slide_idx], use_container_width=True,
                         caption=f"{t('slide_num')} {slide_idx + 1}")
            else:
                st.info(t("render_hint"))
    else:
        # Only show if not already shown above
        pass


# --- RIGHT COLUMN: Editor ---
with col_editor:
    st.markdown(f"### 📝 {t('editor_section_title')}")

    tab_edit, tab_select_data, tab_csv, tab_batch, tab_batch_col = st.tabs([t("tab_edit"), t("tab_select_data"), t("tab_csv"), t("tab_batch"), t("tab_batch_col")])

    # --- EDIT CHART TAB ---
    with tab_edit:
        st.markdown(f"**{selected_chart.chart_title or selected_chart.shape_name}**")
        st.caption(f"{t('chart_type')}: {selected_chart.chart_type_name}")

        pct_cols = [
            col for col in selected_chart.dataframe.columns[1:]
            if is_percentage_format(selected_chart.series_formats.get(col, ""))
        ]
        if pct_cols:
            st.info(t("pct_columns_info", cols=", ".join(pct_cols)))

        st.caption(t("editing_info"))

        # Build column config with widths based on content
        _edit_df = get_chart_df(selected_chart)
        _col_config = {}
        for col in _edit_df.columns:
            max_len = max(len(str(col)), _edit_df[col].astype(str).str.len().max() if len(_edit_df) > 0 else 0)
            if max_len <= 6:
                _col_config[col] = st.column_config.Column(width="small")
            elif max_len <= 15:
                _col_config[col] = st.column_config.Column(width="medium")
            else:
                _col_config[col] = st.column_config.Column(width="large")

        editor_key = f"editor_{selected_chart.slide_index}_{selected_chart.shape_id}"
        edited_df = st.data_editor(
            _edit_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config=_col_config,
            key=editor_key,
            height=300,
        )

        # Push to undo stack if data changed
        prev_df = st.session_state.edited_data.get(chart_key)
        if prev_df is not None and not edited_df.equals(prev_df):
            st.session_state.setdefault("undo_stack", []).append((chart_key, prev_df.copy()))
        st.session_state.edited_data[chart_key] = edited_df

        has_unsaved = not edited_df.equals(selected_chart.dataframe)
        if has_unsaved:
            st.warning(t("unsaved_warning"))

        # Action buttons
        c1, c2, c3 = st.columns(3)
        with c1:
            undo_stack = st.session_state.get("undo_stack", [])
            if st.button(t("undo"), use_container_width=True, disabled=len(undo_stack) == 0):
                if undo_stack:
                    undo_key, undo_df = undo_stack.pop()
                    st.session_state.edited_data[undo_key] = undo_df
                    st.success(t("undo_success"))
                    st.rerun()
        with c2:
            if st.button(t("save_to_pptx"), type="primary", use_container_width=True,
                         disabled=not has_unsaved,
                         help=None if has_unsaved else t("save_disabled_hint")):
                with st.spinner(t("saving_to_pptx")):
                    try:
                        updated_bytes = update_chart_data(
                            st.session_state.pptx_bytes,
                            selected_chart.slide_index,
                            selected_chart.shape_name,
                            edited_df,
                            selected_chart.is_xy,
                            selected_chart.series_formats,
                            current_vis,
                            selected_chart.shape_id,
                        )
                        st.session_state.pptx_bytes = updated_bytes
                        st.session_state.charts_cache = None
                        st.session_state.edited_data.pop(chart_key, None)
                        _schedule_auto_download()
                        st.success(t("saved_to_pptx"))
                        st.rerun()
                    except Exception as e:
                        st.error(f"{t('error_render')}: {e}")
        with c3:
            st.download_button(
                label=t("download"),
                data=st.session_state.pptx_bytes,
                file_name=f"updated_{st.session_state.file_name}",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

    # --- SELECT DATA TAB (Series Visibility) ---
    with tab_select_data:
        st.subheader(t("tab_select_data"))
        st.caption(t("select_data_caption"))

        vis_key = selected_chart.key
        if "series_visibility" not in st.session_state:
            st.session_state.series_visibility = {}
        if vis_key not in st.session_state.series_visibility:
            st.session_state.series_visibility[vis_key] = dict(selected_chart.series_visibility)

        current_visibility = st.session_state.series_visibility[vis_key]

        if selected_chart.is_xy:
            series_display_names = selected_chart.series_names
        else:
            series_display_names = list(selected_chart.series_visibility.keys())

        st.markdown(f"**{t('series_visible_label')}**")
        updated_visibility = {}
        for name in series_display_names:
            visible = current_visibility.get(name, True)
            updated_visibility[name] = st.checkbox(
                name,
                value=visible,
                key=f"vis_{selected_chart.slide_index}_{selected_chart.shape_id}_{name}",
            )

        visible_count = sum(1 for v in updated_visibility.values() if v)
        hidden_count = len(updated_visibility) - visible_count

        if hidden_count > 0:
            st.info(t("hidden_series_count", count=hidden_count))
        else:
            st.success(t("all_series_visible"))

        st.session_state.series_visibility[vis_key] = updated_visibility

        if visible_count == 0:
            st.error(t("at_least_one_series"))
        elif st.button(t("update_visibility"), type="primary", use_container_width=True):
            with st.spinner(t("rendering")):
                try:
                    edited_df = get_chart_df(selected_chart)
                    updated_bytes = update_chart_data(
                        st.session_state.pptx_bytes,
                        selected_chart.slide_index,
                        selected_chart.shape_name,
                        edited_df,
                        selected_chart.is_xy,
                        selected_chart.series_formats,
                        updated_visibility,
                        selected_chart.shape_id,
                    )
                    st.success(t("visibility_updated"))
                    _commit_update(updated_bytes)
                except Exception as e:
                    st.error(f"{t('error_render')}: {e}")

    # --- CSV IMPORT/EXPORT TAB ---
    with tab_csv:
        st.subheader(t("tab_csv"))
        st.caption(t("csv_help"))

        col_export, col_import = st.columns(2, gap="large")

        with col_export:
            st.markdown(f"**{t('export_title')}**")
            st.caption(t("chart_info", name=selected_chart.shape_name, slide=selected_chart.slide_index + 1))

            export_df = get_chart_df(selected_chart)
            csv_buffer = io.StringIO()
            export_df.to_csv(csv_buffer, index=False, encoding="utf-8")
            csv_bytes = ("\ufeff" + csv_buffer.getvalue()).encode("utf-8")

            st.download_button(
                label=t("export_button"),
                data=csv_bytes,
                file_name=f"{selected_chart.shape_name}_slide{selected_chart.slide_index + 1}.csv",
                mime="text/csv",
                use_container_width=True,
            )

        with col_import:
            st.markdown(f"**{t('import_title')}**")
            st.caption(t("import_info"))

            csv_file = st.file_uploader(
                t("import_upload_label"),
                type=["csv"],
                key=f"csv_import_{selected_chart.slide_index}_{selected_chart.shape_id}",
            )

            if csv_file is not None:
                try:
                    imported_df = pd.read_csv(csv_file, encoding="utf-8-sig")

                    expected_cols = len(selected_chart.dataframe.columns)
                    if len(imported_df.columns) != expected_cols:
                        st.error(t("column_mismatch", expected=expected_cols, found=len(imported_df.columns)))
                    else:
                        imported_df.columns = selected_chart.dataframe.columns

                        st.markdown(f"**{t('preview_heading')}**")
                        st.dataframe(imported_df, use_container_width=True)

                        if st.button(t("apply_imported"), type="primary", use_container_width=True):
                            st.session_state.edited_data[chart_key] = imported_df

                            with st.spinner(t("applying_csv")):
                                csv_vis = st.session_state.get("series_visibility", {}).get(
                                    chart_key, selected_chart.series_visibility
                                )
                                updated_bytes = update_chart_data(
                                    st.session_state.pptx_bytes,
                                    selected_chart.slide_index,
                                    selected_chart.shape_name,
                                    imported_df,
                                    selected_chart.is_xy,
                                    selected_chart.series_formats,
                                    csv_vis,
                                    selected_chart.shape_id,
                                )
                                st.success(t("import_success"))
                                _commit_update(updated_bytes)
                except Exception as e:
                    st.error(t("file_read_error", e=e))

    # --- BATCH ADD ROW TAB ---
    with tab_batch:
        st.subheader(t("tab_batch"))
        st.caption(t("batch_caption"))

        # Chart selection
        tab_batch_labels = _build_chart_labels(charts)
        batch_row_charts = _render_chart_checkboxes("batch_row_s3", charts, tab_batch_labels)

        new_category = st.text_input(t("new_category_label"), key="batch_category")

        if new_category:
            if not batch_row_charts:
                st.info(t("batch_no_charts_selected"))
            else:
                st.markdown(f"**{t('batch_preview', name=new_category, count=len(batch_row_charts))}**")

                if st.button(t("batch_button"), type="primary", use_container_width=True):
                    with st.spinner(t("batch_spinner")):
                        try:
                            updates = []
                            for chart_info in batch_row_charts:
                                df = get_chart_df(chart_info)

                                new_row = {df.columns[0]: new_category}
                                for col in df.columns[1:]:
                                    new_row[col] = None
                                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

                                st.session_state.edited_data[chart_info.key] = df
                                updates.append((
                                    chart_info.slide_index,
                                    chart_info.shape_name,
                                    df,
                                    chart_info.is_xy,
                                    chart_info.series_formats,
                                    None,
                                    chart_info.shape_id,
                                ))

                            updated_bytes = update_multiple_charts(
                                st.session_state.pptx_bytes, updates,
                            )
                            st.success(t("batch_success", name=new_category, count=len(batch_row_charts)))
                            _commit_update(updated_bytes)
                        except Exception as e:
                            st.error(t("error_generic", e=e))

    # --- BATCH ADD COLUMN TAB ---
    with tab_batch_col:
        st.subheader(t("tab_batch_col"))
        st.caption(t("batch_col_caption"))

        # Chart selection
        tab_col_labels = _build_chart_labels(charts)
        batch_col_charts = _render_chart_checkboxes("batch_col_s3", charts, tab_col_labels)

        new_series = st.text_input(t("new_series_label"), key="batch_series")

        if new_series:
            if not batch_col_charts:
                st.info(t("batch_no_charts_selected"))
            else:
                # Warn if column already exists in some charts
                conflicts = [ci for ci in batch_col_charts if new_series in get_chart_df(ci).columns]
                if conflicts:
                    st.warning(t("batch_col_exists_warning", name=new_series, count=len(conflicts)))

                st.markdown(f"**{t('batch_col_preview', name=new_series, count=len(batch_col_charts))}**")

                if st.button(t("batch_col_button"), type="primary", use_container_width=True):
                    with st.spinner(t("batch_col_spinner")):
                        try:
                            updates = []
                            for chart_info in batch_col_charts:
                                df = get_chart_df(chart_info)
                                df[new_series] = None
                                st.session_state.edited_data[chart_info.key] = df

                                # Inherit format from last existing series
                                updated_formats = dict(chart_info.series_formats)
                                last_format = list(updated_formats.values())[-1] if updated_formats else "General"
                                updated_formats[new_series] = last_format

                                # New column visible by default
                                updated_visibility = dict(chart_info.series_visibility)
                                updated_visibility[new_series] = True

                                updates.append((
                                    chart_info.slide_index,
                                    chart_info.shape_name,
                                    df,
                                    chart_info.is_xy,
                                    updated_formats,
                                    updated_visibility,
                                    chart_info.shape_id,
                                ))

                            updated_bytes = update_multiple_charts(
                                st.session_state.pptx_bytes, updates,
                            )
                            st.success(t("batch_col_success", name=new_series, count=len(batch_col_charts)))
                            _commit_update(updated_bytes)
                        except Exception as e:
                            st.error(t("error_generic", e=e))

