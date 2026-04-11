"""PPTX Chart Editor - Streamlit App

Split-screen tool for editing PowerPoint chart data with live slide preview.
Features: thumbnail navigation, before/after comparison, CSV import/export,
batch row addition across all charts, Hebrew/English language switching.
"""

import base64
from collections import defaultdict
import io
import re

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from core.data_extractor import extract_all_charts, is_percentage_format
from core.data_writer import update_chart_data, update_multiple_charts
from core.slide_renderer import render_slides
from ui.rtl_support import t, inject_rtl_css

# --- Language Selector (must be before page config uses translated title) ---
if "lang" not in st.session_state:
    st.session_state.lang = "en"

# Page config
st.set_page_config(
    page_title=t("page_title"),
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_rtl_css()

# --- Footer (injected as fixed CSS so it shows regardless of st.stop) ---
st.markdown(
    """
    <style>
    .fixed-footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background: white;
        border-top: 1px solid #eee;
        padding: 6px 0;
        text-align: center;
        color: #888;
        font-size: 0.8rem;
        z-index: 999;
        direction: ltr;
    }
    .fixed-footer a { color: #888; text-decoration: none; }
    .fixed-footer a:hover { color: #4A90D9; text-decoration: underline; }
    .stApp > .main { padding-bottom: 40px; }
    </style>
    <div class="fixed-footer">
        &copy; All Rights Reserved &middot; Dr. Noam Keshet &middot;
        <a href="https://noamkeshet.com" target="_blank">noamkeshet.com</a> &middot;
        <a href="mailto:keshet.noam@gmail.com">keshet.noam@gmail.com</a>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- Language toggle (top of page) ---
col_title, col_lang = st.columns([4, 1])
with col_title:
    st.title(t("page_title"))
with col_lang:
    lang_options = {"עברית": "he", "English": "en"}
    current_label = "עברית" if st.session_state.lang == "he" else "English"
    selected_lang_label = st.selectbox(
        "🌐",
        options=list(lang_options.keys()),
        index=list(lang_options.keys()).index(current_label),
        label_visibility="collapsed",
    )
    new_lang = lang_options[selected_lang_label]
    if new_lang != st.session_state.lang:
        st.session_state.lang = new_lang
        st.session_state.charts_cache = None  # Re-extract with new language
        st.rerun()

st.caption(t("instructions"))


def get_chart_df(chart_info):
    """Get current DataFrame for a chart (edited version if exists, otherwise original)."""
    key = (chart_info.slide_index, chart_info.shape_name)
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
        name_map[(ci.slide_index, ci.shape_name)] = sheet
    return name_map


def _commit_update(updated_bytes: bytes):
    """Save updated PPTX, re-render, trigger auto-download, and rerun."""
    _apply_and_rerender(updated_bytes)
    _schedule_auto_download()
    st.rerun()


# --- File Upload ---
uploaded_file = st.file_uploader(
    t("upload_label"),
    type=["pptx"],
    help=t("upload_help"),
)

if uploaded_file is None:
    st.info(t("upload_label"))
    st.stop()

# --- Initialize Session State ---
if "pptx_bytes" not in st.session_state or st.session_state.get("file_name") != uploaded_file.name:
    st.session_state.pptx_bytes = uploaded_file.getvalue()
    st.session_state.file_name = uploaded_file.name
    st.session_state.slide_images = None
    st.session_state.original_slide_images = None
    st.session_state.edited_data = {}
    st.session_state.selected_slide = None
    st.session_state.show_comparison = False
    st.session_state.charts_cache = None
    st.session_state.series_visibility = {}

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
    st.toast(t("auto_saved_msg"))

# --- Extract Charts (cached in session state) ---
if st.session_state.get("charts_cache") is None:
    st.session_state.charts_cache = extract_all_charts(st.session_state.pptx_bytes)

charts = st.session_state.charts_cache

if not charts:
    st.warning(t("no_charts"))
    st.stop()

# --- Render Slides (lazy — triggered by button or on first load) ---
if st.session_state.slide_images is None:
    render_col, _ = st.columns([1, 3])
    with render_col:
        if st.button(t("render_preview_btn"), type="primary", use_container_width=True):
            with st.spinner(t("rendering")):
                try:
                    st.session_state.slide_images = render_slides(st.session_state.pptx_bytes)
                    st.session_state.original_slide_images = list(st.session_state.slide_images)
                    st.rerun()
                except RuntimeError as e:
                    st.error(str(e))
    st.info(t("render_hint"))

# --- Group charts by slide (computed once) ---
charts_by_slide = defaultdict(list)
for c in charts:
    charts_by_slide[c.slide_index].append(c)

# --- Sidebar: Slide Thumbnails ---
slide_images = st.session_state.slide_images or []

with st.sidebar:
    # Auto-save toggle
    st.session_state.setdefault("auto_save", True)
    auto_save = st.checkbox(
        t("auto_save_label"),
        value=st.session_state.auto_save,
        help=t("auto_save_info"),
    )
    st.session_state.auto_save = auto_save
    st.divider()

    st.subheader(t("slides"))

    for slide_idx in sorted(charts_by_slide):
        chart_count = len(charts_by_slide[slide_idx])
        is_selected = st.session_state.selected_slide == slide_idx
        label = t("slide_n_charts", n=slide_idx + 1, count=chart_count)

        if is_selected:
            st.markdown(f"**► {label}**")
        else:
            st.caption(label)

        if st.sidebar.button(
            t("select_slide_n", n=slide_idx + 1),
            key=f"thumb_{slide_idx}",
            use_container_width=True,
        ):
            st.session_state.selected_slide = slide_idx
            st.rerun()

        if slide_idx < len(slide_images):
            st.image(slide_images[slide_idx], use_container_width=True)
        st.divider()

# --- Chart Selector (filtered by selected slide) ---
if st.session_state.selected_slide is not None:
    filtered_indices = {i for i, c in enumerate(charts) if c.slide_index == st.session_state.selected_slide}
else:
    filtered_indices = set(range(len(charts)))

chart_options = {
    f"{t('slide_num')} {c.slide_index + 1} - {c.shape_name} ({c.chart_type_name})": i
    for i, c in enumerate(charts)
    if i in filtered_indices
}

if not chart_options:
    st.info(t("no_charts_in_slide"))
    st.stop()

# ==================== GLOBAL TABS (all charts) ====================
tab_excel, tab_batch = st.tabs([t("tab_excel"), t("tab_batch")])

# --- EXCEL IMPORT/EXPORT (ALL CHARTS) ---
with tab_excel:
    st.subheader(t("tab_excel"))

    col_export_xl, col_import_xl = st.columns(2, gap="large")

    # --- Export ---
    with col_export_xl:
        st.markdown(f"**{t('excel_export_title')}**")
        st.caption(t("excel_export_caption", count=len(charts)))

        sheet_name_map = _build_sheet_name_map(charts)

        # Cache xlsx bytes — only rebuild when chart data changes
        edit_fingerprint = str(sorted(st.session_state.edited_data.keys()))
        if (st.session_state.get("xl_export_fingerprint") != edit_fingerprint
                or "xl_export_bytes" not in st.session_state):
            xl_buffer = io.BytesIO()
            with pd.ExcelWriter(xl_buffer, engine="openpyxl") as writer:
                for chart_info in charts:
                    sheet = sheet_name_map[(chart_info.slide_index, chart_info.shape_name)]
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
            # Cache import parsing — only re-parse when a new file is uploaded
            xl_cache_key = (xl_file.name, xl_file.size)
            if st.session_state.get("xl_import_cache_key") != xl_cache_key:
                xls = pd.ExcelFile(xl_file, engine="openpyxl")

                # Build reverse lookup: sheet name -> chart_info (using same dedup as export)
                sheet_to_chart = {v: k for k, v in sheet_name_map.items()}
                charts_by_key = {(ci.slide_index, ci.shape_name): ci for ci in charts}

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
                            chart_key = (ci.slide_index, ci.shape_name)
                            st.session_state.edited_data[chart_key] = df
                            updates.append((
                                ci.slide_index,
                                ci.shape_name,
                                df,
                                ci.is_xy,
                                ci.series_formats,
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
with tab_batch:
    st.subheader(t("tab_batch"))
    st.caption(t("batch_caption"))

    new_category = st.text_input(t("new_category_label"), key="batch_category")

    if new_category:
        st.markdown(f"**{t('batch_preview', name=new_category, count=len(charts))}**")

        if st.button(t("batch_button"), type="primary", use_container_width=True):
            with st.spinner(t("batch_spinner")):
                try:
                    updates = []
                    for chart_info in charts:
                        df = get_chart_df(chart_info)

                        new_row = {df.columns[0]: new_category}
                        for col in df.columns[1:]:
                            new_row[col] = None
                        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

                        chart_key = (chart_info.slide_index, chart_info.shape_name)
                        st.session_state.edited_data[chart_key] = df
                        updates.append((
                            chart_info.slide_index,
                            chart_info.shape_name,
                            df,
                            chart_info.is_xy,
                            chart_info.series_formats,
                        ))

                    updated_bytes = update_multiple_charts(
                        st.session_state.pptx_bytes, updates,
                    )
                    st.success(t("batch_success", name=new_category, count=len(charts)))
                    _commit_update(updated_bytes)
                except Exception as e:
                    st.error(t("error_generic", e=e))


st.divider()

# ==================== PER-CHART SECTION ====================
selected_label = st.selectbox(t("select_chart"), options=list(chart_options.keys()))
selected_idx = chart_options[selected_label]
selected_chart = charts[selected_idx]
slide_idx = selected_chart.slide_index

tab_edit, tab_select_data, tab_csv = st.tabs([t("tab_edit"), t("tab_select_data"), t("tab_csv")])

# --- EDIT CHART ---
with tab_edit:
    col_toggle, _ = st.columns([1, 3])
    with col_toggle:
        show_comparison = st.checkbox(t("comparison_toggle"), value=st.session_state.show_comparison)
        st.session_state.show_comparison = show_comparison

    if show_comparison:
        col_before, col_after, col_editor = st.columns([1, 1, 1], gap="medium")
    else:
        col_after, col_editor = st.columns([1, 1], gap="large")
        col_before = None

    # Before (original)
    if col_before is not None:
        with col_before:
            st.subheader(t("before"))
            original_images = st.session_state.original_slide_images or []
            if original_images and slide_idx < len(original_images):
                st.image(original_images[slide_idx], use_container_width=True)

    # After (current)
    with col_after:
        st.subheader(t("after") if show_comparison else t("slide_preview"))
        if slide_images and slide_idx < len(slide_images):
            st.image(
                slide_images[slide_idx],
                use_container_width=True,
                caption=f"{t('slide_num')} {slide_idx + 1}",
            )
        else:
            st.info(t("render_hint"))

    # Data Editor
    with col_editor:
        st.subheader(t("data_editor"))
        st.caption(f"{t('chart_type')}: {selected_chart.chart_type_name}")

        pct_cols = [
            col for col in selected_chart.dataframe.columns[1:]
            if is_percentage_format(selected_chart.series_formats.get(col, ""))
        ]
        if pct_cols:
            st.caption(t("pct_columns_info", cols=", ".join(pct_cols)))
        st.caption(t("editing_info"))

        editor_key = f"editor_{selected_chart.slide_index}_{selected_chart.shape_name}"
        chart_key = (selected_chart.slide_index, selected_chart.shape_name)

        edited_df = st.data_editor(
            get_chart_df(selected_chart),
            num_rows="dynamic",
            use_container_width=True,
            key=editor_key,
        )

        st.session_state.edited_data[chart_key] = edited_df

        if not edited_df.equals(selected_chart.dataframe):
            st.warning(t("unsaved_warning"))

        if st.button(t("update_preview"), type="primary", use_container_width=True):
            with st.spinner(t("rendering")):
                try:
                    # Pass current visibility state if available
                    edit_vis_key = (selected_chart.slide_index, selected_chart.shape_name)
                    current_vis = st.session_state.get("series_visibility", {}).get(
                        edit_vis_key, selected_chart.series_visibility
                    )
                    updated_bytes = update_chart_data(
                        st.session_state.pptx_bytes,
                        selected_chart.slide_index,
                        selected_chart.shape_name,
                        edited_df,
                        selected_chart.is_xy,
                        selected_chart.series_formats,
                        current_vis,
                    )
                    st.success(t("changes_saved"))
                    _commit_update(updated_bytes)
                except Exception as e:
                    st.error(f"{t('error_render')}: {e}")

        st.download_button(
            label=t("download"),
            data=st.session_state.pptx_bytes,
            file_name=f"updated_{st.session_state.file_name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

# --- SELECT DATA (Series Visibility) ---
with tab_select_data:
    st.subheader(t("tab_select_data"))
    st.caption(t("select_data_caption"))

    # Initialize visibility state from chart metadata if not already set
    vis_key = (selected_chart.slide_index, selected_chart.shape_name)
    if "series_visibility" not in st.session_state:
        st.session_state.series_visibility = {}
    if vis_key not in st.session_state.series_visibility:
        st.session_state.series_visibility[vis_key] = dict(selected_chart.series_visibility)

    current_visibility = st.session_state.series_visibility[vis_key]

    # Get series names (skip category column for non-XY charts)
    if selected_chart.is_xy:
        series_display_names = selected_chart.series_names
    else:
        series_display_names = list(selected_chart.series_visibility.keys())

    # Show checkboxes per series
    st.markdown(f"**{t('series_visible_label')}**")
    updated_visibility = {}
    for name in series_display_names:
        visible = current_visibility.get(name, True)
        updated_visibility[name] = st.checkbox(
            name,
            value=visible,
            key=f"vis_{selected_chart.slide_index}_{selected_chart.shape_name}_{name}",
        )

    # Validation — at least one series must be visible
    visible_count = sum(1 for v in updated_visibility.values() if v)
    hidden_count = len(updated_visibility) - visible_count

    if hidden_count > 0:
        st.info(t("hidden_series_count", count=hidden_count))
    else:
        st.success(t("all_series_visible"))

    # Store updated visibility
    st.session_state.series_visibility[vis_key] = updated_visibility

    # Update button
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
                )
                st.success(t("visibility_updated"))
                _commit_update(updated_bytes)
            except Exception as e:
                st.error(f"{t('error_render')}: {e}")

# --- CSV IMPORT/EXPORT ---
with tab_csv:
    st.subheader(t("tab_csv"))

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
            key=f"csv_import_{selected_chart.slide_index}_{selected_chart.shape_name}",
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
                        chart_key = (selected_chart.slide_index, selected_chart.shape_name)
                        st.session_state.edited_data[chart_key] = imported_df

                        with st.spinner(t("rendering")):
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
                            )
                            st.success(t("import_success"))
                            _commit_update(updated_bytes)
            except Exception as e:
                st.error(t("file_read_error", e=e))

