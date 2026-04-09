"""PPTX Chart Editor - Streamlit App

Split-screen tool for editing PowerPoint chart data with live slide preview.
Features: thumbnail navigation, before/after comparison, CSV import/export,
batch row addition across all charts.
"""

import io

import pandas as pd
import streamlit as st

from core.data_extractor import extract_all_charts, _is_percentage_format
from core.data_writer import update_chart_data
from core.slide_renderer import render_slides
from ui.rtl_support import STRINGS, inject_rtl_css

# Page config
st.set_page_config(
    page_title=STRINGS["page_title"],
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_rtl_css()

st.title(STRINGS["page_title"])

# --- File Upload ---
uploaded_file = st.file_uploader(
    STRINGS["upload_label"],
    type=["pptx"],
    help=STRINGS["upload_help"],
)

if uploaded_file is None:
    st.info(STRINGS["upload_label"])
    st.stop()

# --- Initialize Session State ---
file_bytes = uploaded_file.getvalue()
if "pptx_bytes" not in st.session_state or st.session_state.get("file_name") != uploaded_file.name:
    st.session_state.pptx_bytes = file_bytes
    st.session_state.original_bytes = file_bytes  # Keep original for comparison
    st.session_state.file_name = uploaded_file.name
    st.session_state.slide_images = None
    st.session_state.original_slide_images = None
    st.session_state.edited_data = {}
    st.session_state.selected_slide = None
    st.session_state.show_comparison = False

# --- Extract Charts ---
charts = extract_all_charts(st.session_state.pptx_bytes)

if not charts:
    st.warning(STRINGS["no_charts"])
    st.stop()

# --- Render Slides (cached) ---
if st.session_state.slide_images is None:
    with st.spinner(STRINGS["rendering"]):
        try:
            st.session_state.slide_images = render_slides(st.session_state.pptx_bytes)
            st.session_state.original_slide_images = list(st.session_state.slide_images)
        except RuntimeError as e:
            st.error(str(e))
            st.stop()

# --- Sidebar: Slide Thumbnails ---
with st.sidebar:
    st.subheader("שקפים")

    # Get unique slide indices that have charts
    slide_indices = sorted(set(c.slide_index for c in charts))

    for slide_idx in slide_indices:
        slide_charts = [c for c in charts if c.slide_index == slide_idx]
        chart_count = len(slide_charts)

        if slide_idx < len(st.session_state.slide_images):
            # Show thumbnail
            is_selected = st.session_state.selected_slide == slide_idx
            label = f"שקף {slide_idx + 1} ({chart_count} גרפים)"

            if is_selected:
                st.markdown(f"**► {label}**")
            else:
                st.caption(label)

            if st.sidebar.button(
                f"בחר שקף {slide_idx + 1}",
                key=f"thumb_{slide_idx}",
                use_container_width=True,
            ):
                st.session_state.selected_slide = slide_idx
                st.rerun()

            st.image(
                st.session_state.slide_images[slide_idx],
                use_container_width=True,
            )
            st.divider()

# --- Chart Selector (filtered by selected slide) ---
if st.session_state.selected_slide is not None:
    filtered_charts = [c for c in charts if c.slide_index == st.session_state.selected_slide]
else:
    filtered_charts = charts

chart_options = {
    f"{STRINGS['slide_num']} {c.slide_index + 1} - {c.shape_name} ({c.chart_type_name})": i
    for i, c in enumerate(charts)
    if c in filtered_charts
}

if not chart_options:
    st.info("אין גרפים בשקף הנבחר")
    st.stop()

selected_label = st.selectbox(STRINGS["select_chart"], options=list(chart_options.keys()))
selected_idx = chart_options[selected_label]
selected_chart = charts[selected_idx]

# --- Tabs: Edit / Batch Add ---
tab_edit, tab_batch, tab_csv = st.tabs(["עריכת גרף", "הוספת שורה לכל הגרפים", "ייבוא/ייצוא CSV"])

# ==================== TAB 1: EDIT ====================
with tab_edit:
    # --- Comparison toggle ---
    col_toggle, _ = st.columns([1, 3])
    with col_toggle:
        show_comparison = st.checkbox("השוואה לפני/אחרי", value=st.session_state.show_comparison)
        st.session_state.show_comparison = show_comparison

    # --- Split Screen Layout ---
    if show_comparison:
        col_before, col_after, col_editor = st.columns([1, 1, 1], gap="medium")
    else:
        col_after, col_editor = st.columns([1, 1], gap="large")
        col_before = None

    # Before (original)
    if col_before is not None:
        with col_before:
            st.subheader("לפני")
            slide_idx = selected_chart.slide_index
            if st.session_state.original_slide_images and slide_idx < len(st.session_state.original_slide_images):
                st.image(
                    st.session_state.original_slide_images[slide_idx],
                    use_container_width=True,
                )

    # After (current)
    with col_after:
        st.subheader("אחרי" if show_comparison else STRINGS["slide_preview"])
        slide_idx = selected_chart.slide_index
        if slide_idx < len(st.session_state.slide_images):
            st.image(
                st.session_state.slide_images[slide_idx],
                use_container_width=True,
                caption=f"{STRINGS['slide_num']} {slide_idx + 1}",
            )
        else:
            st.warning(STRINGS["error_render"])

    # Data Editor
    with col_editor:
        st.subheader(STRINGS["data_editor"])
        st.caption(f"{STRINGS['chart_type']}: {selected_chart.chart_type_name}")

        pct_cols = [
            col for col in selected_chart.dataframe.columns[1:]
            if _is_percentage_format(selected_chart.series_formats.get(col, ""))
        ]
        if pct_cols:
            st.caption(f"עמודות באחוזים: {', '.join(pct_cols)} (הזן 67 עבור 67%)")
        st.caption(STRINGS["editing_info"])

        editor_key = f"editor_{selected_chart.slide_index}_{selected_chart.shape_name}"
        chart_key = (selected_chart.slide_index, selected_chart.shape_name)

        if chart_key in st.session_state.edited_data:
            current_df = st.session_state.edited_data[chart_key]
        else:
            current_df = selected_chart.dataframe.copy()

        edited_df = st.data_editor(
            current_df,
            num_rows="dynamic",
            use_container_width=True,
            key=editor_key,
        )

        st.session_state.edited_data[chart_key] = edited_df

        # Update Preview button
        if st.button(STRINGS["update_preview"], type="primary", use_container_width=True):
            with st.spinner(STRINGS["rendering"]):
                try:
                    updated_bytes = update_chart_data(
                        st.session_state.pptx_bytes,
                        selected_chart.slide_index,
                        selected_chart.shape_name,
                        edited_df,
                        selected_chart.is_xy,
                        selected_chart.series_formats,
                    )
                    st.session_state.pptx_bytes = updated_bytes
                    st.session_state.slide_images = render_slides(updated_bytes)
                    st.success(STRINGS["changes_saved"])
                    st.rerun()
                except Exception as e:
                    st.error(f"{STRINGS['error_render']}: {e}")

        # Download button
        st.download_button(
            label=STRINGS["download"],
            data=st.session_state.pptx_bytes,
            file_name=f"updated_{st.session_state.file_name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )


# ==================== TAB 2: BATCH ADD ROW ====================
with tab_batch:
    st.subheader("הוספת שורה לכל הגרפים")
    st.caption("הוסף קטגוריה חדשה (למשל תאריך סקר) לכל הגרפים במצגת בבת אחת")

    new_category = st.text_input("שם הקטגוריה החדשה (למשל: 19.3.26)", key="batch_category")

    if new_category:
        st.markdown(f"**ייתווסף שורה חדשה '{new_category}' ל-{len(charts)} גרפים**")

        if st.button("הוסף שורה לכל הגרפים", type="primary", use_container_width=True):
            with st.spinner("מוסיף שורה ומעדכן..."):
                try:
                    current_bytes = st.session_state.pptx_bytes

                    for chart_info in charts:
                        chart_key = (chart_info.slide_index, chart_info.shape_name)

                        # Get current data (edited or original)
                        if chart_key in st.session_state.edited_data:
                            df = st.session_state.edited_data[chart_key].copy()
                        else:
                            df = chart_info.dataframe.copy()

                        # Add new row
                        new_row = {df.columns[0]: new_category}
                        for col in df.columns[1:]:
                            new_row[col] = None
                        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

                        # Store edited data
                        st.session_state.edited_data[chart_key] = df

                        # Update PPTX
                        current_bytes = update_chart_data(
                            current_bytes,
                            chart_info.slide_index,
                            chart_info.shape_name,
                            df,
                            chart_info.is_xy,
                            chart_info.series_formats,
                        )

                    st.session_state.pptx_bytes = current_bytes
                    st.session_state.slide_images = render_slides(current_bytes)
                    st.success(f"שורה '{new_category}' נוספה ל-{len(charts)} גרפים")
                    st.rerun()
                except Exception as e:
                    st.error(f"שגיאה: {e}")


# ==================== TAB 3: CSV IMPORT/EXPORT ====================
with tab_csv:
    st.subheader("ייבוא/ייצוא CSV")

    col_export, col_import = st.columns(2, gap="large")

    with col_export:
        st.markdown("**ייצוא נתוני הגרף הנוכחי**")
        st.caption(f"גרף: {selected_chart.shape_name} (שקף {selected_chart.slide_index + 1})")

        chart_key = (selected_chart.slide_index, selected_chart.shape_name)
        if chart_key in st.session_state.edited_data:
            export_df = st.session_state.edited_data[chart_key]
        else:
            export_df = selected_chart.dataframe.copy()

        # Export to CSV with BOM for Excel Hebrew support
        csv_buffer = io.StringIO()
        export_df.to_csv(csv_buffer, index=False, encoding="utf-8")
        csv_bytes = ("\ufeff" + csv_buffer.getvalue()).encode("utf-8")

        st.download_button(
            label="ייצא ל-CSV",
            data=csv_bytes,
            file_name=f"{selected_chart.shape_name}_slide{selected_chart.slide_index + 1}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col_import:
        st.markdown("**ייבוא CSV לגרף הנוכחי**")
        st.caption("הקובץ חייב להתאים למבנה הגרף: עמודה ראשונה = קטגוריות, שאר העמודות = סדרות")

        csv_file = st.file_uploader(
            "בחר קובץ CSV",
            type=["csv"],
            key=f"csv_import_{selected_chart.slide_index}_{selected_chart.shape_name}",
        )

        if csv_file is not None:
            try:
                imported_df = pd.read_csv(csv_file, encoding="utf-8-sig")

                # Validate column count
                expected_cols = len(selected_chart.dataframe.columns)
                if len(imported_df.columns) != expected_cols:
                    st.error(f"מספר העמודות לא תואם: צפוי {expected_cols}, נמצא {len(imported_df.columns)}")
                else:
                    # Rename columns to match original
                    imported_df.columns = selected_chart.dataframe.columns

                    st.markdown("**תצוגה מקדימה:**")
                    st.dataframe(imported_df, use_container_width=True)

                    if st.button("החל נתונים מיובאים", type="primary", use_container_width=True):
                        chart_key = (selected_chart.slide_index, selected_chart.shape_name)
                        st.session_state.edited_data[chart_key] = imported_df

                        with st.spinner(STRINGS["rendering"]):
                            updated_bytes = update_chart_data(
                                st.session_state.pptx_bytes,
                                selected_chart.slide_index,
                                selected_chart.shape_name,
                                imported_df,
                                selected_chart.is_xy,
                                selected_chart.series_formats,
                            )
                            st.session_state.pptx_bytes = updated_bytes
                            st.session_state.slide_images = render_slides(updated_bytes)
                            st.success("הנתונים יובאו בהצלחה")
                            st.rerun()
            except Exception as e:
                st.error(f"שגיאה בקריאת הקובץ: {e}")
