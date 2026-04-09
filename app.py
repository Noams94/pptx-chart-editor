"""PPTX Chart Editor - Streamlit App

Split-screen tool for editing PowerPoint chart data with live slide preview.
Left panel: actual slide rendering (via LibreOffice)
Right panel: editable data table
"""

import streamlit as st

from core.data_extractor import extract_all_charts
from core.data_writer import update_chart_data
from core.slide_renderer import render_slides
from ui.rtl_support import STRINGS, inject_rtl_css

# Page config
st.set_page_config(
    page_title=STRINGS["page_title"],
    layout="wide",
    initial_sidebar_state="collapsed",
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

# Store original bytes in session state
file_bytes = uploaded_file.getvalue()
if "pptx_bytes" not in st.session_state or st.session_state.get("file_name") != uploaded_file.name:
    st.session_state.pptx_bytes = file_bytes
    st.session_state.file_name = uploaded_file.name
    st.session_state.slide_images = None
    st.session_state.edited_data = {}

# --- Extract Charts ---
charts = extract_all_charts(st.session_state.pptx_bytes)

if not charts:
    st.warning(STRINGS["no_charts"])
    st.stop()

# --- Chart Selector ---
chart_options = {
    f"{STRINGS['slide_num']} {c.slide_index + 1} - {c.shape_name} ({c.chart_type_name})": i
    for i, c in enumerate(charts)
}

selected_label = st.selectbox(STRINGS["select_chart"], options=list(chart_options.keys()))
selected_idx = chart_options[selected_label]
selected_chart = charts[selected_idx]

# --- Render Slides (cached) ---
if st.session_state.slide_images is None:
    with st.spinner(STRINGS["rendering"]):
        try:
            st.session_state.slide_images = render_slides(st.session_state.pptx_bytes)
        except RuntimeError as e:
            st.error(str(e))
            st.stop()

# --- Split Screen Layout ---
col_preview, col_editor = st.columns([1, 1], gap="large")

# Left: Slide Preview
with col_preview:
    st.subheader(STRINGS["slide_preview"])
    slide_idx = selected_chart.slide_index
    if slide_idx < len(st.session_state.slide_images):
        st.image(
            st.session_state.slide_images[slide_idx],
            use_container_width=True,
            caption=f"{STRINGS['slide_num']} {slide_idx + 1}",
        )
    else:
        st.warning(STRINGS["error_render"])

# Right: Data Editor
with col_editor:
    st.subheader(STRINGS["data_editor"])
    st.caption(f"{STRINGS['chart_type']}: {selected_chart.chart_type_name}")
    st.caption(STRINGS["editing_info"])

    # Use a unique key per chart for the data editor
    editor_key = f"editor_{selected_chart.slide_index}_{selected_chart.shape_name}"

    # Get current data (edited or original)
    chart_key = (selected_chart.slide_index, selected_chart.shape_name)
    if chart_key in st.session_state.edited_data:
        current_df = st.session_state.edited_data[chart_key]
    else:
        current_df = selected_chart.dataframe.copy()

    # Data editor
    edited_df = st.data_editor(
        current_df,
        num_rows="dynamic",
        use_container_width=True,
        key=editor_key,
    )

    # Store edited data
    st.session_state.edited_data[chart_key] = edited_df

    # Update Preview button
    if st.button(STRINGS["update_preview"], type="primary", use_container_width=True):
        with st.spinner(STRINGS["rendering"]):
            try:
                # Apply edits to PPTX
                updated_bytes = update_chart_data(
                    st.session_state.pptx_bytes,
                    selected_chart.slide_index,
                    selected_chart.shape_name,
                    edited_df,
                    selected_chart.is_xy,
                )
                st.session_state.pptx_bytes = updated_bytes

                # Re-render slides
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
