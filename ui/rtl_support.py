"""Internationalization (Hebrew/English) and RTL CSS support."""

import streamlit as st
from pptx.enum.chart import XL_CHART_TYPE

# ---------------------------------------------------------------------------
# Translations
# ---------------------------------------------------------------------------

TRANSLATIONS = {
    "he": {
        # Page
        "page_title": "עורך גרפים - PowerPoint",
        "upload_label": "העלה קובץ מצגת (.pptx)",
        "upload_help": "גרור קובץ לכאן או לחץ לבחירה",
        "no_charts": "לא נמצאו גרפים במצגת זו",
        "select_chart": "בחר גרף לעריכה",
        "slide_preview": "תצוגת שקף",
        "data_editor": "עריכת נתונים",
        "update_preview": "עדכן תצוגה",
        "download": "הורד מצגת מעודכנת",
        "rendering": "מרנדר שקף...",
        "chart_type": "סוג גרף",
        "slide_num": "שקף",
        "error_render": "שגיאה ברינדור השקף",
        "changes_saved": "השינויים נשמרו בהצלחה",
        "editing_info": "ערוך את הנתונים בטבלה ולחץ 'עדכן תצוגה' לראות את השינויים",
        # Sidebar
        "slides": "שקפים",
        "slide_n_charts": "שקף {n} ({count} גרפים)",
        "select_slide_n": "בחר שקף {n}",
        "no_charts_in_slide": "אין גרפים בשקף הנבחר",
        # Tabs
        "tab_edit": "עריכת גרף",
        "tab_batch": "הוספת שורה לכל הגרפים",
        "tab_csv": "ייבוא/ייצוא CSV",
        # Edit tab
        "comparison_toggle": "השוואה לפני/אחרי",
        "before": "לפני",
        "after": "אחרי",
        "pct_columns_info": "עמודות באחוזים: {cols} (הזן 67 עבור 67%)",
        # Batch tab
        "batch_caption": "הוסף קטגוריה חדשה (למשל תאריך סקר) לכל הגרפים במצגת בבת אחת",
        "new_category_label": "שם הקטגוריה החדשה (למשל: 19.3.26)",
        "batch_preview": "ייתווסף שורה חדשה '{name}' ל-{count} גרפים",
        "batch_button": "הוסף שורה לכל הגרפים",
        "batch_spinner": "מוסיף שורה ומעדכן...",
        "batch_success": "שורה '{name}' נוספה ל-{count} גרפים",
        "error_generic": "שגיאה: {e}",
        # CSV tab
        "export_title": "ייצוא נתוני הגרף הנוכחי",
        "chart_info": "גרף: {name} (שקף {slide})",
        "export_button": "ייצא ל-CSV",
        "import_title": "ייבוא CSV לגרף הנוכחי",
        "import_info": "הקובץ חייב להתאים למבנה הגרף: עמודה ראשונה = קטגוריות, שאר העמודות = סדרות",
        "import_upload_label": "בחר קובץ CSV",
        "column_mismatch": "מספר העמודות לא תואם: צפוי {expected}, נמצא {found}",
        "preview_heading": "תצוגה מקדימה:",
        "apply_imported": "החל נתונים מיובאים",
        "import_success": "הנתונים יובאו בהצלחה",
        "file_read_error": "שגיאה בקריאת הקובץ: {e}",
        # Data extractor
        "series_n": "סדרה {n}",
        "category": "קטגוריה",
        # Chart type names
        "chart_column_clustered": "עמודות מקובצות",
        "chart_column_stacked": "עמודות מוערמות",
        "chart_column_stacked_100": "עמודות מוערמות 100%",
        "chart_bar_clustered": "מוטות מקובצות",
        "chart_bar_stacked": "מוטות מוערמות",
        "chart_bar_stacked_100": "מוטות מוערמות 100%",
        "chart_line": "קו",
        "chart_line_markers": "קו עם סמנים",
        "chart_line_stacked": "קו מוערם",
        "chart_pie": "עוגה",
        "chart_pie_exploded": "עוגה מפוצלת",
        "chart_doughnut": "סופגנייה",
        "chart_area": "שטח",
        "chart_area_stacked": "שטח מוערם",
        "chart_scatter": "פיזור",
        "chart_generic": "גרף",
    },
    "en": {
        # Page
        "page_title": "Chart Editor - PowerPoint",
        "upload_label": "Upload a presentation file (.pptx)",
        "upload_help": "Drag a file here or click to browse",
        "no_charts": "No charts found in this presentation",
        "select_chart": "Select a chart to edit",
        "slide_preview": "Slide Preview",
        "data_editor": "Edit Data",
        "update_preview": "Update Preview",
        "download": "Download Updated Presentation",
        "rendering": "Rendering slide...",
        "chart_type": "Chart Type",
        "slide_num": "Slide",
        "error_render": "Error rendering slide",
        "changes_saved": "Changes saved successfully",
        "editing_info": "Edit the data in the table and click 'Update Preview' to see the changes",
        # Sidebar
        "slides": "Slides",
        "slide_n_charts": "Slide {n} ({count} charts)",
        "select_slide_n": "Select Slide {n}",
        "no_charts_in_slide": "No charts in the selected slide",
        # Tabs
        "tab_edit": "Edit Chart",
        "tab_batch": "Add Row to All Charts",
        "tab_csv": "CSV Import/Export",
        # Edit tab
        "comparison_toggle": "Before/After Comparison",
        "before": "Before",
        "after": "After",
        "pct_columns_info": "Percentage columns: {cols} (enter 67 for 67%)",
        # Batch tab
        "batch_caption": "Add a new category (e.g. survey date) to all charts in the presentation at once",
        "new_category_label": "New category name (e.g.: 19.3.26)",
        "batch_preview": "A new row '{name}' will be added to {count} charts",
        "batch_button": "Add Row to All Charts",
        "batch_spinner": "Adding row and updating...",
        "batch_success": "Row '{name}' added to {count} charts",
        "error_generic": "Error: {e}",
        # CSV tab
        "export_title": "Export Current Chart Data",
        "chart_info": "Chart: {name} (Slide {slide})",
        "export_button": "Export to CSV",
        "import_title": "Import CSV to Current Chart",
        "import_info": "File must match chart structure: first column = categories, remaining columns = series",
        "import_upload_label": "Choose a CSV file",
        "column_mismatch": "Column count mismatch: expected {expected}, found {found}",
        "preview_heading": "Preview:",
        "apply_imported": "Apply Imported Data",
        "import_success": "Data imported successfully",
        "file_read_error": "Error reading file: {e}",
        # Data extractor
        "series_n": "Series {n}",
        "category": "Category",
        # Chart type names
        "chart_column_clustered": "Clustered Columns",
        "chart_column_stacked": "Stacked Columns",
        "chart_column_stacked_100": "100% Stacked Columns",
        "chart_bar_clustered": "Clustered Bars",
        "chart_bar_stacked": "Stacked Bars",
        "chart_bar_stacked_100": "100% Stacked Bars",
        "chart_line": "Line",
        "chart_line_markers": "Line with Markers",
        "chart_line_stacked": "Stacked Line",
        "chart_pie": "Pie",
        "chart_pie_exploded": "Exploded Pie",
        "chart_doughnut": "Doughnut",
        "chart_area": "Area",
        "chart_area_stacked": "Stacked Area",
        "chart_scatter": "Scatter",
        "chart_generic": "Chart",
    },
}

# Map XL_CHART_TYPE enum values to translation keys
CHART_TYPE_KEYS = {
    XL_CHART_TYPE.COLUMN_CLUSTERED: "chart_column_clustered",
    XL_CHART_TYPE.COLUMN_STACKED: "chart_column_stacked",
    XL_CHART_TYPE.COLUMN_STACKED_100: "chart_column_stacked_100",
    XL_CHART_TYPE.BAR_CLUSTERED: "chart_bar_clustered",
    XL_CHART_TYPE.BAR_STACKED: "chart_bar_stacked",
    XL_CHART_TYPE.BAR_STACKED_100: "chart_bar_stacked_100",
    XL_CHART_TYPE.LINE: "chart_line",
    XL_CHART_TYPE.LINE_MARKERS: "chart_line_markers",
    XL_CHART_TYPE.LINE_STACKED: "chart_line_stacked",
    XL_CHART_TYPE.PIE: "chart_pie",
    XL_CHART_TYPE.PIE_EXPLODED: "chart_pie_exploded",
    XL_CHART_TYPE.DOUGHNUT: "chart_doughnut",
    XL_CHART_TYPE.AREA: "chart_area",
    XL_CHART_TYPE.AREA_STACKED: "chart_area_stacked",
    XL_CHART_TYPE.XY_SCATTER: "chart_scatter",
}


def get_lang() -> str:
    """Get the current language from session state (default: Hebrew)."""
    return st.session_state.get("lang", "he")


def t(key: str, **kwargs) -> str:
    """Translate a string key to the current language.

    Supports format placeholders: t("slide_n_charts", n=1, count=3)
    """
    lang = get_lang()
    text = TRANSLATIONS[lang].get(key, TRANSLATIONS["he"].get(key, key))
    if kwargs:
        text = text.format(**kwargs)
    return text


def chart_type_display_name(chart_type: int) -> str:
    """Get a localized display name for a chart type enum value."""
    key = CHART_TYPE_KEYS.get(chart_type, "chart_generic")
    return t(key)


# Keep STRINGS as a property-like accessor for backward compat if needed
def get_strings() -> dict:
    """Get the full string dictionary for the current language."""
    return TRANSLATIONS[get_lang()]


def inject_rtl_css():
    """Inject RTL CSS when language is Hebrew, or LTR reset for English."""
    if get_lang() != "he":
        # LTR mode - minimal CSS, just slide image styling
        st.markdown(
            """
            <style>
            .slide-preview img {
                border: 2px solid #ddd;
                border-radius: 4px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            }
            </style>
            """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        """
        <style>
        /* RTL for main content */
        .stApp {
            direction: rtl;
        }

        /* Fix selectbox and other inputs */
        .stSelectbox > div,
        .stFileUploader > div,
        .stButton > button {
            direction: rtl;
            text-align: right;
        }

        /* Headers */
        h1, h2, h3, h4, h5, h6 {
            direction: rtl;
            text-align: right;
        }

        /* Paragraphs and text */
        p, span, label, .stMarkdown {
            direction: rtl;
            text-align: right;
        }

        /* Data editor - keep LTR for numbers */
        .stDataFrame td {
            direction: ltr;
            text-align: center;
        }

        /* Data editor headers RTL */
        .stDataFrame th {
            direction: rtl;
            text-align: center;
        }

        /* Success/info messages */
        .stAlert {
            direction: rtl;
            text-align: right;
        }

        /* Download button */
        .stDownloadButton > button {
            direction: rtl;
            width: 100%;
        }

        /* Slide image styling */
        .slide-preview img {
            border: 2px solid #ddd;
            border-radius: 4px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }

        /* Sidebar RTL */
        [data-testid="stSidebar"] {
            direction: rtl;
        }
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span {
            direction: rtl;
            text-align: right;
        }

        /* Tabs RTL */
        .stTabs [data-baseweb="tab-list"] {
            direction: rtl;
        }
        .stTabs [data-baseweb="tab"] {
            direction: rtl;
        }

        /* Text input RTL */
        .stTextInput > div {
            direction: rtl;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
