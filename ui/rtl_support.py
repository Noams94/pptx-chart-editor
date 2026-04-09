"""Hebrew UI strings and RTL CSS support."""

import streamlit as st

# Hebrew UI strings
STRINGS = {
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
    "error_no_libreoffice": "LibreOffice לא מותקן. התקן עם: brew install --cask libreoffice",
    "changes_saved": "השינויים נשמרו בהצלחה",
    "editing_info": "ערוך את הנתונים בטבלה ולחץ 'עדכן תצוגה' לראות את השינויים",
}


def inject_rtl_css():
    """Inject RTL CSS for Hebrew UI."""
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
        </style>
        """,
        unsafe_allow_html=True,
    )
