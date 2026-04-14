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
        "instructions": "העלה קובץ מצגת (.pptx) כדי לצפות ולערוך נתוני גרפים. בחר גרף, שנה ערכים בטבלה, והורד את המצגת המעודכנת.",
        "upload_label": "העלה קובץ מצגת (.pptx)",
        "upload_help": "גרור קובץ לכאן או לחץ לבחירה",
        "no_charts": "לא נמצאו גרפים במצגת זו",
        "select_chart": "בחר גרף לעריכה",
        "slide_preview": "תצוגת שקף",
        "data_editor": "עריכת נתונים",
        "update_preview": "עדכן תצוגה",
        "download": "הורד מצגת מעודכנת",
        "rendering": "מרנדר שקף...",
        "render_preview_btn": "רנדר תצוגה מקדימה",
        "render_hint": "לחץ על 'רנדר תצוגה מקדימה' ליצירת תמונות השקפים. ניתן לערוך נתונים גם ללא רנדור.",
        "chart_type": "סוג גרף",
        "slide_num": "שקף",
        "error_render": "שגיאה ברינדור השקף",
        "changes_saved": "השינויים נשמרו בהצלחה",
        "editing_info": "ערוך את הנתונים בטבלה. התצוגה מתעדכנת מיידית. לחץ 'שמור במצגת' כדי להחיל על הקובץ.",
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
        "chart_comparison_toggle": "לפני/אחרי",
        "slide_comparison_toggle": "לפני/אחרי",
        "before": "לפני",
        "after": "אחרי",
        "chart_preview": "תצוגת גרף",
        "full_slide_preview": "תצוגת שקף מלאה",
        "save_to_pptx": "שמור במצגת",
        "saving_to_pptx": "שומר במצגת ומרנדר...",
        "saved_to_pptx": "השינויים נשמרו במצגת בהצלחה",
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
        # Excel tab
        "tab_excel": "ייבוא/ייצוא Excel (כל הגרפים)",
        "excel_export_title": "ייצוא כל הגרפים לקובץ Excel",
        "excel_export_caption": "הורד קובץ .xlsx אחד עם כל {count} הגרפים, כל אחד בלשונית נפרדת",
        "excel_export_button": "ייצא הכל ל-Excel",
        "excel_import_title": "ייבוא Excel לעדכון גרפים",
        "excel_import_caption": "העלה קובץ .xlsx שיוצא בעבר. הלשוניות יותאמו לגרפים לפי שם.",
        "excel_import_upload_label": "בחר קובץ Excel (.xlsx)",
        "excel_matched_charts": "הותאמו {matched} מתוך {total} גרפים",
        "excel_no_matches": "אף לשונית לא תואמת לגרף. שמות הלשוניות צריכים להיות בפורמט 'Slide{n}_שםהגרף'.",
        "excel_changes_found": "נמצאו {changed} גרפים עם שינויים מתוך {total} שהותאמו",
        "excel_unchanged": "{count} גרפים ללא שינוי",
        "excel_no_changes": "לא נמצאו שינויים — הנתונים באקסל זהים למצגת הנוכחית",
        "excel_apply_button": "החל את כל הנתונים המותאמים",
        "excel_apply_spinner": "מעדכן {count} גרפים...",
        "excel_apply_success": "עודכנו {count} גרפים בהצלחה",
        "excel_column_mismatch_warning": "לשונית '{sheet}': מספר עמודות לא תואם (צפוי {expected}, נמצא {found}) - דולג",
        "excel_sheet_no_match": "לשונית '{sheet}' - לא נמצא גרף תואם",
        # Auto-save
        "auto_save_label": "שמירה אוטומטית אחרי עדכון",
        "auto_save_info": "מוריד אוטומטית את הקובץ המעודכן אחרי לחיצה על 'שמור במצגת'",
        "unsaved_warning": "יש שינויים שלא נשמרו. לחץ 'שמור במצגת' כדי להחיל.",
        "save_disabled_hint": "ערוך נתונים בטבלה כדי להפעיל שמירה",
        # Sidebar filter
        "filter_slides": "סנן שקפים...",
        # Onboarding
        "onboarding_summary": "נמצאו {slides} שקפים עם {charts} גרפים. בחר שקף מהסרגל הצדי כדי להתחיל.",
        "app_subtitle": "עריכת נתוני גרפים ב-PowerPoint בקלות ובמהירות",
        "upload_section_title": "העלאת מצגת",
        "preview_section_title": "תצוגה מקדימה",
        "editor_section_title": "עריכת נתונים",
        "no_slide_selected": "בחר שקף מהסרגל הצדי כדי להתחיל בעריכה",
        "chart_details": "פרטי גרף",
        "actions": "פעולות",
        # Undo
        "undo": "ביטול",
        "undo_success": "השינוי האחרון בוטל",
        "no_undo": "אין שינויים לביטול",
        "auto_saved_msg": "הקובץ נשמר אוטומטית",
        # Select Data tab
        "tab_select_data": "בחירת נתונים",
        "select_data_caption": "בחר אילו סדרות יוצגו בגרף. סדרות מוסתרות שומרות על הנתונים שלהן אבל לא מוצגות בגרף.",
        "series_visible_label": "סדרות מוצגות",
        "at_least_one_series": "חייבת להיות לפחות סדרה אחת מוצגת",
        "visibility_updated": "נראות הסדרות עודכנה בהצלחה",
        "update_visibility": "עדכן נראות",
        "hidden_series_count": "{count} סדרות מוסתרות",
        "all_series_visible": "כל הסדרות מוצגות",
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
        # Progress indicator
        "step_upload": "העלאה",
        "step_select": "עריכת כל הגרפים",
        "step_edit": "עריכת שקף",
        "step_download": "הורדה",
        # Workflow guidance
        "welcome_message": "ברוכים הבאים לעורך הגרפים!",
        "getting_started": "איך מתחילים?",
        "step_1": "העלה קובץ PowerPoint (.pptx)",
        "step_2": "בחר שקף וגרף לעריכה",
        "step_3": "ערוך את הנתונים בטבלה",
        "step_4": "הורד את הקובץ המעודכן",
        # Accessibility
        "filter_help": "סנן שקפים לפי מספר או שם גרף",
        "series_visibility_help": "בחר אילו סדרות יוצגו בגרף. סדרות מוסתרות שומרות על הנתונים שלהן אבל לא מוצגות בגרף.",
        "csv_help": "ייצא את הנתונים הנוכחיים ל-CSV או ייבא נתונים מקובץ CSV",
        # Metrics
        "total_slides": "סה\"כ שקפים",
        "total_charts": "סה\"כ גרפים",
        "chart_types": "סוגי גרפים",
        "overview": "סקירה כללית",
        "overview_description": "בחר שקף מהסרגל הצדי כדי להתחיל בעריכה. כל שקף מכיל מספר גרפים שניתן לערוך בנפרד.",
        "tab_overview": "סקירה",
        "charts": "גרפים",
        "preview": "תצוגה מקדימה",
        # Errors and feedback
        "csv_imported_success": "CSV יובא בהצלחה",
        "csv_import_error": "שגיאה בייבוא CSV",
        "applying_csv": "מחיל נתונים מ-CSV...",
        "undo_not_available": "ביטול לא זמין כרגע",
        "slide_count_info": "{count} שקפים עם גרפים",
        # Getting Started wizard
        "wizard_welcome_title": "ברוכים הבאים לעורך הגרפים!",
        "wizard_welcome_desc": "כלי זה מאפשר לך לערוך נתוני גרפים ישירות בתוך מצגות PowerPoint. בצע את השלבים הבאים כדי להתחיל.",
        "wizard_upload_title": "העלאת מצגת",
        "wizard_upload_desc": "גרור קובץ PowerPoint (.pptx) לאזור ההעלאה בסרגל הצדי, או לחץ על 'Browse files' לבחירת קובץ.",
        "wizard_upload_waiting": "ממתין להעלאת קובץ...",
        "wizard_upload_done": "הקובץ '{name}' הועלה בהצלחה!",
        "wizard_select_title": "עריכת כל הגרפים",
        "wizard_select_desc": "ייצא את כל הגרפים לקובץ Excel, ערוך אותם, וייבא בחזרה. דרך מהירה לעדכן את כל הנתונים בבת אחת.",
        "wizard_edit_title": "עריכת שקף בודד",
        "wizard_edit_desc": "בחר שקף מהסרגל הצדי, ערוך גרף ספציפי בטבלה, צפה בתצוגה מקדימה, ושמור במצגת.",
        "wizard_next": "הבא",
        "wizard_back": "הקודם",
        "wizard_start_editing": "התחל לערוך",
        "wizard_upload_now": "העלה קובץ בסרגל הצדי",
        "wizard_file_uploaded": "קובץ הועלה",
        "quick_start": "התחלה מהירה",
        # User Guide
        "tab_guide": "מדריך למשתמש",
        "guide_overview_title": "סקירה כללית",
        "guide_overview_body": "עורך הגרפים מאפשר לערוך נתוני גרפים בתוך מצגות PowerPoint ישירות בדפדפן. העלו קובץ .pptx, שנו ערכים בטבלה אינטראקטיבית עם תצוגה מקדימה חיה, והורידו את המצגת המעודכנת — ללא צורך ב-PowerPoint.",
        "guide_start_title": "איך מתחילים?",
        "guide_start_1": "**העלאה** — גררו קובץ .pptx לאזור ההעלאה בסרגל הצדי, או לחצו לבחירת קובץ",
        "guide_start_2": "**בחירת שקף** — לחצו על תמונה ממוזערת של שקף בסרגל הצדי",
        "guide_start_3": "**בחירת גרף** — בחרו גרף מהתפריט הנפתח",
        "guide_start_4": "**עריכת נתונים** — שנו ערכים בטבלה; התצוגה המקדימה מתעדכנת מיידית",
        "guide_start_5": "**שמירה** — לחצו \"שמור במצגת\" ולאחר מכן הורידו את הקובץ המעודכן",
        "guide_edit_title": "עריכת גרף",
        "guide_edit_body": """- טבלת הנתונים מציגה קטגוריות (שורות) וסדרות (עמודות)
- ערכו כל תא — תצוגת הגרף מתעדכנת בזמן אמת
- השתמשו בלחצן **לפני/אחרי** להשוואה עם הגרף המקורי
- לחצו **שמור במצגת** כדי לכתוב את השינויים בחזרה לקובץ
- השתמשו ב**ביטול** כדי לחזור לשינוי האחרון""",
        "guide_excel_title": "עריכה מרוכזת עם Excel",
        "guide_excel_body": """- עברו ללשונית **ייבוא/ייצוא Excel**
- לחצו **ייצא הכל ל-Excel** — מוריד קובץ .xlsx עם לשונית לכל גרף
- ערכו את קובץ ה-Excel בכל תוכנת גיליונות
- העלו את הקובץ המעודכן — האפליקציה מתאימה לשוניות לגרפים אוטומטית
- בדקו את השינויים ולחצו **החל את כל הנתונים המותאמים**""",
        "guide_csv_title": "ייבוא/ייצוא CSV",
        "guide_csv_body": """- בלשונית **CSV**, ייצאו או ייבאו נתונים עבור הגרף הנבחר
- קובץ ה-CSV חייב להתאים למבנה הגרף: עמודה ראשונה = קטגוריות, השאר = סדרות""",
        "guide_batch_title": "הוספת שורה לכל הגרפים",
        "guide_batch_body": """- בלשונית **הוספת שורה**, הקלידו שם קטגוריה חדשה (למשל תאריך סקר)
- לחצו **הוסף שורה לכל הגרפים** כדי להוסיף אותה שורה לכל הגרפים בבת אחת""",
        "guide_visibility_title": "בחירת נתונים (נראות סדרות)",
        "guide_visibility_body": """- בלשונית **בחירת נתונים**, בחרו אילו סדרות יוצגו בגרף
- סדרות מוסתרות שומרות על הנתונים שלהן אבל לא מוצגות
- חייבת להיות לפחות סדרה אחת מוצגת""",
        "guide_tips_title": "טיפים",
        "guide_tips_body": """- **אחוזים**: בעמודות אחוזים, הקלידו 67 עבור 67% (המרה אוטומטית)
- **שמירה אוטומטית**: הפעילו \"שמירה אוטומטית אחרי עדכון\" להורדה אוטומטית אחרי כל שמירה
- **סינון שקפים**: השתמשו בתיבת החיפוש בסרגל הצדי לסינון לפי מספר שקף או שם גרף""",
    },
    "en": {
        # Page
        "page_title": "Chart Editor - PowerPoint",
        "instructions": "Upload a PowerPoint file (.pptx) to view and edit chart data. Select a chart, modify values in the table, and download the updated presentation.",
        "upload_label": "Upload a presentation file (.pptx)",
        "upload_help": "Drag a file here or click to browse",
        "no_charts": "No charts found in this presentation",
        "select_chart": "Select a chart to edit",
        "slide_preview": "Slide Preview",
        "data_editor": "Edit Data",
        "update_preview": "Update Preview",
        "download": "Download Updated Presentation",
        "rendering": "Rendering slide...",
        "render_preview_btn": "Render Preview",
        "render_hint": "Click 'Render Preview' to generate slide images. You can edit data without rendering.",
        "chart_type": "Chart Type",
        "slide_num": "Slide",
        "error_render": "Error rendering slide",
        "changes_saved": "Changes saved successfully",
        "editing_info": "Edit the data in the table. The chart preview updates instantly. Click 'Save to Presentation' to apply to the file.",
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
        "chart_comparison_toggle": "Before/After",
        "slide_comparison_toggle": "Before/After",
        "before": "Before",
        "after": "After",
        "chart_preview": "Chart Preview",
        "full_slide_preview": "Full Slide Preview",
        "save_to_pptx": "Save to Presentation",
        "saving_to_pptx": "Saving to presentation and rendering...",
        "saved_to_pptx": "Changes saved to presentation successfully",
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
        # Excel tab
        "tab_excel": "Excel Import/Export (All Charts)",
        "excel_export_title": "Export All Charts to Excel",
        "excel_export_caption": "Download a single .xlsx file with all {count} charts, each on its own sheet",
        "excel_export_button": "Export All to Excel",
        "excel_import_title": "Import Excel to Update Charts",
        "excel_import_caption": "Upload a .xlsx file previously exported. Sheets will be matched to charts by name.",
        "excel_import_upload_label": "Choose an Excel file (.xlsx)",
        "excel_matched_charts": "Matched {matched} of {total} charts",
        "excel_no_matches": "No sheets matched any charts. Sheet names must follow the format 'Slide{n}_ChartName'.",
        "excel_changes_found": "Found {changed} charts with changes out of {total} matched",
        "excel_unchanged": "{count} charts unchanged",
        "excel_no_changes": "No changes detected — the Excel data matches the current presentation",
        "excel_apply_button": "Apply All Matched Data",
        "excel_apply_spinner": "Updating {count} charts...",
        "excel_apply_success": "Successfully updated {count} charts",
        "excel_column_mismatch_warning": "Sheet '{sheet}': column count mismatch (expected {expected}, found {found}) - skipped",
        "excel_sheet_no_match": "Sheet '{sheet}' - no matching chart found",
        # Auto-save
        "auto_save_label": "Auto-save after updates",
        "auto_save_info": "Automatically downloads the updated file after clicking 'Save to Presentation'",
        "unsaved_warning": "You have unsaved edits. Click 'Save to Presentation' to apply.",
        "auto_saved_msg": "File auto-saved",
        "save_disabled_hint": "Edit data in the table to enable saving",
        # Sidebar filter
        "filter_slides": "Filter slides...",
        # Onboarding
        "onboarding_summary": "Found {slides} slides with {charts} charts. Select a slide from the sidebar to begin.",
        "app_subtitle": "Edit PowerPoint chart data quickly and easily",
        "upload_section_title": "Upload Presentation",
        "preview_section_title": "Preview",
        "editor_section_title": "Data Editor",
        "no_slide_selected": "Select a slide from the sidebar to start editing",
        "chart_details": "Chart Details",
        "actions": "Actions",
        # Undo
        "undo": "Undo",
        "undo_success": "Last change undone",
        "no_undo": "Nothing to undo",
        # Select Data tab
        "tab_select_data": "Select Data",
        "select_data_caption": "Choose which series to display in the chart. Hidden series keep their data but are not plotted.",
        "series_visible_label": "Visible Series",
        "at_least_one_series": "At least one series must be visible",
        "visibility_updated": "Series visibility updated successfully",
        "update_visibility": "Update Visibility",
        "hidden_series_count": "{count} series hidden",
        "all_series_visible": "All series visible",
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
        # Progress indicator
        "step_upload": "Upload",
        "step_select": "Edit All Charts",
        "step_edit": "Edit Slide",
        "step_download": "Download",
        # Workflow guidance
        "welcome_message": "Welcome to Chart Editor!",
        "getting_started": "Getting Started",
        "step_1": "Upload a PowerPoint file (.pptx)",
        "step_2": "Select a slide and chart to edit",
        "step_3": "Edit the data in the table",
        "step_4": "Download the updated file",
        # Accessibility
        "filter_help": "Filter slides by number or chart name",
        "series_visibility_help": "Choose which series to display in the chart. Hidden series keep their data but are not plotted.",
        "csv_help": "Export current data to CSV or import data from a CSV file",
        # Metrics
        "total_slides": "Total Slides",
        "total_charts": "Total Charts",
        "chart_types": "Chart Types",
        "overview": "Overview",
        "overview_description": "Select a slide from the sidebar to start editing. Each slide contains multiple charts that can be edited individually.",
        "tab_overview": "Overview",
        "charts": "Charts",
        "preview": "Preview",
        # Errors and feedback
        "csv_imported_success": "CSV imported successfully",
        "csv_import_error": "Error importing CSV",
        "applying_csv": "Applying data from CSV...",
        "undo_not_available": "Undo is not available at this time",
        "slide_count_info": "{count} slides with charts",
        # Getting Started wizard
        "wizard_welcome_title": "Welcome to Chart Editor!",
        "wizard_welcome_desc": "This tool lets you edit chart data directly inside PowerPoint presentations. Follow the steps below to get started.",
        "wizard_upload_title": "Upload Presentation",
        "wizard_upload_desc": "Drag a PowerPoint file (.pptx) into the upload area in the sidebar, or click 'Browse files' to select one.",
        "wizard_upload_waiting": "Waiting for file upload...",
        "wizard_upload_done": "File '{name}' uploaded successfully!",
        "wizard_select_title": "Edit All Charts",
        "wizard_select_desc": "Export all charts to an Excel file, edit them, and import back. A fast way to update all data at once.",
        "wizard_edit_title": "Edit Single Slide",
        "wizard_edit_desc": "Select a slide from the sidebar, edit a specific chart in the table, preview live, and save to the presentation.",
        "wizard_next": "Next",
        "wizard_back": "Back",
        "wizard_start_editing": "Start Editing",
        "wizard_upload_now": "Upload a file in the sidebar",
        "wizard_file_uploaded": "File uploaded",
        "quick_start": "Quick Start",
        # User Guide
        "tab_guide": "User Guide",
        "guide_overview_title": "Overview",
        "guide_overview_body": "PPTX Chart Editor lets you edit chart data inside PowerPoint presentations directly in your browser. Upload a .pptx file, modify chart values in an interactive table with a live preview, and download the updated presentation — no PowerPoint required.",
        "guide_start_title": "Getting Started",
        "guide_start_1": "**Upload** — Drag a .pptx file into the upload area in the sidebar, or click to browse",
        "guide_start_2": "**Select a slide** — Click a slide thumbnail in the sidebar to see its charts",
        "guide_start_3": "**Choose a chart** — Pick a chart from the dropdown to start editing",
        "guide_start_4": "**Edit data** — Modify values in the table; the chart preview updates instantly",
        "guide_start_5": "**Save** — Click \"Save to Presentation\" then download the updated file",
        "guide_edit_title": "Editing a Chart",
        "guide_edit_body": """- The data table shows categories (rows) and series (columns)
- Edit any cell — the Plotly chart preview updates in real time
- Use the **Before/After** toggle to compare your changes with the original
- Click **Save to Presentation** to write changes back to the .pptx file
- Use **Undo** to revert the last change""",
        "guide_excel_title": "Bulk Editing with Excel",
        "guide_excel_body": """- Go to the **Excel Import/Export** tab
- Click **Export All to Excel** — downloads a .xlsx with one sheet per chart
- Edit the Excel file in any spreadsheet app
- Upload the modified .xlsx — the app matches sheets to charts automatically
- Review the changes and click **Apply All Matched Data**""",
        "guide_csv_title": "CSV Import/Export",
        "guide_csv_body": """- In the **CSV** tab, export or import data for the currently selected chart
- The CSV must match the chart structure: first column = categories, remaining = series""",
        "guide_batch_title": "Add Row to All Charts",
        "guide_batch_body": """- In the **Batch** tab, enter a new category name (e.g., a survey date)
- Click **Add Row to All Charts** to add the same row to every chart at once""",
        "guide_visibility_title": "Series Visibility",
        "guide_visibility_body": """- In the **Select Data** tab, toggle which series are shown on the chart
- Hidden series keep their data but are not plotted
- At least one series must remain visible""",
        "guide_tips_title": "Tips",
        "guide_tips_body": """- **Percentages**: For percentage columns, enter 67 for 67% (automatic conversion)
- **Auto-save**: Enable "Auto-save after updates" to automatically download after each save
- **Slide filter**: Use the search box in the sidebar to filter slides by number or chart name""",
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
        @import url('https://fonts.googleapis.com/css2?family=Assistant:wght@300;400;600;700&display=swap');
        .stApp {
            direction: rtl;
            font-family: 'Assistant', sans-serif;
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
            font-family: 'Assistant', sans-serif;
            font-weight: 700;
        }

        /* Paragraphs and text */
        p, span, label, .stMarkdown {
            direction: rtl;
            text-align: right;
            font-family: 'Assistant', sans-serif;
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
