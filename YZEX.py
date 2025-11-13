# YZEX.py - Streamlit version with clickable links, RTL table, and link first column

import streamlit as st
import pandas as pd

# ----- כותרת ראשית -----
st.set_page_config(page_title="YZ Exercise", layout="wide")
st.title("YZ Exercise - Workout Generator")

# ----- טעינת מאגר -----
file_path = "YZEX.xlsx"

@st.cache_data
def load_exercises():
    try:
        df = pd.read_excel(file_path)
        df.columns = [col.strip() for col in df.columns]
        return df
    except Exception as e:
        st.error(f"לא ניתן לטעון את הקובץ: {e}")
        return pd.DataFrame()

exercises_df = load_exercises()

if exercises_df.empty:
    st.warning("המאגר ריק. ודא שהקובץ YZEX.xlsx נמצא באותה תיקיה.")
    st.stop()

# ----- צד שמאל: הגדרות -----
with st.sidebar:
    st.header("Settings")
    num_exercises = st.selectbox("NOE / כמות תרגילים", [3,4,5,6,7,8], index=2)

    st.subheader("Level / רמה")
    difficulty_options = ["קל", "בינוני", "קשה"]
    selected_difficulty = [diff for diff in difficulty_options if st.checkbox(diff, value=True)]

    st.subheader("Equipment / ציוד")
    equipment_options = ["משקל גוף", "TRX", "דאמבלים", "גומיה"]
    selected_equipment = [eq for eq in equipment_options if st.checkbox(eq, value=True)]

# ----- פילטר -----
df_filtered = exercises_df.copy()
if 'רמת קושי' in df_filtered.columns:
    df_filtered = df_filtered[df_filtered['רמת קושי'].isin(selected_difficulty)]
if 'סוג ציוד' in df_filtered.columns:
    df_filtered = df_filtered[df_filtered['סוג ציוד'].isin(selected_equipment)]

if df_filtered.empty:
    st.warning("לא נמצאו תרגילים עם הבחירות שלך.")
    st.stop()

# ----- זיהוי עמודות -----
possible_muscle_cols = ['קבוצת שריר', 'muscle group', 'Muscle', 'Muscle_Group']
possible_name_cols = ['שם', 'תרגיל', 'Name', 'Exercise']
possible_link_cols = ['לינק', 'קישור', 'Link', 'URL']

muscle_col = next((col for col in possible_muscle_cols if col in df_filtered.columns), None)
name_col = next((col for col in possible_name_cols if col in df_filtered.columns), None)
link_col = next((col for col in possible_link_cols if col in df_filtered.columns), None)

if not muscle_col or not name_col:
    st.error("לא נמצאו העמודות הדרושות בקובץ.")
    st.stop()

# ----- פונקציה ליצירת אימון -----
def generate_workout(df_filtered, num_exercises):
    df_shuffled = df_filtered.sample(frac=1)
    workout = []
    used_exercises = set()
    used_muscles = set()
    for _, ex in df_shuffled.iterrows():
        muscle = ex[muscle_col]
        name = ex[name_col]
        if name not in used_exercises and muscle not in used_muscles:
            workout.append(ex)
            used_exercises.add(name)
            used_muscles.add(muscle)
        if len(workout) >= num_exercises:
            break
    if len(workout) < num_exercises:
        remaining_count = num_exercises - len(workout)
        remaining_choices = df_filtered[~df_filtered[name_col].isin(used_exercises)]
        if not remaining_choices.empty:
            extra = remaining_choices.sample(min(remaining_count, len(remaining_choices)))
            for _, ex in extra.iterrows():
                if len(workout) >= num_exercises:
                    break
                name = ex[name_col]
                if name not in used_exercises:
                    workout.append(ex)
                    used_exercises.add(name)
    return pd.DataFrame(workout)

# ----- כפתור "צור אימון" -----
if st.button("Create Workout / צור אימון"):
    workout_df = generate_workout(df_filtered, num_exercises)

    # הפוך את סדר העמודות כדי שהלינק יהיה ראשון
    if link_col and link_col in workout_df.columns:
        cols = [link_col] + [c for c in workout_df.columns if c != link_col]
        workout_df = workout_df[cols]

    st.subheader("Workout Table / טבלת אימון")
    st.markdown(
        "טבלת האימון נוצרה לפי הבחירות שלך. לחץ על 🔗 כדי לפתוח לינק למדריך תרגיל.",
        unsafe_allow_html=True
    )

    # ----- יצירת טבלה HTML עם לינקים ו-RTL -----
    table_html = "<table style='width:100%; border-collapse: collapse; direction: rtl;'>"
    # כותרות
    table_html += "<tr>"
    for col in workout_df.columns:
        table_html += f"<th style='border: 1px solid black; padding: 8px; text-align:center'>{col}</th>"
    table_html += "</tr>"

    # שורות
    for _, row in workout_df.iterrows():
        table_html += "<tr>"
        for col in workout_df.columns:
            val = row[col]
            if col == link_col and isinstance(val, str) and val.startswith("http"):
                val = f"<a href='{val}' target='_blank'>🔗 פתח קישור</a>"
            table_html += f"<td style='border: 1px solid black; padding: 8px; text-align:center'>{val}</td>"
        table_html += "</tr>"

    table_html += "</table>"
    st.markdown(table_html, unsafe_allow_html=True)

    # כפתור רענון
    if st.button("Refresh / רענן"):
        st.experimental_rerun()
