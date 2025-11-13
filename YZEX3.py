import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser

# ----- נתיב קובץ -----
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
file_path = os.path.join(desktop_path, 'YZEX.xlsx')

# ----- פונקציה לטעינת מאגר -----
def load_exercises():
    try:
        df = pd.read_excel(file_path)
        df.columns = [col.strip() for col in df.columns]
        return df
    except Exception as e:
        messagebox.showerror("שגיאה", f"לא ניתן לטעון את הקובץ: {e}")
        return pd.DataFrame()

exercises_df = load_exercises()

# ----- זיהוי עמודת לינק -----
possible_link_cols = ['לינק', 'קישור', 'Link', 'URL']
link_col_name = next((col for col in possible_link_cols if col in exercises_df.columns), None)

# ----- פתיחת לינק -----
def open_link(event):
    region = workout_tree.identify("region", event.x, event.y)
    if region != "cell":
        return

    row_id = workout_tree.identify_row(event.y)
    col_id = workout_tree.identify_column(event.x)
    if not row_id or not col_id:
        return

    col_index = int(col_id.replace("#", "")) - 1
    if link_col_name and columns[col_index] == link_col_name:
        url = original_links.get(row_id, None)
        if isinstance(url, str) and url.startswith("http"):
            webbrowser.open(url)
        else:
            messagebox.showinfo("מידע", "אין קישור תקין לשורה זו.")

# ----- פונקציה ליצירת אימון -----
def generate_workout():
    global exercises_df, original_links
    workout_tree.delete(*workout_tree.get_children())
    original_links = {}  # נאחסן את הקישורים המקוריים לכל שורה

    if exercises_df.empty:
        messagebox.showwarning("אזהרה", "המאגר ריק. רענן או בדוק את הקובץ.")
        return

    try:
        num_exercises = int(num_var.get())
    except:
        messagebox.showwarning("אזהרה", "מספר תרגילים לא חוקי")
        return

    selected_difficulty = [d for d, var in difficulty_vars.items() if var.get()]
    selected_equipment = [e for e, var in equipment_vars.items() if var.get()]

    if not selected_difficulty or not selected_equipment:
        messagebox.showwarning("אזהרה", "בחר לפחות רמת קושי וסוג ציוד אחד.")
        return

    df_filtered = exercises_df.copy()
    if 'רמת קושי' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['רמת קושי'].isin(selected_difficulty)]
    if 'סוג ציוד' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['סוג ציוד'].isin(selected_equipment)]

    if df_filtered.empty:
        messagebox.showwarning("אזהרה", "לא נמצאו תרגילים עם הבחירות שלך.")
        return

    possible_muscle_cols = ['קבוצת שריר', 'muscle group', 'Muscle', 'Muscle_Group']
    possible_name_cols = ['שם', 'תרגיל', 'Name', 'Exercise']

    muscle_col = next((col for col in possible_muscle_cols if col in df_filtered.columns), None)
    name_col = next((col for col in possible_name_cols if col in df_filtered.columns), None)

    if not muscle_col or not name_col:
        msg = f"לא נמצאו העמודות הדרושות בקובץ.\nעמודות קיימות: {df_filtered.columns.tolist()}"
        messagebox.showerror("שגיאה", msg)
        return

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

    # הצגה ב-GUI
    for ex in workout:
        row = []
        for col in columns:
            if col == link_col_name:
                url = ex[col] if col in ex else ""
                if isinstance(url, str) and url.startswith("http"):
                    row.append("🔗 פתח קישור")
                else:
                    row.append("")
            else:
                row.append(ex[col] if col in ex else "")
        item_id = workout_tree.insert("", "end", values=row)

        if link_col_name and link_col_name in ex and isinstance(ex[link_col_name], str):
            original_links[item_id] = ex[link_col_name]

# ----- פונקציה לרענון המאגר -----
def refresh_exercises():
    global exercises_df
    exercises_df = load_exercises()
    for widget in equipment_frame.winfo_children():
        widget.destroy()
    equipment_vars.clear()

    all_equipment = ["משקל גוף", "TRX", "דאמבלים", "גומיה"]
    for i, eq in enumerate(all_equipment):
        var = tk.BooleanVar(value=True)
        cb = tk.Checkbutton(equipment_frame, text=eq, variable=var, bg="#FFECB3", font=("Arial", 11))
        cb.grid(row=0, column=i, sticky="nsew", padx=5)
        equipment_frame.grid_columnconfigure(i, weight=1)
        equipment_vars[eq] = var

    messagebox.showinfo("הצלחה", "המאגר עודכן בהצלחה!")

# ----- GUI -----
root = tk.Tk()
root.title("YZ Exercise")
root.configure(bg="#FFD590")  # צבע רקע ראשי

base_width = 1450
base_height = 650
root.geometry(f"{int(base_width*1.1)}x{int(base_height*1.1)}")

# --- תוויות ---
tk.Label(root, text="NOE / כמות תרגילים", bg="#FFECB3", font=("Arial", 15, "bold")).pack()
num_var = tk.StringVar(value="5")
num_menu = ttk.Combobox(root, textvariable=num_var, values=[3,4,5,6,7,8], width=3, state="readonly", font=("Arial", 15))
num_menu.pack(pady=5)

# --- רמות קושי ---
difficulty_frame = tk.LabelFrame(root, text="Level \ רמה", bg="#FFECB3", fg="black", font=("Arial", 15, "bold"), padx=5, pady=5)
difficulty_frame.pack(pady=10)

difficulty_vars = {}
for i, diff in enumerate(["קל", "בינוני", "קשה"]):
    var = tk.BooleanVar(value=True)
    cb = tk.Checkbutton(difficulty_frame, text=diff, variable=var, bg="#FFECB3", font=("Arial", 15))
    cb.grid(row=0, column=i, sticky="nsew", padx=5)
    difficulty_frame.grid_columnconfigure(i, weight=1)
    difficulty_vars[diff] = var

# --- ציוד ---
equipment_frame = tk.LabelFrame(root, text="Equipment \ ציוד", bg="#FFECB3", fg="black", font=("Arial", 15, "bold"), padx=5, pady=5)
equipment_frame.pack(pady=5)

equipment_vars = {}
all_equipment = ["משקל גוף", "TRX", "דאמבלים", "גומיה"]
for i, eq in enumerate(all_equipment):
    var = tk.BooleanVar(value=True)
    cb = tk.Checkbutton(equipment_frame, text=eq, variable=var, bg="#FFECB3", font=("Arial", 15))
    cb.grid(row=0, column=i, sticky="nsew", padx=5)
    equipment_frame.grid_columnconfigure(i, weight=1)
    equipment_vars[eq] = var

# --- כפתורים ---
btn_frame = tk.Frame(root, bg="#FFD580")
btn_frame.pack(pady=10)
tk.Button(btn_frame, text="Create Exercise / צור אימון", command=generate_workout, width=25, bg="#FFA500", fg="Black", font=("Arial", 11, "bold")).pack(side="right", padx=5)
tk.Button(btn_frame, text="Refresh \ רענן", command=refresh_exercises, width=15, bg="#FFA500", fg="black", font=("Arial", 11, "bold")).pack(side="right", padx=5)

# --- טבלת אימון ---
table_frame = tk.Frame(root, bg="#FFD580")
table_frame.pack(fill="both", expand=True, padx=10, pady=10)

columns = exercises_df.columns.tolist()[::-1] if not exercises_df.empty else ["תרגיל"]
workout_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)

scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=workout_tree.yview)
scrollbar_y.pack(side="right", fill="y")
workout_tree.configure(yscrollcommand=scrollbar_y.set)

scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=workout_tree.xview)
scrollbar_x.pack(side="bottom", fill="x")
workout_tree.configure(xscrollcommand=scrollbar_x.set)

style = ttk.Style()
style.theme_use("xpnative")  # מאפשר צבע רקע לשורות
style.configure("Treeview", font=("Arial", 11, "bold"), background="#FFF3E0", foreground="black", fieldbackground="#FFF3E0", rowheight=45)
style.configure("Treeview.Heading", font=("Arial", 15, "bold"), background="#FFB74D", foreground="black")

for col in columns:
    workout_tree.heading(col, text=col, anchor='center')  # יישור לימין
    workout_tree.column(col, width=160, anchor='center')  # יישור לימין

workout_tree.pack(fill="both", expand=True)

# --- לחיצה כפולה לפתיחת קישור ---
if link_col_name:
    workout_tree.bind("<Double-1>", open_link)

original_links = {}

root.mainloop()
