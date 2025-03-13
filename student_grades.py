import pandas as pd
from tkinter import *
from tkinter import messagebox, ttk

# دالة لتوحيد الحروف المتشابهة
def normalize_arabic_text(text):
    replacements = {
        'ى': 'ي',
        'إ': 'أ',
        'آ': 'أ',
        'ة': 'ه',
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text

# تحميل ملف الإكسل
def load_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء تحميل الملف: {e}")
        return None

# حفظ التغييرات في ملف الإكسل
def save_excel(df, file_path):
    try:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("نجاح", "تم حفظ التغييرات بنجاح!")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء حفظ الملف: {e}")

# عرض قائمة الطلاب المطابقة للبحث
def show_matched_students(event=None):
    search_value = entry_search.get().strip()

    if not search_value:
        messagebox.showwarning("تحذير", "يرجى إدخال قيمة 'م' أو جزء من الاسم")
        return

    # توحيد الحروف في قيمة البحث
    search_value_normalized = normalize_arabic_text(search_value)

    # البحث عن الطالب باستخدام 'م' أو الاسم
    if search_value.isdigit():
        matched_students = df[df['م'] == int(search_value)]
    else:
        matched_students = df[df['الاسم'].apply(normalize_arabic_text).str.contains(search_value_normalized, case=False, na=False)]

    if len(matched_students) == 0:
        messagebox.showinfo("نتيجة البحث", "لم يتم العثور على أي طالب بهذه القيمة أو الاسم")
    else:
        # عرض قائمة الطلاب المطابقة
        combo_students['values'] = [f"{row['م']} - {row['الاسم']}" for index, row in matched_students.iterrows()]
        combo_students.set('')  # إزالة الاسم القديم

# عرض بيانات الطالب المحدد
def show_student_data(event):
    selected_student = combo_students.get()
    if not selected_student:
        return

    student_id = int(selected_student.split(" - ")[0])
    student_data = df[df['م'] == student_id].iloc[0]

    # عرض البيانات في الواجهة
    label_id.config(text=f"م: {student_data['م']}")
    label_name.config(text=f"الاسم: {student_data['الاسم']}")

    # عرض الدرجات
    for col in df.columns[2:]:
        entry_grades[col].delete(0, END)  # مسح القيمة القديمة
        entry_grades[col].insert(0, student_data[col])  # إدخال القيمة الجديدة

# حفظ التغييرات
def save_changes():
    if not edit_enabled.get():
        messagebox.showwarning("تحذير", "التعديل غير مفعل. يرجى تفعيله أولاً.")
        return

    selected_student = combo_students.get()
    if not selected_student:
        messagebox.showwarning("تحذير", "يرجى اختيار طالب")
        return

    student_id = int(selected_student.split(" - ")[0])

    # تحديث البيانات في DataFrame
    for col in df.columns[2:]:
        value = entry_grades[col].get()
        try:
            df.at[student_id - 1, col] = float(value)  # قبول القيم العشرية
        except ValueError:
            messagebox.showwarning("تحذير", f"يرجى إدخال قيمة رقمية للعمود {col}")
            return  # إيقاف الحفظ إذا كانت القيمة غير رقمية

    # حفظ التغييرات في ملف الإكسل
    save_excel(df, 'students.xlsx')

# إضافة عمود جديد
def add_new_column():
    new_column_name = entry_new_column.get().strip()
    if not new_column_name:
        messagebox.showwarning("تحذير", "يرجى إدخال اسم العمود الجديد")
        return

    scale = scale_var.get()
    if scale not in [1, 10, 100]:
        messagebox.showwarning("تحذير", "يرجى اختيار مقياس صحيح (1، 10، أو 100)")
        return

    if new_column_name in df.columns:
        messagebox.showwarning("تحذير", "العمود موجود بالفعل")
        return

    df[new_column_name] = scale
    save_excel(df, 'students.xlsx')
    update_grades_frame()
    combo_columns['values'] = df.columns[2:]

# تحديث إطار الدرجات
def update_grades_frame():
    for widget in frame_grades.winfo_children():
        widget.destroy()

    entry_grades.clear()
    columns = df.columns[2:]
    for i in range(0, len(columns), 2):  # عرض عمودين جنبًا إلى جنب
        frame_row = Frame(frame_grades, bg="#f0f0f0")
        frame_row.pack(pady=10)

        for j in range(2):
            if i + j < len(columns):
                col = columns[i + j]
                frame_grade = Frame(frame_row, bg="#f0f0f0")
                frame_grade.pack(side="left", padx=10)

                label_grade = Label(frame_grade, text=f"{col}:", bg="#f0f0f0", font=("Arial", 12))
                label_grade.grid(row=0, column=0, padx=5, pady=5)

                entry_grade = Entry(frame_grade, font=("Arial", 12))
                entry_grade.grid(row=0, column=1, padx=5, pady=5)
                entry_grades[col] = entry_grade

# تعديل اسم العمود
def edit_column_name():
    selected_column = combo_columns.get()
    new_column_name = entry_new_column_name.get().strip()

    if not selected_column or not new_column_name:
        messagebox.showwarning("تحذير", "يرجى اختيار عمود وإدخال اسم جديد")
        return

    if new_column_name in df.columns:
        messagebox.showwarning("تحذير", "الاسم الجديد موجود بالفعل")
        return

    df.rename(columns={selected_column: new_column_name}, inplace=True)
    save_excel(df, 'students.xlsx')
    update_grades_frame()
    combo_columns['values'] = df.columns[2:]

# فتح القائمة المنسدلة تلقائيًا عند النقر على حقل البحث
def open_combobox(event):
    combo_students.event_generate('<Down>')

# دالة لإعادة تحميل ملف الإكسل
def reload_excel():
    global df
    df = load_excel('students.xlsx')
    if df is not None:
        messagebox.showinfo("نجاح", "تم إعادة تحميل الملف بنجاح!")
        update_grades_frame()
        combo_columns['values'] = df.columns[2:]
    else:
        messagebox.showerror("خطأ", "حدث خطأ أثناء إعادة تحميل الملف")

# إنشاء الواجهة الرئيسية مع Scrollbar
root = Tk()
root.title("إدارة درجات الطلاب")
root.geometry("1000x700")  # تحديد حجم النافذة
root.configure(bg="#f0f0f0")  # لون خلفية النافذة

# إنشاء Canvas وإضافة Scrollbar
canvas = Canvas(root, bg="#f0f0f0")
scrollbar = Scrollbar(root, orient="vertical", command=canvas.yview)
scrollable_frame = Frame(canvas, bg="#f0f0f0")

# توصيل Canvas بالScrollbar
canvas.configure(yscrollcommand=scrollbar.set)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

# تحديث Scrollbar عند تغيير حجم النافذة
scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

# تمكين التمرير باستخدام عجلة الماوس
def on_mouse_wheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

canvas.bind_all("<MouseWheel>", on_mouse_wheel)

# وضع Canvas وScrollbar في الواجهة
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# زر الحفظ في الأعلى
button_save = Button(scrollable_frame, text="حفظ التغييرات", command=save_changes, bg="#4CAF50", fg="white", font=("Arial", 12))
button_save.pack(pady=20)

# مربع اختيار لتحديد ما إذا كان التعديل مفعلًا
edit_enabled = BooleanVar(value=True)
checkbox_edit = Checkbutton(scrollable_frame, text="تفعيل التعديل", variable=edit_enabled, bg="#f0f0f0", font=("Arial", 12))
checkbox_edit.pack(pady=10)

# إطار البحث
frame_search = Frame(scrollable_frame, bg="#f0f0f0")
frame_search.pack(pady=20)

label_search = Label(frame_search, text="ابحث بقيمة 'م' أو جزء من الاسم:", bg="#f0f0f0", font=("Arial", 12))
label_search.grid(row=0, column=0, padx=5, pady=5)

entry_search = Entry(frame_search, font=("Arial", 12))
entry_search.grid(row=0, column=1, padx=5, pady=5)
entry_search.bind('<Button-1>', open_combobox)  # فتح القائمة المنسدلة عند النقر
entry_search.bind('<Return>', show_matched_students)  # البحث عند الضغط على Enter

button_search = Button(frame_search, text="بحث", command=show_matched_students, bg="#4CAF50", fg="white", font=("Arial", 12))
button_search.grid(row=0, column=2, padx=5, pady=5)

# قائمة الطلاب المطابقة
combo_students = ttk.Combobox(scrollable_frame, state="readonly", width=50, font=("Arial", 12))
combo_students.pack(pady=20)
combo_students.bind('<<ComboboxSelected>>', show_student_data)

# إطار بيانات الطالب
frame_student_data = Frame(scrollable_frame, bg="#f0f0f0")
frame_student_data.pack(pady=20)

label_id = Label(frame_student_data, text="م: ", bg="#f0f0f0", font=("Arial", 12))
label_id.pack(pady=10)

label_name = Label(frame_student_data, text="الاسم: ", bg="#f0f0f0", font=("Arial", 12))
label_name.pack(pady=10)

# إطار الدرجات
frame_grades = Frame(scrollable_frame, bg="#f0f0f0")
frame_grades.pack(pady=20)

entry_grades = {}

# إطار إضافة عمود جديد
frame_new_column = Frame(scrollable_frame, bg="#f0f0f0")
frame_new_column.pack(pady=20)

label_new_column = Label(frame_new_column, text="اسم العمود الجديد:", bg="#f0f0f0", font=("Arial", 12))
label_new_column.grid(row=0, column=0, padx=5, pady=5)

entry_new_column = Entry(frame_new_column, font=("Arial", 12))
entry_new_column.grid(row=0, column=1, padx=5, pady=5)

scale_var = IntVar(value=1)
scale_1 = Radiobutton(frame_new_column, text="1", variable=scale_var, value=1, bg="#f0f0f0", font=("Arial", 12))
scale_1.grid(row=0, column=2, padx=5, pady=5)
scale_10 = Radiobutton(frame_new_column, text="10", variable=scale_var, value=10, bg="#f0f0f0", font=("Arial", 12))
scale_10.grid(row=0, column=3, padx=5, pady=5)
scale_100 = Radiobutton(frame_new_column, text="100", variable=scale_var, value=100, bg="#f0f0f0", font=("Arial", 12))
scale_100.grid(row=0, column=4, padx=5, pady=5)

button_add_column = Button(frame_new_column, text="إضافة عمود", command=add_new_column, bg="#4CAF50", fg="white", font=("Arial", 12))
button_add_column.grid(row=0, column=5, padx=5, pady=5)

# إطار تعديل اسم العمود
frame_edit_column = Frame(scrollable_frame, bg="#f0f0f0")
frame_edit_column.pack(pady=20)

label_columns = Label(frame_edit_column, text="اختر العمود:", bg="#f0f0f0", font=("Arial", 12))
label_columns.grid(row=0, column=0, padx=5, pady=5)

combo_columns = ttk.Combobox(frame_edit_column, state="readonly", font=("Arial", 12))
combo_columns.grid(row=0, column=1, padx=5, pady=5)

label_new_column_name = Label(frame_edit_column, text="اسم العمود الجديد:", bg="#f0f0f0", font=("Arial", 12))
label_new_column_name.grid(row=0, column=2, padx=5, pady=5)

entry_new_column_name = Entry(frame_edit_column, font=("Arial", 12))
entry_new_column_name.grid(row=0, column=3, padx=5, pady=5)

button_edit_column = Button(frame_edit_column, text="تعديل اسم العمود", command=edit_column_name, bg="#4CAF50", fg="white", font=("Arial", 12))
button_edit_column.grid(row=0, column=4, padx=5, pady=5)

# زر إعادة تحميل الملف
button_reload = Button(scrollable_frame, text="إعادة تحميل الملف", command=reload_excel, bg="#4CAF50", fg="white", font=("Arial", 12))
button_reload.pack(pady=20)

# تحميل ملف الإكسل
df = load_excel('students.xlsx')

# إذا لم يكن الملف موجودًا، إنشاء ملف جديد
if df is None:
    df = pd.DataFrame({'م': [], 'الاسم': []})
    save_excel(df, 'students.xlsx')

# تحديث قائمة الأعمدة
combo_columns['values'] = df.columns[2:]

# إنشاء حقول إدخال الدرجات
update_grades_frame()

root.mainloop()