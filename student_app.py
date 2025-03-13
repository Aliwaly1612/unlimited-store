import gspread

# إعداد الاتصال بGoogle Sheets
def setup_google_sheets():
    # استخدم رابط الملف مباشرةً
    gc = gspread.oauth()
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1HLLoPza8mCottyOUXt_ux1uVz4q_JhhEOSVWlwK2gWc/edit?usp=sharing').sheet1  # استبدل YOUR_SHEET_URL برابط الملف
    return sheet

# تحميل البيانات من Google Sheets
def load_data(sheet):
    return sheet.get_all_records()

# حفظ البيانات في Google Sheets
def save_data(sheet, data):
    sheet.clear()
    sheet.append_row(list(data[0].keys()))  # إضافة العناوين
    for row in data:
        sheet.append_row(list(row.values()))

# اختبار الكود
sheet = setup_google_sheets()
data = load_data(sheet)
print(data)  # عرض البيانات

# حفظ التغييرات (اختياري)
# save_data(sheet, data)