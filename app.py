import streamlit as st
import pandas as pd

# عنوان التطبيق
st.title("إدارة درجات الطلاب")

# تحميل ملف الإكسل
uploaded_file = st.file_uploader("اختر ملف الإكسل", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # عرض البيانات
    st.write("بيانات الطلاب:")
    st.dataframe(df)

    # البحث عن طالب
    search_option = st.radio("ابحث باستخدام:", ["م", "الاسم"])
    search_value = st.text_input("أدخل قيمة البحث")

    if search_value:
        if search_option == "م":
            matched_students = df[df['م'] == int(search_value)]
        else:
            matched_students = df[df['الاسم'].str.contains(search_value, case=False)]

        if len(matched_students) == 0:
            st.warning("لم يتم العثور على أي طالب بهذه القيمة")
        else:
            st.write("الطلاب المطابقون:")
            st.dataframe(matched_students)

    # إضافة عمود جديد
    st.subheader("إضافة عمود جديد")
    new_column_name = st.text_input("اسم العمود الجديد")
    scale = st.selectbox("اختر المقياس", [1, 10, 100])

    if st.button("إضافة العمود"):
        if new_column_name in df.columns:
            st.warning("العمود موجود بالفعل!")
        else:
            df[new_column_name] = scale
            st.success(f"تمت إضافة العمود {new_column_name} بنجاح!")
            st.dataframe(df)

    # حفظ التغييرات
    if st.button("حفظ التغييرات"):
        df.to_excel("updated_students.xlsx", index=False)
        st.success("تم حفظ التغييرات بنجاح!")