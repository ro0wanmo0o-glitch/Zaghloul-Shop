import streamlit as st
import pandas as pd
import gspread

# ==========================================
# 1. الاتصال بـ Google Sheets مجاناً
# ==========================================
# ضعي هنا رابط الشيت الخاص بكِ بعد إتاحة صلاحية Editing لأي شخص مع الرابط
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/https://docs.google.com/spreadsheets/d/1VsHWS_D3Ebtm-YTYsBbYmjm3kYUg6BIelexAYoxpIxY/edit?usp=sharing/edit"

@st.cache_resource
def get_gsheet_client():
    # استخدام الاتصال المباشر عبر gspread
    gc = gspread.public_authorize() # أو الاتصال الفردي
    return gc

# دالة لقراءة صفحة معينة من الشيت
def read_sheet(sheet_name):
    try:
        # قراءة البيانات مباشرة عبر رابط CSV العام للشيت لسرعة ومجانية فائقة
        sheet_id = SPREADSHEET_URL.split("/d/")[1].split("/")[0]
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
        df = pd.read_csv(url)
        return df
    except Exception as e:
        return pd.DataFrame()

# دالة لتحديث صفحة (حذف / إضافة / تعديل)
def write_sheet(sheet_name, df):
    # يمكنك استخدام gspread لتحديث الصفحات
    gc = gspread.service_account_from_dict({}) # أو الحفظ المباشر
    sh = gc.open_by_url(SPREADSHEET_URL)
    worksheet = sh.worksheet(sheet_name)
    worksheet.clear()
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())

# تحميل البيانات
df_stock = read_sheet("Stock")
df_customers = read_sheet("Customers")
df_suppliers = read_sheet("Suppliers")

# ==========================================
# 2. الواجهة وتطبيق إدارة النواقص والموردين
# ==========================================
tab4, tab5 = st.tabs(["⚠️ جرد النواقص", "👥 إدارة الموردين والعملاء"])

# ------------------------------------------
# جرد النواقص
# ------------------------------------------
with tab4:
    st.markdown("### ⚠️ جرد النواقص")
    if not df_stock.empty and "الكمية المتاحة" in df_stock.columns:
        df_stock["الكمية المتاحة"] = pd.to_numeric(df_stock["الكمية المتاحة"], errors='coerce').fillna(0)
        low_stock = df_stock[df_stock["الكمية المتاحة"] <= 5]
        
        if not low_stock.empty:
            st.warning("الأصناف التالية وصل رصيدها إلى 5 قطع أو أقل:")
            st.dataframe(low_stock[["كود المنتج", "اسم المنتج", "الكمية المتاحة", "سعر البيع"]], use_container_width=True)
        else:
            st.success("جميع الأصناف متوفرة بكميات آمنة داخل المخزون.")
    else:
        st.info("لا توجد بيانات مخزون حالياً.")

# ------------------------------------------
# إدارة الموردين والعملاء
# ------------------------------------------
with tab5:
    st.subheader("👥 إدارة الموردين والعملاء")
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.markdown("### 📞 دليل العملاء")
        if not df_customers.empty:
            st.dataframe(df_customers, use_container_width=True)
        else:
            st.info("لا يوجد عملاء مسجلين.")
            
    with col_b:
        st.markdown("### 🏢 دليل الموردين")
        if not df_suppliers.empty:
            st.dataframe(df_suppliers, use_container_width=True)
        else:
            st.info("لا يوجد موردين مسجلين.")
