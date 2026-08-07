import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
import re

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

EXCEL_FILE = "Items Body Care.xlsx"
SUPPLIERS_FILE = "Suppliers.xlsx"
CUSTOMERS_FILE = "Customers.xlsx"

st.set_page_config(page_title="ZAGHLOUL - World Of Care", page_icon="🛍️", layout="wide")

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.user_name = ""

if "sales_history" not in st.session_state:
    st.session_state.sales_history = []

def sort_stock_df(df):
    """ترتيب جدول المخزون حسب الكود (الكلمة الحرفية أولاً ثم الرقم تسلسلياً)"""
    if df.empty or "كود المنتج" not in df.columns:
        return df
    
    df_sorted = df.copy()
    
    def extract_sort_keys(code):
        code_str = str(code).strip()
        parts = code_str.split("-")
        prefix = parts[0]
        num = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
        return (prefix, num)

    keys = df_sorted["كود المنتج"].apply(extract_sort_keys)
    df_sorted["_prefix"] = [k[0] for k in keys]
    df_sorted["_num"] = [k[1] for k in keys]
    
    df_sorted = df_sorted.sort_values(by=["_prefix", "_num"]).drop(columns=["_prefix", "_num"]).reset_index(drop=True)
    return df_sorted

def init_empty_excel():
    columns = ["كود المنتج", "اسم المنتج", "سعر الشراء (التكلفة)", "سعر البيع", "الكمية المتاحة", "تاريخ الإضافة"]
    df_empty = pd.DataFrame(columns=columns)
    df_empty.to_excel(EXCEL_FILE, index=False)

def init_suppliers_excel():
    columns = ["اسم المورد", "رقم الهاتف", "ملاحظات / الأصناف الموردة"]
    df_empty = pd.DataFrame(columns=columns)
    df_empty.to_excel(SUPPLIERS_FILE, index=False)

def init_customers_excel():
    columns = ["اسم العميل", "رقم الهاتف", "تاريخ التسجيل", "الموظف المسجل"]
    df_empty = pd.DataFrame(columns=columns)
    df_empty.to_excel(CUSTOMERS_FILE, index=False)

if not os.path.exists(EXCEL_FILE):
    init_empty_excel()

if not os.path.exists(SUPPLIERS_FILE):
    init_suppliers_excel()

if not os.path.exists(CUSTOMERS_FILE):
    init_customers_excel()

def normalize_arabic(text):
    if not text:
        return ""
    text = re.sub(r"[إأآا]", "ا", text)
    text = re.sub(r"ى", "ي", text)
    text = re.sub(r"ؤ", "و", text)
    text = re.sub(r"ئ", "ي", text)
    text = re.sub(r"ة", "ه", text)
    return text.strip().lower()

def get_auto_code_prefix(product_name):
    p_name = normalize_arabic(product_name)
    
    rules = [
        ("بادي سبلاش", "BSP"),
        ("بادي", "BSP"),
        ("سبلاش", "BSP"),
        ("مزيل عرق كريم", "DCR"),
        ("مزيل عرق سبراي", "DSP"),
        ("مزيل عرق", "DCR"),
        ("مزيل", "DCR"),
        ("مسك طهاره", "MST"),
        ("مسك", "MST"),
        ("مخمريه", "MKH"),
        ("برفان", "PRF"),
        ("عطر", "PRF"),
        ("حلق", "ERG"),
        ("سلسله", "NLK"),
        ("طرحه", "SCV"),
        ("بندانه", "BND"),
        ("معصم", "SLV"),
        ("شنطه هدايا", "GBG"),
        ("بوكس هدايا", "GBX"),
        ("بوكس", "GBX"),
        ("شيت ماسك", "SMK"),
        ("ماسك", "SMK"),
        ("اظافر", "NLC"),
        ("ظافر", "NLC"),
        ("عنايه باليدين", "HNC"),
        ("يدين", "HNC"),
        ("عنايه بالقدمين", "FTC"),
        ("قدمين", "FTC"),
        ("شنطه ظهر", "BPK"),
        ("شنطه كروس", "CRB"),
        ("محفظه", "WLT"),
        ("شنطه", "CRB")
    ]
    
    for key, code in rules:
        if normalize_arabic(key) in p_name:
            return code
            
    return "ZAG"

def generate_sequential_code(product_name, df_stock_current):
    prefix = get_auto_code_prefix(product_name)
    
    max_num = 100
    if not df_stock_current.empty and "كود المنتج" in df_stock_current.columns:
        existing_codes = df_stock_current["كود المنتج"].astype(str).tolist()
        for code in existing_codes:
            if code.startswith(f"{prefix}-"):
                parts = code.split("-")
                if len(parts) > 1 and parts[1].isdigit():
                    max_num = max(max_num, int(parts[1]))
    
    next_num = max_num + 1
    return f"{prefix}-{next_num}"

def generate_pdf_report(today_rev, today_profit, month_rev, month_profit, stock_cost, stock_sale, total_actual_profit):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=18,
        alignment=1,
        spaceAfter=20
    )
    
    elements = []
    elements.append(Paragraph("ZAGHLOUL - World Of Care", title_style))
    elements.append(Paragraph("Financial & Inventory Summary Report", styles['Heading2']))
    elements.append(Paragraph(f"Generated Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles['Normal']))
    elements.append(Spacer(1, 20))
    
    data = [
        ["Metric Description", "Value (EGP)"],
        ["Today Total Sales", f"{today_rev:,.2f}"],
        ["Today Net Profit", f"{today_profit:,.2f}"],
        ["Current Month Total Sales", f"{month_rev:,.2f}"],
        ["Current Month Net Profit", f"{month_profit:,.2f}"],
        ["Total Actual Net Sales Profit", f"{total_actual_profit:,.2f}"],
        ["Total Invested Capital (Stock Cost)", f"{stock_cost:,.2f}"],
        ["Total Stock Expected Sales Value", f"{stock_sale:,.2f}"]
    ]
    
    t = Table(data, colWidths=[300, 150])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1B365D")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#F8F9FA")),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#CCCCCC")),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
    ]))
    
    elements.append(t)
    doc.build(elements)
    buffer.seek(0)
    return buffer

if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center; color: #1B365D;'>ZAGHLOUL - World Of Care</h1>", unsafe_allow_html=True)
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("🔑 تسجيل الدخول للنظام")
        username = st.text_input("اسم المستخدم")
        password = st.text_input("كلمة السر", type="password")
        
        if st.button("دخول للنظام", type="primary", use_container_width=True):
            if username == "admin" and password == "182":
                st.session_state.logged_in = True
                st.session_state.role = "admin"
                st.session_state.user_name = "محمد زغلول"
                st.rerun()
            elif username == "emp1" and password == "132":
                st.session_state.logged_in = True
                st.session_state.role = "cashier"
                st.session_state.user_name = "موظف 1"
                st.rerun()
            elif username == "emp2" and password == "232":
                st.session_state.logged_in = True
                st.session_state.role = "cashier"
                st.session_state.user_name = "موظف 2"
                st.rerun()
            else:
                st.error("❌ اسم المستخدم أو كلمة السر غير صحيحة!")

else:
    st.sidebar.title(f"👤 مرحباً: {st.session_state.user_name}")
    if st.sidebar.button("تسجيل الخروج"):
        st.session_state.logged_in = False
        st.rerun()

    st.markdown("# ZAGHLOUL - World Of Care")
    st.markdown(f"##### المستخدم الحالي: **{st.session_state.user_name}**")
    st.markdown("---")

    try:
        df_stock = pd.read_excel(EXCEL_FILE)
        df_stock = sort_stock_df(df_stock)
    except Exception:
        df_stock = pd.DataFrame(columns=["كود المنتج", "اسم المنتج", "سعر الشراء (التكلفة)", "سعر البيع", "الكمية المتاحة", "تاريخ الإضافة"])

    try:
        df_suppliers = pd.read_excel(SUPPLIERS_FILE)
    except Exception:
        df_suppliers = pd.DataFrame(columns=["اسم المورد", "رقم الهاتف", "ملاحظات / الأصناف الموردة"])

    try:
        df_customers = pd.read_excel(CUSTOMERS_FILE)
    except Exception:
        df_customers = pd.DataFrame(columns=["اسم العميل", "رقم الهاتف", "تاريخ التسجيل", "الموظف المسجل"])

    if not df_stock.empty and "الكمية المتاحة" in df_stock.columns:
        low_stock_items = df_stock[df_stock["الكمية المتاحة"] <= 5]
        if not low_stock_items.empty:
            st.error(f"🚨 **تنبيه هام:** يوجد ({len(low_stock_items)}) منتجات أوشكت على النفاد!")

    if st.session_state.role == "admin":
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "🛒 شاشة المبيعات", 
            "📦 إضافة مشتريات جديدة", 
            "📋 إدارة المخزون والجرد", 
            "📊 التقارير والأرباح والجرد المالي",
            "👥 إدارة الموردين والعملاء"
        ])
    else:
        tab1, tab3 = st.tabs(["🛒 شاشة المبيعات", "📋 عرض المخزون المتاح"])
        tab2, tab4, tab5 = None, None, None

    # 1. شاشة المبيعات
    with tab1:
        st.subheader("تسجيل فاتورة بيع جديدة وحفظ العميل")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            customer_name = st.text_input("اسم المشتري / العميل", value="عميل نقدي")
        with col_c2:
            customer_phone = st.text_input("رقم هاتف العميل", value="")
            
        if st.button("💾 تسجيل/حفظ بيانات العميل فقط"):
            c_name_clean = customer_name.strip()
            c_phone_clean = customer_phone.strip()
            
            if c_name_clean and c_name_clean != "عميل نقدي":
                new_cust = {
                    "اسم العميل": c_name_clean,
                    "رقم الهاتف": c_phone_clean,
                    "تاريخ التسجيل": str(datetime.now().strftime("%Y-%m-%d %H:%M")),
                    "الموظف المسجل": st.session_state.user_name
                }
                df_customers = pd.concat([df_customers, pd.DataFrame([new_cust])], ignore_index=True)
                df_customers.to_excel(CUSTOMERS_FILE, index=False)
                st.success(f"تم تسجيل وحفظ العميل ({c_name_clean}) بنجاح لتصل تلقائياً لسجل المدير!")
                st.rerun()
            else:
                st.warning("يرجى إدخال اسم عميل صحيح للتسجيل.")

        st.markdown("---")
        
        chosen_item = ""
        selected_code = ""
        selected_name = ""
        cost_price = 0.0
        sell_price = 0.0
        available_qty = 0

        if not df_stock.empty:
            item_options = [""] + [f"{row['اسم المنتج']} | [كود: {row['كود المنتج']}]" for _, row in df_stock.iterrows()]
            chosen_item = st.selectbox("ابحث عن اسم المنتج (اكتب اسم أو كود المنتج):", options=item_options, key="sale_product_select")
            
            if chosen_item:
                selected_name = chosen_item.split(" | [كود: ")[0]
                extracted_code = chosen_item.split(" | [كود: ")[1].replace("]", "").strip()
                
                match = df_stock[df_stock["كود المنتج"] == extracted_code]
                if not match.empty:
                    idx = match.index[0]
                    selected_code = extracted_code
                    cost_price = float(df_stock.at[idx, "سعر الشراء (التكلفة)"])
                    sell_price = float(df_stock.at[idx, "سعر البيع"])
                    available_qty = int(df_stock.at[idx, "الكمية المتاحة"])

            col_info1, col_info2, col_info3 = st.columns(3)
            with col_info1:
                st.info(f"📌 **الكود:** {selected_code if selected_code else 'غير محدد'}")
            with col_info2:
                st.info(f"💵 **سعر البيع:** {sell_price:,.2f} ج.م")
            with col_info3:
                st.info(f"📦 **المتاح:** {available_qty} قطعة")
        else:
            st.info("لا توجد منتجات في المخزون بعد.")

        col_qty1, col_qty2 = st.columns(2)
        with col_qty1:
            qty = st.number_input("الكمية المباعة", min_value=1, value=1)
            if chosen_item and available_qty == 0:
                st.error("⚠️ هذا المنتج نفد بالكامل من المخزون!")
        with col_qty2:
            if st.session_state.role == "admin":
                sale_date = st.date_input("تاريخ الفاتورة", value=datetime.now())
            else:
                sale_date = datetime.now().date()
                st.info(f"تاريخ الفاتورة: {sale_date}")

        if st.button("💾 تسجيل عملية البيع", type="primary"):
            if not selected_code:
                st.error("يرجى اختيار منتج أولاً!")
            else:
                final_code = selected_code
                final_name = selected_name
                
                match_idx = df_stock[df_stock["كود المنتج"] == final_code].index
                if not match_idx.empty:
                    idx = match_idx[0]
                    cost_price = float(df_stock.at[idx, "سعر الشراء (التكلفة)"])
                    sell_price = float(df_stock.at[idx, "سعر البيع"])
                    
                    current_qty = int(df_stock.at[idx, "الكمية المتاحة"])
                    new_qty = max(0, current_qty - qty)
                    df_stock.at[idx, "الكمية المتاحة"] = new_qty
                    
                    df_stock = sort_stock_df(df_stock)
                    df_stock.to_excel(EXCEL_FILE, index=False)

                c_name_clean = customer_name.strip() if customer_name.strip() else "عميل نقدي"
                c_phone_clean = customer_phone.strip()

                if c_name_clean != "عميل نقدي":
                    new_cust = {
                        "اسم العميل": c_name_clean,
                        "رقم الهاتف": c_phone_clean,
                        "تاريخ التسجيل": str(datetime.now().strftime("%Y-%m-%d %H:%M")),
                        "الموظف المسجل": st.session_state.user_name
                    }
                    df_customers = pd.concat([df_customers, pd.DataFrame([new_cust])], ignore_index=True)
                    df_customers.to_excel(CUSTOMERS_FILE, index=False)

                st.session_state.sales_history.append({
                    "التاريخ": str(sale_date),
                    "العميل": c_name_clean,
                    "رقم الهاتف": c_phone_clean,
                    "كود المنتج": final_code,
                    "اسم المنتج": final_name,
                    "الكمية": qty,
                    "تكلفة القطعة": cost_price,
                    "سعر البيع للقطعة": sell_price,
                    "إجمالي البيع": qty * sell_price,
                    "صافي الربح": qty * (sell_price - cost_price),
                    "الموظف": st.session_state.user_name
                })
                st.success("تم تسجيل عملية البيع وحفظ بيانات العميل بنجاح!")
                st.rerun()

        st.markdown("---")
        st.subheader("📋 سجل عمليات البيع المسجلة")
        if st.session_state.sales_history:
            sales_df = pd.DataFrame(st.session_state.sales_history)
            st.dataframe(sales_df, use_container_width=True)

            if st.session_state.role == "admin":
                st.markdown("---")
                st.markdown("##### 🔴 إلغاء / حذف عملية بيع:")
                sale_options = [f"عملية رقم {idx+1}: [{row['كود المنتج']}] {row['اسم المنتج']} - {row['إجمالي البيع']} ج.م ({row['الموظف']})" for idx, row in sales_df.iterrows()]
                selected_sale_to_delete = st.selectbox("اختر عملية البيع المراد إلغائها:", options=[""] + sale_options)

                if st.button("حذف عملية البيع وإعادة الكمية للمخزون"):
                    if selected_sale_to_delete:
                        sale_idx = int(selected_sale_to_delete.split(":")[0].replace("عملية رقم ", "").strip()) - 1
                        sale_item = st.session_state.sales_history[sale_idx]
                        
                        item_code = sale_item.get("كود المنتج", "")
                        qty_sold = sale_item.get("الكمية", 0)
                        if item_code and item_code != "بدون كود":
                            match_idx = df_stock[df_stock["كود المنتج"] == item_code].index
                            if not match_idx.empty:
                                idx = match_idx[0]
                                df_stock.at[idx, "الكمية المتاحة"] = int(df_stock.at[idx, "الكمية المتاحة"]) + qty_sold
                                df_stock = sort_stock_df(df_stock)
                                df_stock.to_excel(EXCEL_FILE, index=False)

                        st.session_state.sales_history.pop(sale_idx)
                        st.success("تم إلغاء عملية البيع وإعادة الكمية للمخزون!")
                        st.rerun()
                    else:
                        st.warning("يرجى اختيار عملية بيع أولاً.")
            else:
                st.info("🔒 ملاحظة: إلغاء أو حذف الفواتير المسجلة متاح فقط للمدير المسؤول.")
        else:
            st.info("لم يتم تسجيل أي عمليات بيع حتى الآن.")

    # 2. إضافة مشتريات
    if tab2 and st.session_state.role == "admin":
        with tab2:
            st.subheader("📦 إضافة منتج جديد للمخزون")
            
            col_s1, col_s2, col_s3 = st.columns(3)
            with col_s1:
                supplier_name = st.text_input("اسم المورد")
            with col_s2:
                supplier_phone = st.text_input("رقم هاتف المورد")
            with col_s3:
                supplier_notes = st.text_input("ملاحظات عن الأصناف الموردة")

            if st.button("💾 حفظ بيانات المورد فقط في السجل"):
                if supplier_name.strip():
                    new_sup = {
                        "اسم المورد": supplier_name.strip(),
                        "رقم الهاتف": supplier_phone.strip(),
                        "ملاحظات / الأصناف الموردة": supplier_notes.strip()
                    }
                    df_suppliers = pd.concat([df_suppliers, pd.DataFrame([new_sup])], ignore_index=True)
                    df_suppliers.to_excel(SUPPLIERS_FILE, index=False)
                    st.success(f"تم حفظ المورد ({supplier_name}) في سجل الموردين!")
                    st.rerun()
                else:
                    st.error("يرجى إدخال اسم المورد أولاً.")

            st.markdown("---")
            col_p1, col_p2, col_p3 = st.columns(3)

            with col_p1:
                p_name = st.text_input("اسم المنتج بالعربي")
                p_code_custom = st.text_input("كود المنتج (اتركه فارغاً للتوليد التلقائي والتسلسلي)", help="اتركه فارغاً لإنشاء كود تلقائي مثل BSP-101 ثم BSP-102")

            with col_p2:
                p_cost = st.number_input("سعر التكلفة", min_value=0.0, value=0.0)
                p_price = st.number_input("سعر البيع للجمهور", min_value=0.0, value=0.0)

            with col_p3:
                p_qty = st.number_input("الكمية", min_value=0, value=1)
                p_date = st.date_input("تاريخ الشراء", value=datetime.now())

            if st.button("➕ إضافة المنتج للمخزون", type="primary"):
                final_name = p_name.strip()
                if not final_name:
                    st.error("يرجى إدخال اسم المنتج بالعربي أولاً!")
                else:
                    if p_code_custom.strip():
                        final_code = p_code_custom.strip()
                    else:
                        final_code = generate_sequential_code(final_name, df_stock)

                    new_row = {
                        "كود المنتج": final_code,
                        "اسم المنتج": final_name,
                        "سعر الشراء (التكلفة)": p_cost,
                        "سعر البيع": p_price,
                        "الكمية المتاحة": p_qty,
                        "تاريخ الإضافة": str(p_date)
                    }
                    df_stock = pd.concat([df_stock, pd.DataFrame([new_row])], ignore_index=True)
                    
                    # إعادة الفرز والترتيب حسب الكود
                    df_stock = sort_stock_df(df_stock)
                    df_stock.to_excel(EXCEL_FILE, index=False)

                    if supplier_name.strip():
                        new_sup = {
                            "اسم المورد": supplier_name.strip(),
                            "رقم الهاتف": supplier_phone.strip(),
                            "ملاحظات / الأصناف الموردة": f"توريد: {final_name} - {supplier_notes.strip()}".strip(" -")
                        }
                        df_suppliers = pd.concat([df_suppliers, pd.DataFrame([new_sup])], ignore_index=True)
                        df_suppliers.to_excel(SUPPLIERS_FILE, index=False)

                    st.success(f"تمت إضافة المنتج ({final_name}) وترتيب المخزون بنجاح بالكود ({final_code})!")
                    st.rerun()

    # 3. إدارة المخزون والجرد
    with tab3:
        st.subheader("📋 جدول جرد وتعديل المخزون (مرتب كودياً تلقائياً)")
        
        if not df_stock.empty:
            if st.session_state.role == "cashier":
                cols_to_show = [c for c in df_stock.columns if "التكلفة" not in c and "الشراء" not in c]
                st.dataframe(df_stock[cols_to_show], use_container_width=True)
            else:
                st.info("💡 **طريقة التعديل المباشر:** اضغط على الخلية المراد تعديلها، ثم اضغط زر **'حفظ التعديلات'** بالأسفل ليتم الترتيب وتحديث البيانات.")
                
                edited_df = st.data_editor(
                    df_stock,
                    use_container_width=True,
                    num_rows="dynamic",
                    disabled=["تاريخ الإضافة"],
                    key="stock_editor"
                )

                if st.button("💾 حفظ التعديلات على المخزون", type="primary"):
                    edited_df = sort_stock_df(edited_df)
                    edited_df.to_excel(EXCEL_FILE, index=False)
                    st.success("تم تحديث المخزون وإعادة ترتيبه كودياً بنجاح!")
                    st.rerun()

                st.markdown("---")
                st.subheader("🗑️ قسم حذف وإدارة المخزون")
                col_del1, col_del2 = st.columns(2)
                
                with col_del1:
                    st.markdown("##### 🔴 حذف منتج محدد:")
                    delete_options = [f"{row['كود المنتج']} - {row['اسم المنتج']}" for _, row in df_stock.iterrows()]
                    item_to_delete = st.selectbox("اختر المنتج المراد حذفه:", options=[""] + delete_options)
                    
                    if st.button("حذف المنتج المحدد"):
                        if item_to_delete:
                            code_to_del = item_to_delete.split(" - ")[0]
                            df_stock = df_stock[df_stock["كود المنتج"] != code_to_del]
                            df_stock = sort_stock_df(df_stock)
                            df_stock.to_excel(EXCEL_FILE, index=False)
                            st.success("تم حذف المنتج المحدد من المخزون بنجاح!")
                            st.rerun()
                        else:
                            st.warning("يرجى اختيار منتج أولاً.")

                with col_del2:
                    st.markdown("##### 🚨 تفريغ المخزون بالكامل:")
                    if st.button("مسح كافة منتجات المخزون"):
                        init_empty_excel()
                        st.success("تم مسح المخزون بنجاح!")
                        st.rerun()
        else:
            st.info("المخزون فارغ حالياً.")

    # 4. التقارير المالية والتحميل بـ PDF
    if tab4 and st.session_state.role == "admin":
        with tab4:
            st.subheader("📊 التقارير المالية والأرباح")
            
            pdf_today_rev = 0.0
            pdf_today_profit = 0.0
            pdf_month_rev = 0.0
            pdf_month_profit = 0.0
            pdf_stock_cost = 0.0
            pdf_stock_sale = 0.0
            total_actual_profit = 0.0

            st.markdown("### 💵 تقرير مبيعات وأرباح اليوم")
            if st.session_state.sales_history:
                df_s = pd.DataFrame(st.session_state.sales_history)
                df_s["التاريخ_dt"] = pd.to_datetime(df_s["التاريخ"])
                today_str = str(datetime.now().date())
                
                df_today = df_s[df_s["التاريخ"] == today_str]
                if not df_today.empty:
                    pdf_today_rev = float(df_today["إجمالي البيع"].sum())
                    pdf_today_profit = float(df_today["صافي الربح"].sum())
                    today_orders = len(df_today)

                    d1, d2, d3 = st.columns(3)
                    d1.metric("مبيعات اليوم الإجمالية", f"{pdf_today_rev:,.2f} ج.م")
                    d2.metric("صافي أرباح اليوم", f"{pdf_today_profit:,.2f} ج.م")
                    d3.metric("عدد عمليات البيع اليوم", f"{today_orders}")
                else:
                    st.info("لا توجد عمليات بيع مسجلة بتاريخ اليوم.")
            else:
                st.info("لا توجد مبيعات مسجلة اليوم.")

            st.markdown("---")

            current_month = datetime.now().month
            current_year = datetime.now().year
            st.markdown(f"### 📅 تقرير مبيعات وأرباح الشهر الحالي ({current_month}/{current_year})")
            
            if st.session_state.sales_history:
                df_month = df_s[(df_s["التاريخ_dt"].dt.month == current_month) & (df_s["التاريخ_dt"].dt.year == current_year)]
                if not df_month.empty:
                    pdf_month_rev = float(df_month["إجمالي البيع"].sum())
                    pdf_month_profit = float(df_month["صافي الربح"].sum())
                    month_orders = len(df_month)
                    avg_order_val = pdf_month_rev / month_orders if month_orders > 0 else 0.0

                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("إجمالي مبيعات الشهر", f"{pdf_month_rev:,.2f} ج.م")
                    m2.metric("إجمالي أرباح الشهر الصافية", f"{pdf_month_profit:,.2f} ج.م")
                    m3.metric("إجمالي فواتير الشهر", f"{month_orders}")
                    m4.metric("متوسط قيمة الفاتورة", f"{avg_order_val:,.2f} ج.م")
                    
                    st.markdown("##### تفاصيل فواتير الشهر الحالي:")
                    st.dataframe(df_month.drop(columns=["التاريخ_dt"]), use_container_width=True)
                else:
                    st.info("لا توجد مبيعات مسجلة خلال هذا الشهر حتى الآن.")
            else:
                st.info("لم يتم تسجيل أي مبيعات في النظام لحساب تقارير الشهر.")

            st.markdown("---")

            st.markdown("### 💰 تقييم رأس المال والمخزون الحقيقي والأرباح")
            
            if st.session_state.sales_history:
                df_all_sales = pd.DataFrame(st.session_state.sales_history)
                total_actual_profit = float(df_all_sales["صافي الربح"].sum())
            
            if not df_stock.empty:
                df_stock["إجمالي قيمة التكلفة"] = df_stock["سعر الشراء (التكلفة)"] * df_stock["الكمية المتاحة"]
                df_stock["إجمالي قيمة البيع المتوقعة"] = df_stock["سعر البيع"] * df_stock["الكمية المتاحة"]

                pdf_stock_cost = float(df_stock["إجمالي قيمة التكلفة"].sum())
                pdf_stock_sale = float(df_stock["إجمالي قيمة البيع المتوقعة"].sum())
                expected_stock_profit = pdf_stock_sale - pdf_stock_cost

                k1, k2, k3, k4 = st.columns(4)
                k1.metric("رأس المال المستثمر بالمخزون", f"{pdf_stock_cost:,.2f} ج.م")
                k2.metric("القيمة البيعية للمخزون الحالي", f"{pdf_stock_sale:,.2f} ج.م")
                k3.metric("الأرباح المتوقعة عند بيع المخزون", f"{expected_stock_profit:,.2f} ج.م")
                k4.metric("🔥 صافي الأرباح الحقيقية للمبيعات", f"{total_actual_profit:,.2f} ج.م")
            else:
                st.metric("🔥 صافي الأرباح الحقيقية للمبيعات الفعلية", f"{total_actual_profit:,.2f} ج.م")
                st.info("المخزون فارغ حالياً.")

            st.markdown("---")

            st.markdown("### 📄 تصدير التقرير المالي")
            pdf_data = generate_pdf_report(pdf_today_rev, pdf_today_profit, pdf_month_rev, pdf_month_profit, pdf_stock_cost, pdf_stock_sale, total_actual_profit)
            
            st.download_button(
                label="📥 تحميل التقرير المالي (PDF)",
                data=pdf_data,
                file_name=f"ZAGHLOUL_Financial_Report_{datetime.now().strftime('%Y_%m_%d')}.pdf",
                mime="application/pdf",
                type="primary"
            )

            st.markdown("---")

            st.markdown("### ⚠️ جرد النواقص")
            if not df_stock.empty:
                low_stock = df_stock[df_stock["الكمية المتاحة"] <= 5]
                if not low_stock.empty:
                    st.warning("الأصناف التالية وصل رصيدها إلى 5 قطع أو أقل:")
                    st.dataframe(low_stock[["كود المنتج", "اسم المنتج", "الكمية المتاحة", "سعر البيع"]], use_container_width=True)
                else:
                    st.success("جميع الأصناف متوفرة بكميات آمنة داخل المخزون.")
            else:
                st.info("المخزون فارغ حالياً.")

    # 5. إدارة الموردين والعملاء
    if tab5 and st.session_state.role == "admin":
        with tab5:
            st.subheader("👥 إدارة الموردين والعملاء (خـاص بالمدير)")
            
            col_tab_a, col_tab_b = st.columns(2)
            
            with col_tab_a:
                st.markdown("### 📞 دليل وسجل العملاء المسجلين")
                if not df_customers.empty:
                    st.dataframe(df_customers, use_container_width=True)
                    
                    st.markdown("---")
                    st.markdown("##### 🔴 حذف عميل من السجل:")
                    cust_list = [f"{idx+1}: {row['اسم العميل']} ({row['رقم الهاتف']})" for idx, row in df_customers.iterrows()]
                    selected_cust_del = st.selectbox("اختر العميل المراد حذفه:", options=[""] + cust_list)
                    
                    if st.button("حذف العميل المحدد"):
                        if selected_cust_del:
                            cust_idx = int(selected_cust_del.split(":")[0]) - 1
                            df_customers = df_customers.drop(cust_idx).reset_index(drop=True)
                            df_customers.to_excel(CUSTOMERS_FILE, index=False)
                            st.success("تم حذف العميل بنجاح!")
                            st.rerun()
                else:
                    st.info("لا يوجد عملاء مسجلين حالياً في ملف البيانات.")
                    
            with col_tab_b:
                st.markdown("### 🏢 دليل وسجل الموردين")
                if not df_suppliers.empty:
                    st.dataframe(df_suppliers, use_container_width=True)
                    
                    st.markdown("---")
                    st.markdown("##### 🔴 حذف مورد من السجل:")
                    sup_list = [f"{idx+1}: {row['اسم المورد']} ({row['رقم الهاتف']})" for idx, row in df_suppliers.iterrows()]
                    selected_sup_del = st.selectbox("اختر المورد المراد حذفه:", options=[""] + sup_list)
                    
                    if st.button("حذف المورد المحدد"):
                        if selected_sup_del:
                            sup_idx = int(selected_sup_del.split(":")[0]) - 1
                            df_suppliers = df_suppliers.drop(sup_idx).reset_index(drop=True)
                            df_suppliers.to_excel(SUPPLIERS_FILE, index=False)
                            st.success("تم حذف المورد بنجاح!")
                            st.rerun()
                else:
                    st.info("لا يوجد موردين مسجلين حالياً.")
