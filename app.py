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

st.set_page_config(page_title="ZAGHLOUL - World Of Care", page_icon="🛍️", layout="wide")

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.user_name = ""

if "sales_history" not in st.session_state:
    st.session_state.sales_history = []

def init_empty_excel():
    columns = ["كود المنتج", "اسم المنتج", "سعر الشراء (التكلفة)", "سعر البيع", "الكمية المتاحة", "تاريخ الإضافة"]
    df_empty = pd.DataFrame(columns=columns)
    df_empty.to_excel(EXCEL_FILE, index=False)

if not os.path.exists(EXCEL_FILE):
    init_empty_excel()

def normalize_arabic(text):
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

def generate_pdf_report(today_rev, today_profit, month_rev, month_profit, stock_cost, stock_sale):
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
    except Exception:
        df_stock = pd.DataFrame(columns=["كود المنتج", "اسم المنتج", "سعر الشراء (التكلفة)", "سعر البيع", "الكمية المتاحة", "تاريخ الإضافة"])

    if not df_stock.empty and "الكمية المتاحة" in df_stock.columns:
        low_stock_items = df_stock[df_stock["الكمية المتاحة"] <= 5]
        if not low_stock_items.empty:
            st.error(f"🚨 **تنبيه هام:** يوجد ({len(low_stock_items)}) منتجات أوشكت على النفاد!")

    if st.session_state.role == "admin":
        tab1, tab2, tab3, tab4 = st.tabs([
            "🛒 شاشة المبيعات", 
            "📦 إضافة مشتريات جديدة", 
            "📋 إدارة المخزون والجرد", 
            "📊 التقارير والأرباح والجرد المالي"
        ])
    else:
        tab1, tab3 = st.tabs(["🛒 شاشة المبيعات", "📋 عرض المخزون المتاح"])
        tab2, tab4 = None, None

    # 1. شاشة المبيعات
    with tab1:
        st.subheader("تسجيل فاتورة بيع جديدة")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            customer_name = st.text_input("اسم المشتري", value="عميل نقدي")
        with col_c2:
            customer_phone = st.text_input("رقم الهاتف", value="")
            
        st.markdown("---")
        
        col_prod1, col_prod2 = st.columns(2)
        
        selected_code = ""
        selected_name = ""
        cost_price = 0.0
        sell_price = 0.0
        available_qty = 0

        with col_prod1:
            if not df_stock.empty:
                item_options = [""] + [f"{row['كود المنتج']} - {row['اسم المنتج']}" for _, row in df_stock.iterrows()]
                chosen_item = st.selectbox("اختر المنتج من المخزون:", options=item_options)
                
                if chosen_item:
                    selected_code = chosen_item.split(" - ")[0]
                    match = df_stock[df_stock["كود المنتج"] == selected_code]
                    if not match.empty:
                        idx = match.index[0]
                        selected_name = df_stock.at[idx, "اسم المنتج"]
                        cost_price = float(df_stock.at[idx, "سعر الشراء (التكلفة)"])
                        sell_price = float(df_stock.at[idx, "سعر البيع"])
                        available_qty = int(df_stock.at[idx, "الكمية المتاحة"])
            else:
                st.info("لا توجد منتجات في المخزون بعد.")

        with col_prod2:
            p_code_input = st.text_input("كود المنتج", value=selected_code, key="sale_code_in")
            p_name_input = st.text_input("اسم المنتج", value=selected_name, key="sale_name_in")

        col_qty1, col_qty2 = st.columns(2)
        with col_qty1:
            qty = st.number_input("الكمية المباعة", min_value=1, value=1)
            if available_qty > 0:
                st.caption(f"الكمية المتاحة حالياً بالمخزون: **{available_qty}** قطعة")
        with col_qty2:
            if st.session_state.role == "admin":
                sale_date = st.date_input("تاريخ الفاتورة", value=datetime.now())
            else:
                sale_date = datetime.now().date()
                st.info(f"تاريخ الفاتورة: {sale_date}")

        if st.button("💾 تسجيل عملية البيع", type="primary"):
            final_code = p_code_input.strip() if p_code_input.strip() else "بدون كود"
            final_name = p_name_input.strip() if p_name_input.strip() else "منتج غير مسمى"
            
            if final_code != "بدون كود":
                match_idx = df_stock[df_stock["كود المنتج"] == final_code].index
                if not match_idx.empty:
                    idx = match_idx[0]
                    cost_price = float(df_stock.at[idx, "سعر الشراء (التكلفة)"])
                    sell_price = float(df_stock.at[idx, "سعر البيع"])
                    
                    current_qty = int(df_stock.at[idx, "الكمية المتاحة"])
                    new_qty = max(0, current_qty - qty)
                    df_stock.at[idx, "الكمية المتاحة"] = new_qty
                    df_stock.to_excel(EXCEL_FILE, index=False)

            st.session_state.sales_history.append({
                "التاريخ": str(sale_date),
                "العميل": customer_name,
                "رقم الهاتف": customer_phone,
                "كود المنتج": final_code,
                "اسم المنتج": final_name,
                "الكمية": qty,
                "تكلفة القطعة": cost_price,
                "سعر البيع للقطعة": sell_price,
                "إجمالي البيع": qty * sell_price,
                "صافي الربح": qty * (sell_price - cost_price),
                "الموظف": st.session_state.user_name
            })
            st.success("تم تسجيل عملية البيع بنجاح!")
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
            
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                supplier_name = st.text_input("اسم المورد")
            with col_s2:
                supplier_phone = st.text_input("رقم هاتف المورد")
                
            st.markdown("---")
            col_p1, col_p2, col_p3 = st.columns(3)
            
            with col_p1:
                p_name = st.text_input("اسم المنتج بالعربي", key="new_p_name")
                
                auto_prefix = get_auto_code_prefix(p_name) if p_name else "ZAG"
                suggested_code = f"{auto_prefix}-{101 + len(df_stock)}"
                
                p_code = st.text_input("كود المنتج", value=suggested_code, key="new_p_code")

            with col_p2:
                p_cost = st.number_input("سعر التكلفة", min_value=0.0, value=0.0)
                p_price = st.number_input("سعر البيع للجمهور", min_value=0.0, value=0.0)

            with col_p3:
                p_qty = st.number_input("الكمية", min_value=0, value=1)
                p_date = st.date_input("تاريخ الشراء", value=datetime.now())

            if st.button("➕ إضافة للمخزون", type="primary"):
                final_name = p_name.strip() if p_name.strip() else "منتج جديد بدون اسم"
                final_code = p_code.strip() if p_code.strip() else suggested_code

                new_row = {
                    "كود المنتج": final_code,
                    "اسم المنتج": final_name,
                    "سعر الشراء (التكلفة)": p_cost,
                    "سعر البيع": p_price,
                    "الكمية المتاحة": p_qty,
                    "تاريخ الإضافة": str(p_date)
                }
                df_stock = pd.concat([df_stock, pd.DataFrame([new_row])], ignore_index=True)
                df_stock.to_excel(EXCEL_FILE, index=False)
                st.success(f"تمت إضافة المنتج ({final_name}) بنجاح!")
                st.rerun()

    # 3. إدارة المخزون والجرد
    with tab3:
        st.subheader("📋 جدول جرد المخزون")
        
        if not df_stock.empty:
            if st.session_state.role == "cashier":
                cols_to_show = [c for c in df_stock.columns if "التكلفة" not in c and "الشراء" not in c]
                st.dataframe(df_stock[cols_to_show], use_container_width=True)
            else:
                st.dataframe(df_stock, use_container_width=True)

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

            st.markdown("### 💰 تقييم رأس المال والمخزون الحقيقي")
            if not df_stock.empty:
                df_stock["إجمالي قيمة التكلفة"] = df_stock["سعر الشراء (التكلفة)"] * df_stock["الكمية المتاحة"]
                df_stock["إجمالي قيمة البيع المتوقعة"] = df_stock["سعر البيع"] * df_stock["الكمية المتاحة"]
                df_stock["الربح المتوقع"] = df_stock["إجمالي قيمة البيع المتوقعة"] - df_stock["إجمالي قيمة التكلفة"]

                pdf_stock_cost = float(df_stock["إجمالي قيمة التكلفة"].sum())
                pdf_stock_sale = float(df_stock["إجمالي قيمة البيع المتوقعة"].sum())
                expected_stock_profit = pdf_stock_sale - pdf_stock_cost

                k1, k2, k3 = st.columns(3)
                k1.metric("رأس المال المستثمر", f"{pdf_stock_cost:,.2f} ج.م")
                k2.metric("القيمة البيعية الكلية للمخزون", f"{pdf_stock_sale:,.2f} ج.م")
                k3.metric("الأرباح المتوقعة عند بيع المخزون", f"{expected_stock_profit:,.2f} ج.م")
            else:
                st.info("المخزون فارغ حالياً.")

            st.markdown("---")

            st.markdown("### 📄 تصدير التقرير المالي")
            pdf_data = generate_pdf_report(pdf_today_rev, pdf_today_profit, pdf_month_rev, pdf_month_profit, pdf_stock_cost, pdf_stock_sale)
            
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
