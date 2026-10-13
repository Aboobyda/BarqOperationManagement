import io
import os
import re
import sqlite3
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

# 1. إعداد الصفحة والـ CSS للاتجاه العربي RTL
st.set_page_config(
    page_title="نظام إدارة حركة أسطول الرحلات",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        html, body, [class*="css"] {
            direction: rtl;
            text-align: right;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        
        .metric-card {
            background: linear-gradient(135deg, #1f4e78 0%, #2c3e50 100%);
            padding: 18px 20px;
            border-radius: 12px;
            color: white !important;
            box-shadow: 0 4px 10px rgba(0,0,0,0.08);
            text-align: center;
            margin-bottom: 15px;
            border-right: 6px solid #00b4d8;
        }
        .metric-card-green {
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            border-right: 6px solid #10b981;
        }
        .metric-card-orange {
            background: linear-gradient(135deg, #f857a6 0%, #ff5858 100%);
            border-right: 6px solid #ef4444;
        }
        .metric-card-purple {
            background: linear-gradient(135deg, #6a11cb 0%, #2575fc 100%);
            border-right: 6px solid #6366f1;
        }
        .metric-title {
            font-size: 15px;
            font-weight: bold;
            opacity: 0.95;
            margin-bottom: 8px;
        }
        .metric-value {
            font-size: 26px;
            font-weight: 800;
            letter-spacing: 0.5px;
        }

        .stDataFrame, div[data-testid="stDataEditor"] { 
            direction: rtl !important; 
            width: 100% !important; 
        }

        div[data-testid="stDataEditor"] div[role="columnheader"],
        .stDataFrame div[role="columnheader"] {
            white-space: nowrap !important;
            direction: rtl !important;
            text-align: right !important;
        }

        @media print {
            [data-testid="stSidebar"], header, footer, button, .no-print, [data-testid="stHeader"] {
                display: none !important;
            }
            body, .main, .block-container {
                direction: rtl !important;
                text-align: right !important;
                padding: 0 !important;
                margin: 0 !important;
                width: 100% !important;
            }
        }
    </style>
""",
    unsafe_allow_html=True,
)

DB_PATH = "trips_db.db"

ARABIC_MONTHS = {
    1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل",
    5: "مايو", 6: "يونيو", 7: "يوليو", 8: "أغسطس",
    9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر",
}

ARABIC_DAYS = {
    "Saturday": "السبت", "Sunday": "الأحد", "Monday": "الإثنين",
    "Tuesday": "الثلاثاء", "Wednesday": "الأربعاء", "Thursday": "الخميس", "Friday": "الجمعة",
}

NUMBER_TO_ARABIC_ORDINAL = {
    1: "الأولى", 2: "الثانية", 3: "الثالثة", 4: "الرابعة", 5: "الخامسة",
    6: "السادسة", 7: "السابعة", 8: "الثامنة", 9: "التاسعة", 10: "العاشرة",
}

EXCLUDED_COMPANIES = ["برق الشمال", "الكرتون", "تفريق حاويات"]

DRIVER_NAME_MAP = {
    "ابوبكر": "أبو بكر",
    "أبو بكر": "أبو بكر",
}

# دالة حازمة لتطهير النصوص من أي رموز ASCII الخفية التي يرفضها openpyxl
def sanitize_text(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, bytes):
        try:
            val = val.decode("utf-8")
        except Exception:
            try:
                val = val.decode("windows-1256")
            except Exception:
                val = val.decode("utf-8", errors="ignore")
    
    text = str(val).strip()
    # حذف رموز التحكم غير الصالحة في XML / Excel
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]', '', text)
    return text


def get_db_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.text_factory = lambda x: x.decode('utf-8', 'ignore') if isinstance(x, bytes) else x
    return conn


# 2. تحميل البيانات
@st.cache_data(ttl=60)
def load_data():
    if not os.path.exists(DB_PATH):
        st.error("⚠️ لم يتم العثور على قاعدة البيانات! يرجى تشغيل سكريبت migrate.py أولاً.")
        return pd.DataFrame(), [], [], [], [], [], []

    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM trips ORDER BY id ASC", conn)
    
    df = df.rename(columns={
        "id": "الرقم",
        "company": "اسم الشركة",
        "day": "اليوم",
        "date": "التاريخ",
        "waste_type": "المخلفه",
        "driver": "السائق",
        "plate": "رقم اللوحة",
        "quantity": "كمية المخلفات",
        "equipment": "المعدة",
        "amount": "سعر الرحلة",
        "driver_fee": "مبلغ السائق",
        "receipt": "السند",
        "station": "محطة الاتلاف",
        "notes": "الملاحظات",
        "month": "الشهر"
    })

    df["سعر_الرحلة_رقم"] = pd.to_numeric(df["سعر الرحلة"], errors="coerce").fillna(0)
    df["مبلغ_السائق_رقم"] = pd.to_numeric(df["مبلغ السائق"], errors="coerce").fillna(15)

    string_cols = ["اسم الشركة", "اليوم", "التاريخ", "المخلفه", "السائق", "رقم اللوحة", 
                   "كمية المخلفات", "المعدة", "سعر الرحلة", "مبلغ السائق", "السند", "محطة الاتلاف", "الملاحظات", "الشهر"]
    for col in string_cols:
        if col in df.columns:
            df[col] = df[col].apply(sanitize_text)

    def fetch_list_merged(table_name, df_col_name):
        res_db = pd.read_sql_query(f"SELECT name FROM {table_name}", conn)
        db_items = [sanitize_text(x) for x in res_db["name"].dropna().tolist() if sanitize_text(x)]
        df_items = [sanitize_text(x) for x in df[df_col_name].dropna().unique().tolist() if sanitize_text(x)]
        all_items = sorted(list(set([x for x in db_items + df_items if x not in ["nan", "الكل", ""]])))
        return all_items

    companies_list = fetch_list_merged("companies", "اسم الشركة")
    wastes_list = fetch_list_merged("wastes", "المخلفه")
    drivers_list = fetch_list_merged("drivers", "السائق")
    equipments_list = fetch_list_merged("equipments", "المعدة")
    stations_list = fetch_list_merged("stations", "محطة الاتلاف")
    plates_list = fetch_list_merged("plates", "رقم اللوحة")

    conn.close()
    return df, companies_list, wastes_list, drivers_list, equipments_list, stations_list, plates_list


(df, companies_list, wastes_list, drivers_list, equipments_list, stations_list, plates_list) = load_data()

if df.empty and not os.path.exists(DB_PATH):
    st.stop()

ALL_DISPLAY_COLUMNS = [
    "الرقم", "اسم الشركة", "اليوم", "التاريخ", "المخلفه", "السائق", "رقم اللوحة",
    "كمية المخلفات", "المعدة", "سعر الرحلة", "مبلغ السائق", "السند", "محطة الاتلاف", "الملاحظات"
]
VALID_ALL_COLS = [c for c in ALL_DISPLAY_COLUMNS if c in df.columns]


# دالة التصدير إلى Excel المعالجة لحظر IllegalCharacterError
def convert_df_to_excel(df_to_export, cols_order):
    output = io.BytesIO()
    export_df = df_to_export.copy()
    valid_cols = [c for c in cols_order if c in export_df.columns]
    export_df = export_df[valid_cols]

    # تنظيف جميع الخلايا النصية قبل كتابتها بالملف
    for col in export_df.columns:
        export_df[col] = export_df[col].apply(sanitize_text)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="الرحلات")
    return output.getvalue()


def generate_custom_print_html(data_df, chosen_cols, title_text):
    base_cols = [c for c in ALL_DISPLAY_COLUMNS if c in chosen_cols]

    html_cols = "".join([f"<th style='padding:8px; border:1px solid #000; background-color:#1F4E78; color:white; font-size:14px;'>{col}</th>" for col in base_cols])
    rows_html = ""
    for _, row in data_df.iterrows():
        rows_html += "<tr>" + "".join([f"<td style='padding:6px; border:1px solid #000; text-align:center; font-size:13px;'>{sanitize_text(row[col]) if col in row and sanitize_text(row[col]) else '-'}</td>" for col in base_cols]) + "</tr>"

    return f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="utf-8"><title>{title_text}</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, sans-serif; direction: rtl; text-align: right; padding: 15px; }}
            h2 {{ text-align: center; color: #1F4E78; margin-bottom: 15px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th, td {{ border: 1px solid #000; padding: 6px; text-align: center; }}
            @media print {{ .no-print {{ display: none !important; }} }}
        </style>
    </head>
    <body>
        <button class="no-print" onclick="window.print()" style="padding:10px 22px; font-size:16px; background-color:#28a745; color:white; border:none; border-radius:5px; cursor:pointer; margin-bottom:15px;">🖨️ اضغط هنا لتأكيد طباعة الحقول المحددة / PDF</button>
        <h2>{title_text}</h2>
        <table><thead><tr>{html_cols}</tr></thead><tbody>{rows_html}</tbody></table>
        <script>window.onload = function() {{ window.print(); }}</script>
    </body>
    </html>"""


st.sidebar.title("🚚 إدارة أسطول الرحلات")
st.sidebar.markdown("---")

def clear_active_actions():
    st.session_state["edit_trip_id"] = None
    st.session_state["confirm_delete_id"] = None
    st.session_state["confirm_del_item"] = None
    st.session_state["editor_key_version"] = st.session_state.get("editor_key_version", 0) + 1

menu = st.sidebar.radio(
    "القائمة الرئيسية:",
    [
        "📊 لوحة التحكم الإحصائية",
        "➕ إضافة رحلة جديدة",
        "📄 التقرير اليومي التلقائي",
        "🏢 تقارير الشركات",
        "🚚 تقارير السائقين",
        "📂 قاعدة البيانات الموحدة",
        "⚙️ إدارة القوائم (إضافة/حذف)",
    ],
    on_change=clear_active_actions
)

# ---------------------------------------------------------
# 1. لوحة التحكم الإحصائية
# ---------------------------------------------------------
if menu == "📊 لوحة التحكم الإحصائية":
    st.title("📊 لوحة التحكم الشاملة لعمليات أسطول الرحلات")
    st.markdown("---")

    avail_months_dash = ["الكل"] + [m for m in list(ARABIC_MONTHS.values()) if m in df["الشهر"].dropna().unique()]
    selected_dash_month = st.selectbox("🗓️ تصفية البيانات حسب الشهر:", avail_months_dash, key="dash_month_select")

    dash_df = df if selected_dash_month == "الكل" else df[df["الشهر"] == selected_dash_month]

    st.markdown("---")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">🚚 إجمالي عدد الرحلات</div>
                <div class="metric-value">{len(dash_df):,}</div>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
            <div class="metric-card metric-card-green">
                <div class="metric-title">💰 إجمالي إيرادات الرحلات</div>
                <div class="metric-value">{dash_df['سعر_الرحلة_رقم'].sum():,.2f} ريال</div>
            </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
            <div class="metric-card metric-card-orange">
                <div class="metric-title">💳 إجمالي مستحقات السائقين</div>
                <div class="metric-value">{dash_df['مبلغ_السائق_رقم'].sum():,.2f} ريال</div>
            </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
            <div class="metric-card metric-card-purple">
                <div class="metric-title">🏢 الشركات / السائقين</div>
                <div class="metric-value">{dash_df['اسم الشركة'].nunique()} / {dash_df['السائق'].nunique()}</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    r1_c1, r1_c2 = st.columns(2)
    with r1_c1:
        st.subheader("📅 عدد الرحلات حسب الأشهر")
        monthly = dash_df["الشهر"].value_counts().reset_index()
        monthly.columns = ["الشهر", "عدد الرحلات"]
        fig_m = px.bar(
            monthly, 
            x="الشهر", 
            y="عدد الرحلات", 
            text_auto=True, 
            color="عدد الرحلات", 
            color_continuous_scale="Blues",
            labels={"الشهر": "اسم الشهر", "عدد الرحلات": "عدد الرحلات"}
        )
        fig_m.update_xaxes(type="category")
        st.plotly_chart(fig_m, use_container_width=True)

    with r1_c2:
        st.subheader("♻️ توزيع أنواع المخلفات")
        waste = dash_df["المخلفه"].value_counts().reset_index()
        waste.columns = ["نوع المخلفة", "العدد"]
        fig_w = px.pie(waste, names="نوع المخلفة", values="العدد", hole=0.3)
        st.plotly_chart(fig_w, use_container_width=True)

    r2_c1, r2_c2 = st.columns(2)
    with r2_c1:
        st.subheader("🏢 توزيع جميع الشركات طلباً للرحلات")
        comp = dash_df["اسم الشركة"].value_counts().reset_index()
        comp.columns = ["الشركة", "عدد الرحلات"]
        fig_c = px.bar(
            comp, 
            x="الشركة", 
            y="عدد الرحلات", 
            text_auto=True, 
            color="عدد الرحلات", 
            color_continuous_scale="Viridis",
            labels={"الشركة": "اسم الشركة", "عدد الرحلات": "عدد الرحلات"}
        )
        fig_c.update_xaxes(categoryorder="total descending")
        st.plotly_chart(fig_c, use_container_width=True)

    with r2_c2:
        st.subheader("👨‍✈️ أكثر السائقين تنفيذ للرحلات")
        driv = dash_df["السائق"].value_counts().head(10).reset_index()
        driv.columns = ["السائق", "عدد الرحلات"]
        fig_d = px.bar(driv, x="السائق", y="عدد الرحلات", text_auto=True, color="عدد الرحلات", color_continuous_scale="Teal")
        st.plotly_chart(fig_d, use_container_width=True)

# ---------------------------------------------------------
# 2. إضافة رحلة جديدة
# ---------------------------------------------------------
elif menu == "➕ إضافة رحلة جديدة":
    st.title("➕ تسجيل رحلة جديدة في قاعدة البيانات")
    st.markdown("---")

    next_id = int(pd.to_numeric(df["الرقم"], errors="coerce").max() + 1) if not df.empty else 1

    col_date1, col_date2 = st.columns(2)
    with col_date1:
        date_input = st.date_input("📅 التاريخ (تلقائي تاريخ اليوم):", datetime.now(), key="trip_date")

    day_name_en = date_input.strftime("%A")
    calculated_day = ARABIC_DAYS.get(day_name_en, "غير محدد")

    with col_date2:
        st.text_input("📆 اليوم (يتغير تلقائياً):", value=calculated_day, disabled=True)

    opt_companies = ["اختر اسم الشركة..."] + companies_list
    opt_drivers = ["اختر اسم السائق..."] + drivers_list
    opt_wastes = ["اختر نوع المخلفة..."] + wastes_list
    opt_plates = ["اختر رقم اللوحة..."] + plates_list
    opt_equipments = ["اختر نوع المعدة..."] + equipments_list
    opt_stations = ["اختر محطة الاتلاف..."] + stations_list

    with st.form("add_trip_main_form"):
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.number_input("الرقم (تلقائي):", value=next_id, disabled=True)
            company_input = st.selectbox("اسم الشركة *", opt_companies, index=0)
        with col_b:
            driver_input = st.selectbox("اسم السائق *", opt_drivers, index=0)
            plate_input = st.selectbox("رقم اللوحة / السيارة *", opt_plates, index=0)
        with col_c:
            waste_input = st.selectbox("نوع المخلفة *", opt_wastes, index=0)
            equipment_input = st.selectbox("المعدة *", opt_equipments, index=0)

        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            receipt_input = st.text_input("رقم السند * (إجباري)", placeholder="ادخل رقم السند هنا...")
        with col_r2:
            quantity_input = st.text_input("كمية المخلفات:", placeholder="مثال: 7 طن أو 30 يارده")
        with col_r3:
            amount_input = st.number_input("سعر الرحلة (ريال):", min_value=0, value=400, step=50)

        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            driver_fee_input = st.number_input("مبلغ السائق (ريال):", min_value=0, value=15, step=5)
        with col_s2:
            station_input = st.selectbox("محطة الاتلاف / المردم *", opt_stations, index=0)
        with col_s3:
            notes_input = st.text_input("الملاحظات:", placeholder="أي ملاحظات إضافية...")

        submit_trip = st.form_submit_button("💾 حفظ الرحلة الجديدة")

        if submit_trip:
            missing_fields = []
            if company_input == "اختر اسم الشركة...": missing_fields.append("اسم الشركة")
            if driver_input == "اختر اسم السائق...": missing_fields.append("اسم السائق")
            if waste_input == "اختر نوع المخلفة...": missing_fields.append("نوع المخلفة")
            if plate_input == "اختر رقم اللوحة...": missing_fields.append("رقم اللوحة")
            if equipment_input == "اختر نوع المعدة...": missing_fields.append("المعدة")
            if station_input == "اختر محطة الاتلاف...": missing_fields.append("محطة الاتلاف")
            if not receipt_input or not receipt_input.strip(): missing_fields.append("رقم السند (إجباري)")

            if missing_fields:
                st.error(f"⚠️ **تنبيه:** تعذر حفظ الرحلة! يرجى إكمال البيانات التالية: ({' - '.join(missing_fields)})")
            else:
                conn = get_db_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO trips (id, company, day, date, waste_type, driver, plate, quantity, equipment, amount, driver_fee, receipt, station, notes, month)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    next_id, sanitize_text(company_input), sanitize_text(calculated_day), date_input.strftime("%d/%m/%Y"),
                    sanitize_text(waste_input), sanitize_text(DRIVER_NAME_MAP.get(driver_input, driver_input)), sanitize_text(plate_input),
                    sanitize_text(quantity_input), sanitize_text(equipment_input), amount_input, driver_fee_input,
                    sanitize_text(receipt_input), sanitize_text(station_input),
                    sanitize_text(notes_input) if notes_input and sanitize_text(notes_input) else None,
                    ARABIC_MONTHS[date_input.month]
                ))
                conn.commit()
                conn.close()

                st.toast("✅ تم إضافة الرحلة وحفظها بنجاح!", icon="🎉")
                st.cache_data.clear()
                st.rerun()

# ---------------------------------------------------------
# 3. التقرير اليومي التلقائي
# ---------------------------------------------------------
elif menu == "📄 التقرير اليومي التلقائي":
    st.title("📄 مولد تقرير العمليات اليومي التلقائي")
    st.markdown("---")

    rep_date_input = st.date_input("📅 اختر تاريخ التقرير:", datetime.now(), key="daily_rep_date")
    target_date_str = rep_date_input.strftime("%d/%m/%Y")
    arabic_day = ARABIC_DAYS.get(rep_date_input.strftime("%A"), "غير محدد")

    daily_trips = df[(df["التاريخ"] == target_date_str) & (~df["اسم الشركة"].astype(str).str.strip().isin(EXCLUDED_COMPANIES))].copy()
    total_daily_trips = len(daily_trips)

    if rep_date_input == datetime.now().date():
        now_time = datetime.now().strftime("%I:%M").lstrip("0")
        period = "مساءً" if datetime.now().hour >= 12 else "صباحاً"
        time_text = f" (تمام الساعة {now_time} {period})"
    else:
        time_text = ""

    st.info(f"📊 **عدد الرحلات المشمولة:** {total_daily_trips} رحلة")

    if st.button("🚀 توليد التقرير اليومي الفوري", type="primary", use_container_width=True):
        operations_text = ""
        for idx, (_, row) in enumerate(daily_trips.iterrows(), 1):
            ordinal_name = NUMBER_TO_ARABIC_ORDINAL.get(idx, f"الـ {idx}")
            details = f"عملية نقل/استبدال ({row.get('المعدة', 'حاوية')})"
            if pd.notna(row.get("كمية المخلفات")) and sanitize_text(row.get("كمية المخلفات")):
                details += f" بمقاس/كمية ({row.get('كمية المخلفات')})"
            details += f" لنفايات ({row.get('المخلفه', 'نفايات')})"
            if pd.notna(row.get("الملاحظات")) and sanitize_text(row.get("الملاحظات")):
                details += f" - {row.get('الملاحظات')}"

            operations_text += f" العملية {ordinal_name}    ©\n   وهي عبارة عن {details} لصالح شركة ({row.get('اسم الشركة')})، وقد نفذت العملية بواسطة السائق ({row.get('السائق')}).\n\n"

        report_template = f"""بسم الله الرحمن الرحيم

السادة بشركة برق الشمال للخدمات البيئية..... 
السلام عليكم ورحمة الله تعالى وبركاته، أسعد الله أمسيتكم وطابت أوقاتكم ودمتم آمنين في حلكم وترحالكم....

أما بعد..  
نفيد سيادتكم عبر هذه النافذة لنعرب عن تقرير عملياتنا المنجزة خلال هذا اليوم وبهذا الإشعار الذي يوافق يوم {arabic_day} / {target_date_str} م{time_text} - والتي أُجريت خلال هذا اليوم وبفضل الله قد نفذنا عدد ({total_daily_trips}) عمليات وكانت على النحو التالي: 👇

{operations_text}= = =
وهذا توضيح منا لمجريات عملياتنا لهذا اليوم، ونأمل بأن نكون قد قدمنا في هذا التقرير معلومات دقيقة وموثقة لمجريات سير العمليات بمقر ومباني شركة برق الشمال للخدمات البيئية.

© ودمتم في أمان الله وحفظه ©

مع خالص التحايا... 
عمليات شركة برق الشمال للخدمات البيئية...."""
        st.code(report_template, language="text")

# ---------------------------------------------------------
# 4. تقارير الشركات الشهري
# ---------------------------------------------------------
elif menu == "🏢 تقارير الشركات":
    st.title("🏢 تقرير رحلات وعمليات الشركات الشهري")
    st.markdown("---")

    avail_months = ["الكل"] + [m for m in list(ARABIC_MONTHS.values()) if m in df["الشهر"].dropna().unique()]
    avail_companies = ["الكل"] + companies_list

    c_f1, c_f2 = st.columns(2)
    with c_f1:
        selected_company = st.selectbox("اختر اسم الشركة:", avail_companies)
    with c_f2:
        selected_month = st.selectbox("اختر الشهر:", avail_months)

    filtered_df = df.copy()
    if selected_company != "الكل":
        filtered_df = filtered_df[filtered_df["اسم الشركة"] == selected_company]
    if selected_month != "الكل":
        filtered_df = filtered_df[filtered_df["الشهر"] == selected_month]

    filtered_df = filtered_df.reset_index(drop=True)
    filtered_df["الرقم"] = range(1, len(filtered_df) + 1)

    st.markdown("---")
    m1, m2 = st.columns(2)
    with m1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">📌 إجمالي عدد الرحلات</div>
                <div class="metric-value">{len(filtered_df):,} رحلة</div>
            </div>
        """, unsafe_allow_html=True)
    with m2:
        st.markdown(f"""
            <div class="metric-card metric-card-green">
                <div class="metric-title">💰 إجمالي إيرادات الرحلات</div>
                <div class="metric-value">{filtered_df['سعر_الرحلة_رقم'].sum():,.2f} ريال</div>
            </div>
        """, unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("🎯 **تخصيص الحقول المطلوبة في التقرير والطباعة والتصدير:**")
    selected_cols = st.multiselect(
        "اختر الحقول التي ترغب بإظهارها وطباعتها وتصديرها فقط:",
        options=VALID_ALL_COLS,
        default=VALID_ALL_COLS,
    )

    if not selected_cols:
        st.warning("يرجى اختيار حقل واحد على الأقل للعرض والطباعة والتصدير.")
    else:
        ordered_selected = [c for c in ALL_DISPLAY_COLUMNS if c in selected_cols]
        st_display_cols = list(reversed([c for c in ALL_DISPLAY_COLUMNS if c in filtered_df.columns]))

        st.dataframe(
            filtered_df[st_display_cols],
            use_container_width=True,
            hide_index=True,
        )

        col_p1, col_p2 = st.columns(2)
        with col_p1:
            btn_print_comp = st.button("🖨️ طباعة التقرير / حفظ كـ PDF", use_container_width=True)
            if btn_print_comp:
                report_title = f"تقرير رحلات شركة: ({selected_company}) - شهر: ({selected_month})"
                print_code = generate_custom_print_html(filtered_df, ordered_selected, report_title)
                st.components.v1.html(print_code, height=600, scrolling=True)

        with col_p2:
            excel_data = convert_df_to_excel(filtered_df, ordered_selected)
            st.download_button(
                label="📥 تصدير التقرير الحالي إلى Excel",
                data=excel_data,
                file_name=f"Company_Report_{selected_company}_{selected_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ---------------------------------------------------------
# 5. تقارير السائقين الشهري
# ---------------------------------------------------------
elif menu == "🚚 تقارير السائقين":
    st.title("🚚 تقرير رحلات حركة السائقين الشهري")
    st.markdown("---")

    avail_months = ["الكل"] + [m for m in list(ARABIC_MONTHS.values()) if m in df["الشهر"].dropna().unique()]
    avail_drivers = ["الكل"] + drivers_list

    d_f1, d_f2 = st.columns(2)
    with d_f1:
        selected_driver = st.selectbox("اختر اسم السائق:", avail_drivers)
    with d_f2:
        selected_month_d = st.selectbox("اختر الشهر:", avail_months)

    filtered_driv = df.copy()
    if selected_driver != "الكل":
        filtered_driv = filtered_driv[filtered_driv["السائق"] == selected_driver]
    if selected_month_d != "الكل":
        filtered_driv = filtered_driv[filtered_driv["الشهر"] == selected_month_d]

    filtered_driv = filtered_driv.reset_index(drop=True)
    filtered_driv["الرقم"] = range(1, len(filtered_driv) + 1)

    st.markdown("---")
    s1, s2, s3 = st.columns(3)
    with s1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">🚚 إجمالي عدد رحلات السائق</div>
                <div class="metric-value">{len(filtered_driv):,} رحلة</div>
            </div>
        """, unsafe_allow_html=True)
    with s2:
        st.markdown(f"""
            <div class="metric-card metric-card-green">
                <div class="metric-title">💵 إجمالي إيرادات الرحلات</div>
                <div class="metric-value">{filtered_driv['سعر_الرحلة_رقم'].sum():,.2f} ريال</div>
            </div>
        """, unsafe_allow_html=True)
    with s3:
        st.markdown(f"""
            <div class="metric-card metric-card-orange">
                <div class="metric-title">💳 إجمالي مستحقات السائق</div>
                <div class="metric-value">{filtered_driv['مبلغ_السائق_رقم'].sum():,.2f} ريال</div>
            </div>
        """, unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("🎯 **تخصيص الحقول المطلوبة في التقرير والطباعة والتصدير:**")
    selected_cols_d = st.multiselect(
        "اختر الحقول التي ترغب بإظهارها وطباعتها وتصديرها فقط:",
        options=VALID_ALL_COLS,
        default=VALID_ALL_COLS,
        key="ms_driver",
    )

    if not selected_cols_d:
        st.warning("يرجى اختيار حقل واحد على الأقل للعرض والطباعة والتصدير.")
    else:
        ordered_selected_d = [c for c in ALL_DISPLAY_COLUMNS if c in selected_cols_d]
        st_display_cols_d = list(reversed([c for c in ALL_DISPLAY_COLUMNS if c in filtered_driv.columns]))

        st.dataframe(
            filtered_driv[st_display_cols_d],
            use_container_width=True,
            hide_index=True,
        )

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            btn_print_driver = st.button("🖨️ طباعة تقرير السائق / حفظ كـ PDF", use_container_width=True)
            if btn_print_driver:
                report_title = f"تقرير حركة السائق: ({selected_driver}) - شهر: ({selected_month_d})"
                print_code = generate_custom_print_html(filtered_driv, ordered_selected_d, report_title)
                st.components.v1.html(print_code, height=600, scrolling=True)

        with col_d2:
            excel_data_d = convert_df_to_excel(filtered_driv, ordered_selected_d)
            st.download_button(
                label="📥 تصدير تقرير السائق الحالي إلى Excel",
                data=excel_data_d,
                file_name=f"Driver_Report_{selected_driver}_{selected_month_d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ---------------------------------------------------------
# 6. قاعدة البيانات الموحدة
# ---------------------------------------------------------
elif menu == "📂 قاعدة البيانات الموحدة":
    st.title("📂 قاعدة البيانات الموحدة لجميع الرحلات")
    st.markdown("---")

    col_s_box, col_s_btn = st.columns([3, 1])
    with col_s_box:
        search_term = st.text_input("🔍 بحث سريع (شركة، سائق، لوحة...):")
    
    df_db_show = df.copy()

    if search_term and search_term.strip():
        mask = df_db_show[VALID_ALL_COLS].astype(str).apply(lambda row: row.str.contains(search_term, case=False)).any(axis=1)
        df_db_show = df_db_show[mask]

    with col_s_btn:
        st.write("")
        excel_db = convert_df_to_excel(df_db_show, ALL_DISPLAY_COLUMNS)
        st.download_button(
            label="📥 تصدير البيانات المعروضة إلى Excel",
            data=excel_db,
            file_name=f"Trips_Export_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    is_editing = ("edit_trip_id" in st.session_state and st.session_state["edit_trip_id"] is not None)
    is_deleting = ("confirm_delete_id" in st.session_state and st.session_state["confirm_delete_id"] is not None)

    if is_editing:
        trip_id = st.session_state["edit_trip_id"]
        target_rows = df[df["الرقم"].astype(str) == str(trip_id)]

        if not target_rows.empty:
            row_data = target_rows.iloc[0]
            st.warning(f"✏️ **تعديل بيانات الرحلة رقم ({trip_id}) لشركة: {row_data['اسم الشركة']}**")

            curr_date_val = datetime.now()
            if pd.notna(row_data.get("التاريخ")) and sanitize_text(row_data.get("التاريخ")):
                try:
                    curr_date_val = datetime.strptime(sanitize_text(row_data["التاريخ"]), "%d/%m/%Y")
                except ValueError:
                    pass

            with st.form(f"edit_form_{trip_id}"):
                e_date_col1, e_date_col2 = st.columns(2)
                with e_date_col1:
                    e_date_input = st.date_input("📅 التاريخ:", value=curr_date_val)
                
                e_day_name_en = e_date_input.strftime("%A")
                e_calculated_day = ARABIC_DAYS.get(e_day_name_en, "غير محدد")
                
                with e_date_col2:
                    st.text_input("📆 اليوم (يتغير تلقائياً):", value=e_calculated_day, disabled=True)

                e_col1, e_col2, e_col3 = st.columns(3)
                with e_col1:
                    e_comp = st.selectbox("اسم الشركة:", companies_list, index=companies_list.index(sanitize_text(row_data["اسم الشركة"])) if sanitize_text(row_data["اسم الشركة"]) in companies_list else 0)
                    e_driver = st.selectbox("السائق:", drivers_list, index=drivers_list.index(sanitize_text(row_data["السائق"])) if sanitize_text(row_data["السائق"]) in drivers_list else 0)
                    e_waste = st.selectbox("نوع المخلفة:", wastes_list, index=wastes_list.index(sanitize_text(row_data["المخلفه"])) if sanitize_text(row_data["المخلفه"]) in wastes_list else 0)
                with e_col2:
                    e_plate = st.selectbox("رقم اللوحة:", plates_list, index=plates_list.index(sanitize_text(row_data["رقم اللوحة"])) if sanitize_text(row_data["رقم اللوحة"]) in plates_list else 0)
                    e_qty = st.text_input("الكمية:", value=sanitize_text(row_data["كمية المخلفات"]))
                    e_equip = st.selectbox("المعدة:", equipments_list, index=equipments_list.index(sanitize_text(row_data["المعدة"])) if sanitize_text(row_data["المعدة"]) in equipments_list else 0)
                with e_col3:
                    e_amount = st.text_input("سعر الرحلة:", value=sanitize_text(row_data["سعر الرحلة"]))
                    e_driver_fee = st.text_input("مبلغ السائق:", value=sanitize_text(row_data["مبلغ السائق"]) if sanitize_text(row_data["مبلغ السائق"]) else "15")
                    e_receipt = st.text_input("السند:", value=sanitize_text(row_data["السند"]))
                    e_station = st.selectbox("محطة الاتلاف:", stations_list, index=stations_list.index(sanitize_text(row_data["محطة الاتلاف"])) if sanitize_text(row_data["محطة الاتلاف"]) in stations_list else 0)

                e_notes = st.text_input("الملاحظات:", value=sanitize_text(row_data["الملاحظات"]))

                b_save, b_cancel = st.columns(2)
                submit_edit = b_save.form_submit_button("💾 حفظ التعديلات")
                cancel_edit = b_cancel.form_submit_button("❌ إلغاء والتراجع")

                if submit_edit:
                    formatted_e_date = e_date_input.strftime("%d/%m/%Y")
                    extracted_e_month = ARABIC_MONTHS[e_date_input.month]

                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute("""
                        UPDATE trips SET 
                        company=?, day=?, date=?, waste_type=?, driver=?, plate=?, quantity=?, equipment=?, amount=?, driver_fee=?, receipt=?, station=?, notes=?, month=?
                        WHERE id=?
                    """, (
                        sanitize_text(e_comp), sanitize_text(e_calculated_day), formatted_e_date, sanitize_text(e_waste), sanitize_text(DRIVER_NAME_MAP.get(e_driver, e_driver)),
                        sanitize_text(e_plate), sanitize_text(e_qty), sanitize_text(e_equip), sanitize_text(e_amount), sanitize_text(e_driver_fee), sanitize_text(e_receipt), sanitize_text(e_station),
                        sanitize_text(e_notes) if e_notes else None, extracted_e_month, trip_id
                    ))
                    conn.commit()
                    conn.close()

                    clear_active_actions()
                    st.toast("✅ تم حفظ التعديلات بنجاح!", icon="💾")
                    st.cache_data.clear()
                    st.rerun()

                if cancel_edit:
                    clear_active_actions()
                    st.rerun()
        else:
            clear_active_actions()

    elif is_deleting:
        del_id = st.session_state["confirm_delete_id"]
        st.error(f"🚨 **تنبيه:** هل أنت أصلًا برغبتك متأكد من حذف الرحلة رقم ({del_id}) نهائياً؟")
        col_yes, col_no = st.columns(2)
        if col_yes.button("✅ نعم، تأكيد الحذف النهائي"):
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM trips WHERE id=?", (del_id,))
            conn.commit()
            conn.close()

            clear_active_actions()
            st.toast(f"✅ تم حذف الرحلة رقم ({del_id}) بنجاح!", icon="🗑️")
            st.cache_data.clear()
            st.rerun()
        
        if col_no.button("❌ إلغاء والتراجع"):
            clear_active_actions()
            st.rerun()

    if not is_editing and not is_deleting:
        df_db_show["✏️ تعديل"] = False
        df_db_show["🗑️ حذف"] = False

        st_display_db = ["✏️ تعديل", "🗑️ حذف"] + list(reversed(VALID_ALL_COLS))

        editor_key = f"data_editor_db_{st.session_state.get('editor_key_version', 0)}"

        edited_db = st.data_editor(
            df_db_show[st_display_db],
            use_container_width=True,
            hide_index=True,
            key=editor_key,
            column_config={
                "✏️ تعديل": st.column_config.CheckboxColumn("✏️ تعديل", default=False),
                "🗑️ حذف": st.column_config.CheckboxColumn("🗑️ حذف", default=False),
            },
            disabled=VALID_ALL_COLS
        )

        edited_rows = edited_db[edited_db["✏️ تعديل"] == True]
        deleted_rows = edited_db[edited_db["🗑️ حذف"] == True]

        if not edited_rows.empty:
            st.session_state["edit_trip_id"] = str(edited_rows.iloc[0]["الرقم"])
            st.rerun()

        if not deleted_rows.empty:
            st.session_state["confirm_delete_id"] = str(deleted_rows.iloc[0]["الرقم"])
            st.rerun()

# ---------------------------------------------------------
# 7. إدارة القوائم المنسدلة
# ---------------------------------------------------------
elif menu == "⚙️ إدارة القوائم (إضافة/حذف)":
    st.title("⚙️ إدارة القوائم المنسدلة (شركة / نفايات / سائق / معدة / مردم / سيارة)")
    st.markdown("---")

    tab_add, tab_del = st.tabs(["➕ إضافة عنصر جديد", "🗑️ حذف عنصر من القوائم"])

    with tab_add:
        c1, c2, c3, c4, c5, c6 = st.columns(6)

        def add_item_to_db(table_name, item_value):
            cleaned = sanitize_text(item_value)
            if cleaned:
                conn = get_db_connection()
                cursor = conn.cursor()
                cursor.execute(f"INSERT OR IGNORE INTO {table_name} (name) VALUES (?)", (cleaned,))
                conn.commit()
                conn.close()
                st.toast("✅ تم الإضافة بنجاح!")
                st.cache_data.clear()
                st.rerun()

        with c1:
            st.subheader("🏢 شركة جديدة")
            new_comp = st.text_input("اسم الشركة:", key="nc")
            if st.button("حفظ الشركة"): add_item_to_db("companies", new_comp)

        with c2:
            st.subheader("♻️ نوع نفايات")
            new_waste = st.text_input("نوع النفايات:", key="nw")
            if st.button("حفظ النفايات"): add_item_to_db("wastes", new_waste)

        with c3:
            st.subheader("🚚 سائق جديد")
            new_driv = st.text_input("اسم السائق:", key="nd")
            if st.button("حفظ السائق"): add_item_to_db("drivers", DRIVER_NAME_MAP.get(sanitize_text(new_driv), sanitize_text(new_driv)))

        with c4:
            st.subheader("🚜 معدة جديدة")
            new_eq = st.text_input("اسم المعدة:", key="ne")
            if st.button("حفظ المعدة"): add_item_to_db("equipments", new_eq)

        with c5:
            st.subheader("🏞️ مردم / محطة")
            new_st = st.text_input("اسم المردم:", key="ns")
            if st.button("حفظ المردم"): add_item_to_db("stations", new_st)

        with c6:
            st.subheader("🚘 سيارة / لوحة")
            new_pl = st.text_input("رقم اللوحة:", key="np")
            if st.button("حفظ السيارة"): add_item_to_db("plates", new_pl)

    with tab_del:
        st.subheader("🗑️ إزالة عنصر من القوائم المنسدلة")
        list_type = st.selectbox(
            "اختر نوع القائمة المراد حذف عنصر منها:",
            ["الشركات", "أنواع النفايات / المخلفات", "السائقين", "المعدات", "المحطات", "السيارات"],
        )

        items_map = {
            "الشركات": (companies_list, "companies"),
            "أنواع النفايات / المخلفات": (wastes_list, "wastes"),
            "السائقين": (drivers_list, "drivers"),
            "المعدات": (equipments_list, "equipments"),
            "المحطات": (stations_list, "stations"),
            "السيارات": (plates_list, "plates"),
        }

        items, table_name = items_map[list_type]

        if items:
            item_to_delete = st.selectbox(f"اختر العنصر المراد حذفه من قائمة ({list_type}):", items)
            
            if st.button("❌ حذف العنصر المحدد"):
                st.session_state["confirm_del_item"] = item_to_delete
                st.session_state["confirm_del_table"] = table_name

            if "confirm_del_item" in st.session_state and st.session_state["confirm_del_item"] == item_to_delete:
                st.warning(f"⚠️ هل أنت برغبتك متأكد من حذف '{item_to_delete}' من قائمة {list_type}؟")
                c_del_yes, c_del_no = st.columns(2)
                if c_del_yes.button("✅ نعم، تأكيد الحذف"):
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute(f"DELETE FROM {table_name} WHERE name=?", (item_to_delete,))
                    conn.commit()
                    conn.close()

                    st.session_state["confirm_del_item"] = None
                    st.toast(f"✅ تم حذف '{item_to_delete}' بنجاح!", icon="🗑️")
                    st.cache_data.clear()
                    st.rerun()
                if c_del_no.button("❌ إلغاء والتراجع"):
                    st.session_state["confirm_del_item"] = None
                    st.rerun()
        else:
            st.info("لا توجد عناصر متاحة للحذف في هذه القائمة.")