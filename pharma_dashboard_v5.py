import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import plotly.express as px
import plotly.graph_objects as go

PRIMARY_COLOR = "#0A6ED1"
UP_COLOR = "#1AA260"
DOWN_COLOR = "#D64545"

st.set_page_config(layout="wide", page_title="Con chào Mí Du ", initial_sidebar_state="expanded")

st.markdown(f"""
    <style>
    /* Tab label bigger */
    [role="tablist"] button[role="tab"] {{
        font-size:15px !important;
        padding: 8px 14px !important;
        margin-right:6px !important;
    }}
    /* Compact uploader */
    .stFileUploader > label {{ font-size:13px; }}
    .stFileUploader input[type=file] {{ height:32px; }}

    /* KPI card */
    .kpi-card {{ background-color:#F5FBFF; border-radius:8px; padding:8px; text-align:center; }}

    /* Make dataframe horizontally scrollable */
    .dataframe-container {{ overflow-x:auto; }}

    </style>
    """, unsafe_allow_html=True)

# ---------------------------
# Helpers: Excel export (tối ưu: dùng xlsxwriter ưu tiên)
# ---------------------------
def try_excel_writer(output_stream):
    try:
        import xlsxwriter  # noqa
        return pd.ExcelWriter(output_stream, engine='xlsxwriter')
    except Exception:
        return pd.ExcelWriter(output_stream, engine='openpyxl')

def export_to_excel_bytes(**dataframes):
    out = io.BytesIO()
    writer = try_excel_writer(out)
    for sheet_name, df in dataframes.items():
        name = str(sheet_name)[:31]
        try:
            if hasattr(df, "to_excel"):
                df.to_excel(writer, sheet_name=name, index=False)
            else:
                pd.DataFrame(df).to_excel(writer, sheet_name=name, index=False)
        except Exception:
            pd.DataFrame(df).to_excel(writer, sheet_name=name, index=False)
    try:
        writer.close()
    except Exception:
        try:
            writer.save()
        except Exception:
            pass
    out.seek(0)
    return out

# ---------------------------
# Small util functions (tối ưu: vectorized)
# ---------------------------
def clean_code(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    try:
        f = float(s)
        if np.isfinite(f) and float(int(f)) == f:
            return str(int(f))
    except Exception:
        pass
    if s.endswith('.0'):
        return s[:-2]
    return s

def add_total_row(df, value_cols=None):
    if df is None or df.empty:
        return df
    value_cols = value_cols or df.select_dtypes(include=[np.number]).columns.tolist()
    totals = {c: (df[c].sum() if c in value_cols else "") for c in df.columns}
    total_row = pd.DataFrame([totals])
    first_nonnum = next((c for c in df.columns if c not in value_cols), None)
    if first_nonnum:
        total_row.at[0, first_nonnum] = "TỔNG CỘNG"
    return pd.concat([df, total_row], ignore_index=True)

def style_wide_doanhso(df, col_name='Doanh số'):
    try:
        fmt_dict = {col_name: "{:,.0f}"}
        if isinstance(df, pd.DataFrame):
            sty = df.style.format(fmt_dict) if col_name in df.columns else df.style
        else:
            sty = df
        styles = [
            {'selector': 'th', 'props': [('min-width', '140px')]},
            {'selector': 'td', 'props': [('padding', '6px 8px')]},
        ]
        sty = sty.set_table_styles(styles)
        return sty
    except Exception:
        return df

def quarter_label(year, q):
    return f"Q{q} {year}"

def quarter_start_end(year, q):
    if q == 1:
        return datetime(year,1,1), datetime(year,3,31)
    if q == 2:
        return datetime(year,4,1), datetime(year,6,30)
    if q == 3:
        return datetime(year,7,1), datetime(year,9,30)
    return datetime(year,10,1), datetime(year,12,31)

# ---------------------------
# Load & Standardize data (tối ưu: cache, vectorized clean)
# ---------------------------
@st.cache_data(ttl=3600)
def load_and_standardize(uploaded, sheet_name):
    if uploaded is None:
        return pd.DataFrame()
    try:
        raw = pd.read_excel(uploaded, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"Lỗi đọc sheet {sheet_name}: {e}")
        return pd.DataFrame()

    df = raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    def find_col(possible_names, fallback_idx=None):
        cols = df.columns.tolist()
        for p in possible_names:
            for c in cols:
                if p.lower() == c.lower().strip():
                    return c
        for p in possible_names:
            for c in cols:
                if p.lower() in c.lower():
                    return c
        if fallback_idx is not None and 0 <= fallback_idx < len(cols):
            return cols[fallback_idx]
        return None

    date_col = find_col(['billing date','billing_date','date','ngày','ngay'], fallback_idx=0)
    cust_code_col = find_col(['customer','mã kh','cust code','mã khách','customer code'], fallback_idx=1)
    cust_name_col = find_col(['name','customer name','tên khách','tên'], fallback_idx=2)
    drug_code_col = find_col(['material','mã thuốc','item code','material code','mã hàng'], fallback_idx=3)
    drug_name_col = find_col(['item description','item','tên thuốc','description'], fallback_idx=4)
    qty_col = find_col(['số lượng','quantity','qty','sl'], fallback_idx=5)
    revenue_col = find_col(['ds đã trừ ck','ds','doanh thu','doanh','giá trị','revenue'], fallback_idx=7)
    emp_col = find_col(['tên tdv','tdv','nhân viên','sales','rep','employee'], fallback_idx=10)
    channel_col = find_col(['kênh','kenh','channel'], fallback_idx=12)

    df['cust_code_raw'] = df[cust_code_col].apply(clean_code).astype(str) if cust_code_col else ""
    df['cust_name_raw'] = df[cust_name_col].astype(str).fillna("").str.strip() if cust_name_col else ""
    df['drug_code_raw'] = df[drug_code_col].apply(clean_code).astype(str) if drug_code_col else ""
    df['drug_name_raw'] = df[drug_name_col].astype(str).fillna("").str.strip() if drug_name_col else ""
    df['customer_full'] = df['cust_code_raw'] + " - " + df['cust_name_raw']
    df['customer_full'] = df['customer_full'].str.strip(" -")
    df['drug_full'] = df['drug_code_raw'] + " - " + df['drug_name_raw']
    df['drug_full'] = df['drug_full'].str.strip(" -")
    df['employee'] = df[emp_col].astype(str).fillna("Không rõ") if emp_col else "Không rõ"
    df['channel'] = df[channel_col].astype(str).fillna("Không rõ") if channel_col else "Không rõ"
    df['quantity'] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0) if qty_col else 0
    df['revenue'] = pd.to_numeric(df[revenue_col], errors='coerce').fillna(0) if revenue_col else 0

    if date_col:
        df['date'] = pd.to_datetime(df[date_col], errors='coerce').dt.normalize()
    else:
        try:
            df['date'] = pd.to_datetime(df.iloc[:,0], errors='coerce').dt.normalize()
        except Exception:
            df['date'] = pd.NaT

    df = df[~df['date'].isna()].sort_values('date').reset_index(drop=True)
    if df.empty:
        return df

    df['year_month'] = df['date'].dt.to_period('M').astype(str)
    df['year'] = df['date'].dt.year
    df['month'] = df['date'].dt.month
    df['quarter'] = df['date'].dt.quarter
    return df

# ---------------------------
# Upload UI (thay đổi: chỉ 1 file, sidebar, 1-2 sheets)
# ---------------------------
st.markdown(f"<h1 style='color: {PRIMARY_COLOR}'>Phân tích kinh doanh dược phẩm</h1>", unsafe_allow_html=True)

with st.sidebar:
    st.header("Tải file")
    uploaded_file = st.file_uploader("File Excel (Sheet1: năm hiện tại, Sheet2: năm trước nếu có)", type=["xlsx","xls"], key="up_file")

if uploaded_file is None:
    st.info("Vui lòng tải file Excel để bắt đầu phân tích.")
    st.stop()

# Load sheets
excel_file = pd.ExcelFile(uploaded_file)
sheets = excel_file.sheet_names
df_now = load_and_standardize(uploaded_file, sheets[0]) if sheets else pd.DataFrame()
df_prev = load_and_standardize(uploaded_file, sheets[1]) if len(sheets) >= 2 else pd.DataFrame()

# ---------------------------
# Filters (Sidebar) - Multi-select global (thay đổi: tách date input)
# ---------------------------
st.sidebar.header("Bộ lọc (Multi-select) - Áp dụng cho tất cả Tabs")
customers_all = sorted(df_now['customer_full'].dropna().unique()) if not df_now.empty else []
sel_customers = st.sidebar.multiselect("Khách hàng", options=customers_all)

if sel_customers:
    drugs_all = sorted(df_now[df_now['customer_full'].isin(sel_customers)]['drug_full'].dropna().unique())
else:
    drugs_all = sorted(df_now['drug_full'].dropna().unique()) if not df_now.empty else []
sel_drugs = st.sidebar.multiselect("Sản phẩm", options=drugs_all)

mask_em = pd.Series(True, index=df_now.index)
if sel_customers:
    mask_em &= df_now['customer_full'].isin(sel_customers)
if sel_drugs:
    mask_em &= df_now['drug_full'].isin(sel_drugs)
emps_all = sorted(df_now[mask_em]['employee'].dropna().unique()) if not df_now.empty else []
sel_emps = st.sidebar.multiselect("Nhân viên (TDV)", options=emps_all)

channels_all = sorted(df_now['channel'].dropna().unique()) if not df_now.empty else []
sel_channels = st.sidebar.multiselect("Kênh", options=channels_all)

min_date = df_now['date'].min().date() if not df_now.empty else datetime.today().date()
max_date = df_now['date'].max().date() if not df_now.empty else datetime.today().date()
from_date = st.sidebar.date_input("Từ ngày", value=min_date, min_value=min_date, max_value=max_date)
to_date = st.sidebar.date_input("Đến ngày", value=max_date, min_value=min_date, max_value=max_date)
sel_date_range = [from_date, to_date]

# ---------------------------
# Apply filters globally (tối ưu: avoid copy where possible)
# ---------------------------
df_filtered = df_now[
    df_now['customer_full'].isin(sel_customers) if sel_customers else pd.Series(True, index=df_now.index)
]
if sel_drugs:
    df_filtered = df_filtered[df_filtered['drug_full'].isin(sel_drugs)]
if sel_emps:
    df_filtered = df_filtered[df_filtered['employee'].isin(sel_emps)]
if sel_channels:
    df_filtered = df_filtered[df_filtered['channel'].isin(sel_channels)]
if len(sel_date_range) == 2:
    start_dt = pd.to_datetime(sel_date_range[0])
    end_dt = pd.to_datetime(sel_date_range[1]) + timedelta(days=1) - timedelta(seconds=1)
    df_filtered = df_filtered[(df_filtered['date'] >= start_dt) & (df_filtered['date'] <= end_dt)]

# Filter prev similarly
def apply_filters_to_prev(df_prev, sel_customers, sel_drugs, sel_emps, sel_channels, sel_date_range):
    if df_prev.empty:
        return pd.DataFrame()
    dfp = df_prev[
        df_prev['customer_full'].isin(sel_customers) if sel_customers else pd.Series(True, index=df_prev.index)
    ]
    if sel_drugs:
        dfp = dfp[dfp['drug_full'].isin(sel_drugs)]
    if sel_emps:
        dfp = dfp[dfp['employee'].isin(sel_emps)]
    if sel_channels:
        dfp = dfp[dfp['channel'].isin(sel_channels)]
    if len(sel_date_range) == 2:
        start_dt = pd.to_datetime(sel_date_range[0])
        end_dt = pd.to_datetime(sel_date_range[1]) + timedelta(days=1) - timedelta(seconds=1)
        start_prev = start_dt - pd.DateOffset(years=1)
        end_prev = end_dt - pd.DateOffset(years=1)
        dfp = dfp[(dfp['date'] >= start_prev) & (dfp['date'] <= end_prev)]
    return dfp

df_prev_filtered = apply_filters_to_prev(df_prev, sel_customers, sel_drugs, sel_emps, sel_channels, sel_date_range)

# ---------------------------
# Compute aggregates (tối ưu: cache ttl ngắn hơn, compute only needed)
# ---------------------------
@st.cache_data(ttl=300)
def compute_aggregates(df):
    out = {}
    if df.empty:
        return out
    out['total_revenue'] = df['revenue'].sum()
    df_orders = df.assign(order_key=df['customer_full'].astype(str) + '|' + df['date'].dt.strftime('%Y-%m-%d'))
    out['total_orders'] = df_orders['order_key'].nunique()
    out['total_customers'] = df['customer_full'].nunique()
    out['total_products'] = df['drug_full'].nunique()
    out['monthly'] = df.groupby('year_month')['revenue'].sum().reset_index().sort_values('year_month')
    out['top_products'] = df.groupby('drug_full')['revenue'].sum().reset_index().sort_values('revenue', ascending=False)
    out['top_customers'] = df.groupby('customer_full')['revenue'].sum().reset_index().sort_values('revenue', ascending=False)
    out['prod_pareto'] = out['top_products'].copy()
    out['prod_pareto']['cum'] = out['prod_pareto']['revenue'].cumsum()
    total_prod_rev = out['prod_pareto']['revenue'].sum()
    out['prod_pareto']['cum_pct'] = out['prod_pareto']['cum'] / total_prod_rev if total_prod_rev != 0 else 0
    out['cust_pareto'] = out['top_customers'].copy()
    out['cust_pareto']['cum'] = out['cust_pareto']['revenue'].cumsum()
    total_cust_rev = out['cust_pareto']['revenue'].sum()
    out['cust_pareto']['cum_pct'] = out['cust_pareto']['cum'] / total_cust_rev if total_cust_rev != 0 else 0
    out['channel_summary'] = df.groupby('channel').agg({'revenue':'sum','customer_full':'nunique','drug_full':'nunique'}).reset_index().rename(columns={'revenue':'Doanh số','customer_full':'Số KH','drug_full':'Số SP'}).sort_values('Doanh số', ascending=False)
    out['cust_month_pivot'] = df.groupby(['customer_full','year_month'])['revenue'].sum().unstack(fill_value=0)
    out['prod_month_pivot'] = df.groupby(['drug_full','year_month'])['revenue'].sum().unstack(fill_value=0)
    out['last_date'] = df['date'].max()
    return out

ag = compute_aggregates(df_filtered)

# ---------------------------
# Tabs (không thay đổi cấu trúc)
# ---------------------------
tabs = st.tabs(["Tổng quan", "Pareto & TOP", "Kênh", "So sánh Quý", "So sánh năm trước", "Khách hàng & Sản phẩm", "Xuất báo cáo"])

# TAB: Tổng quan (tối ưu: avoid copy)
with tabs[0]:
    st.markdown(f"## Tổng quan (Phạm vi: {sel_date_range[0]} → {sel_date_range[1]})" if len(sel_date_range)==2 else "## Tổng quan")
    total_revenue = ag.get('total_revenue', 0)
    total_orders = ag.get('total_orders', 0)
    total_customers = ag.get('total_customers', 0)
    total_products = ag.get('total_products', 0)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Doanh số (VND)", f"{int(total_revenue):,}")
    k2.metric("Số đơn hàng", f"{int(total_orders):,}")
    k3.metric("Số khách hàng", f"{int(total_customers):,}")
    k4.metric("Số chủng loại thuốc", f"{int(total_products):,}")

    st.markdown("### Biểu đồ doanh số theo tháng")
    monthly = ag.get('monthly', pd.DataFrame())
    if monthly.empty:
        st.write("Không có dữ liệu để vẽ biểu đồ.")
    else:
        monthly['dt'] = pd.to_datetime(monthly['year_month'] + "-01", errors='coerce')
        fig = px.line(monthly, x='dt', y='revenue', markers=True, title="Doanh số theo tháng", labels={'dt':'Tháng','revenue':'Doanh số'})
        fig.update_traces(line=dict(color=PRIMARY_COLOR))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Top 10 sản phẩm theo doanh số (bảng)")
    top10_products = ag.get('top_products', pd.DataFrame()).head(10)
    if not top10_products.empty:
        top10_products[['Mã Thuốc','Tên Thuốc']] = top10_products['drug_full'].str.split(' - ', n=1, expand=True)
        df_show = top10_products[['Mã Thuốc','Tên Thuốc','revenue']].rename(columns={'revenue':'Doanh số'})
        st.dataframe(style_wide_doanhso(add_total_row(df_show)))

    st.markdown("### Top 10 khách hàng theo doanh số (bảng)")
    top10_customers = ag.get('top_customers', pd.DataFrame()).head(10)
    if not top10_customers.empty:
        top10_customers[['Mã KH','Tên KH']] = top10_customers['customer_full'].str.split(' - ', n=1, expand=True)
        df_show = top10_customers[['Mã KH','Tên KH','revenue']].rename(columns={'revenue':'Doanh số'})
        st.dataframe(style_wide_doanhso(add_total_row(df_show)))

    st.markdown("### Danh sách nguy cơ - khách hàng giảm doanh số")
    df_monthly_cust = ag.get('cust_month_pivot', pd.DataFrame())
    last_date = ag.get('last_date', None)
    if last_date is not None and not df_monthly_cust.empty:
        last_month_str = last_date.strftime('%Y-%m')
        start_10m = (pd.to_datetime(last_month_str + "-01") - pd.DateOffset(months=10)).strftime('%Y-%m')
        prev_mask_cols = [c for c in df_monthly_cust.columns if start_10m <= c < last_month_str]
        warns = []
        for cust, row in df_monthly_cust.iterrows():
            mean_prev = row[prev_mask_cols].mean() if prev_mask_cols else 0
            last_r = row.get(last_month_str, 0)
            if mean_prev > 0 and (last_r < mean_prev * 0.85):
                warns.append({'customer_full': cust, 'doanh_thu_thang': last_r, 'trung_binh_10th_truoc': mean_prev, 'pct_change': (last_r-mean_prev)/mean_prev})
        warns_df = pd.DataFrame(warns)
        if not warns_df.empty:
            warns_df = warns_df.sort_values('pct_change')
            warns_df[['Mã KH','Tên KH']] = warns_df['customer_full'].str.split(' - ', n=1, expand=True)
            warns_df = warns_df[['Mã KH','Tên KH','doanh_thu_thang','trung_binh_10th_truoc','pct_change']].rename(columns={
                'doanh_thu_thang':'Doanh thu tháng gần nhất','trung_binh_10th_truoc':'Trung bình 10 tháng trước','pct_change':'Tỷ lệ thay đổi'
            })
            st.dataframe(warns_df.style.format({"Doanh thu tháng gần nhất":"{:,.0f}","Trung bình 10 tháng trước":"{:,.0f}","Tỷ lệ thay đổi":"{:.2%}"}))
        else:
            st.write("Không tìm thấy khách hàng giảm vượt ngưỡng 15% so với trung bình 10 tháng trước.")
    else:
        st.write("Không có dữ liệu để đánh giá danh sách nguy cơ.")

    st.markdown("### Khách hàng không phát sinh doanh số (liệt kê những tháng không có phát sinh) — chỉ hiển thị THÁNG")
    cust_month_pivot = ag.get('cust_month_pivot', pd.DataFrame())
    last_6_months = sorted(df_filtered['year_month'].unique())[-6:] if not df_filtered.empty else []
    no_sale_rows = []
    for cust, row in cust_month_pivot.iterrows():
        zero_months = [m.split('-')[1] for m in last_6_months if (m in row.index and row.at[m] == 0)]
        if zero_months:
            no_sale_rows.append({'customer_full': cust, 'months_no_sale': ', '.join(zero_months)})
    no_sale_df = pd.DataFrame(no_sale_rows)
    if not no_sale_df.empty:
        no_sale_df[['Mã KH','Tên KH']] = no_sale_df['customer_full'].str.split(' - ', n=1, expand=True)
        st.dataframe(add_total_row(no_sale_df[['Mã KH','Tên KH','months_no_sale']].rename(columns={'months_no_sale':'Tháng không phát sinh'})))
    else:
        st.write("Không có khách hàng nào thiếu phát sinh trong 6 tháng gần nhất.")

    st.markdown("### Sản phẩm không phát sinh doanh số (liệt kê những tháng không có phát sinh) — chỉ hiển thị THÁNG")
    prod_month_pivot = ag.get('prod_month_pivot', pd.DataFrame())
    prod_no_sale = []
    for prod, row in prod_month_pivot.iterrows():
        zero_months = [m.split('-')[1] for m in last_6_months if (m in row.index and row.at[m] == 0)]
        if zero_months:
            prod_no_sale.append({'drug_full': prod, 'months_no_sale': ', '.join(zero_months)})
    prod_no_sale_df = pd.DataFrame(prod_no_sale)
    if not prod_no_sale_df.empty:
        prod_no_sale_df[['Mã Thuốc','Tên Thuốc']] = prod_no_sale_df['drug_full'].str.split(' - ', n=1, expand=True)
        st.dataframe(add_total_row(prod_no_sale_df[['Mã Thuốc','Tên Thuốc','months_no_sale']].rename(columns={'months_no_sale':'Tháng không phát sinh'})))
    else:
        st.write("Không có sản phẩm nào thiếu phát sinh trong 6 tháng gần nhất.")

# TAB: Pareto & TOP
with tabs[1]:
    st.markdown("## Pareto 80/20 & TOP")
    prod_pareto = ag.get('prod_pareto', pd.DataFrame())
    if not prod_pareto.empty:
        prod_pareto[['Mã Thuốc', 'Tên Thuốc']] = prod_pareto['drug_full'].str.split(' - ', n=1, expand=True)
    else:
        prod_pareto['Mã Thuốc'] = ''
        prod_pareto['Tên Thuốc'] = ''

    cust_pareto = ag.get('cust_pareto', pd.DataFrame())
    if not cust_pareto.empty:
        cust_pareto[['Mã KH','Tên KH']] = cust_pareto['customer_full'].str.split(' - ', n=1, expand=True)

    st.markdown("### Pareto - Sản phẩm (bảng + biểu đồ)")
    if not prod_pareto.empty:
        st.dataframe(style_wide_doanhso(add_total_row(prod_pareto[['Mã Thuốc','Tên Thuốc','revenue','cum_pct']].rename(columns={'revenue':'Doanh số','cum_pct':'Tỷ lệ lũy kế'}))))
        fig = go.Figure()
        fig.add_trace(go.Bar(x=prod_pareto['Tên Thuốc'].fillna(prod_pareto['Mã Thuốc']), y=prod_pareto['revenue'], name='Doanh số'))
        fig.add_trace(go.Scatter(x=prod_pareto['Tên Thuốc'].fillna(prod_pareto['Mã Thuốc']), y=prod_pareto['cum_pct'], name='Tỷ lệ lũy kế', yaxis='y2', mode='lines+markers'))
        fig.update_layout(title='Pareto Sản phẩm (Doanh số & Tỷ lệ lũy kế)', xaxis_tickangle=-45,
                          yaxis=dict(title='Doanh số'), yaxis2=dict(title='Tỷ lệ lũy kế', overlaying='y', side='right', tickformat='.0%'))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.write("Không có dữ liệu sản phẩm cho Pareto.")

    st.markdown("### Pareto - Khách hàng (bảng + biểu đồ)")
    if not cust_pareto.empty:
        st.dataframe(style_wide_doanhso(add_total_row(cust_pareto[['Mã KH','Tên KH','revenue','cum_pct']].rename(columns={'revenue':'Doanh số','cum_pct':'Tỷ lệ lũy kế'}))))
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=cust_pareto['Tên KH'].fillna(cust_pareto['Mã KH']), y=cust_pareto['revenue'], name='Doanh số'))
        fig2.add_trace(go.Scatter(x=cust_pareto['Tên KH'].fillna(cust_pareto['Mã KH']), y=cust_pareto['cum_pct'], name='Tỷ lệ lũy kế', yaxis='y2', mode='lines+markers'))
        fig2.update_layout(title='Pareto Khách hàng (Doanh số & Tỷ lệ lũy kế)', xaxis_tickangle=-45,
                          yaxis=dict(title='Doanh số'), yaxis2=dict(title='Tỷ lệ lũy kế', overlaying='y', side='right', tickformat='.0%'))
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.write("Không có dữ liệu khách hàng cho Pareto.")

    st.markdown("### TOP (bảng)")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.write("Top 10 Sản phẩm (Doanh số)")
        tprod = ag.get('top_products', pd.DataFrame())
        if not tprod.empty:
            tprod[['Mã Thuốc','Tên Thuốc']] = tprod['drug_full'].str.split(' - ', n=1, expand=True)
            st.dataframe(style_wide_doanhso(add_total_row(tprod[['Mã Thuốc','Tên Thuốc','revenue']].rename(columns={'revenue':'Doanh số'})).head(11)))
    with c2:
        st.write("Top 10 Khách hàng (Doanh số)")
        tcust = ag.get('top_customers', pd.DataFrame())
        if not tcust.empty:
            tcust[['Mã KH','Tên KH']] = tcust['customer_full'].str.split(' - ', n=1, expand=True)
            st.dataframe(style_wide_doanhso(add_total_row(tcust[['Mã KH','Tên KH','revenue']].rename(columns={'revenue':'Doanh số'})).head(11)))
    with c3:
        st.write("Top 10 Nhân viên (Doanh số)")
        temp = df_filtered.groupby('employee')['revenue'].sum().reset_index().sort_values('revenue', ascending=False).rename(columns={'revenue':'Doanh số'})
        st.dataframe(style_wide_doanhso(add_total_row(temp.head(10))))

# TAB: Kênh
with tabs[2]:
    st.markdown("## Phân tích Kênh")
    ch_sum = ag.get('channel_summary', pd.DataFrame())
    if not ch_sum.empty:
        st.dataframe(style_wide_doanhso(add_total_row(ch_sum)))
        fig = px.pie(ch_sum, values='Doanh số', names='channel', title='Cơ cấu doanh thu theo kênh')
        fig.update_traces(textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.write("Không có dữ liệu kênh.")

    st.markdown("### Doanh số theo Kênh & Nhân viên")
    ch_emp_agg = df_filtered.groupby(['channel','employee']).agg({'revenue':'sum','customer_full':'nunique','drug_full':'nunique'}).reset_index().rename(columns={'revenue':'Doanh số','customer_full':'Số KH','drug_full':'Số SP'})
    st.dataframe(style_wide_doanhso(add_total_row(ch_emp_agg.sort_values(['channel','Doanh số'], ascending=[True,False]))))

    st.markdown("### Doanh số theo Kênh & Khách hàng")
    ch_cust = df_filtered.groupby(['channel','customer_full'])['revenue'].sum().reset_index().sort_values(['channel','revenue'], ascending=[True,False])
    for ch in ch_cust['channel'].unique():
        st.markdown(f"**Kênh: {ch}**")
        sub = ch_cust[ch_cust['channel']==ch]
        sub[['Mã KH','Tên KH']] = sub['customer_full'].str.split(' - ', n=1, expand=True)
        st.dataframe(style_wide_doanhso(add_total_row(sub[['Mã KH','Tên KH','revenue']].rename(columns={'revenue':'Doanh số'}))))

    st.markdown("### Doanh số theo Kênh & Sản phẩm")
    ch_prod = df_filtered.groupby(['channel','drug_full'])['revenue'].sum().reset_index().sort_values(['channel','revenue'], ascending=[True,False])
    for ch in ch_prod['channel'].unique():
        st.markdown(f"**Kênh: {ch}**")
        subp = ch_prod[ch_prod['channel']==ch]
        subp[['Mã Thuốc','Tên Thuốc']] = subp['drug_full'].str.split(' - ', n=1, expand=True)
        st.dataframe(style_wide_doanhso(add_total_row(subp[['Mã Thuốc','Tên Thuốc','revenue']].rename(columns={'revenue':'Doanh số'}))))

# TAB: So sánh Quý
with tabs[3]:
    st.markdown("## So sánh Quý")
    q_summary = df_filtered.groupby(['year','quarter'])['revenue'].sum().reset_index().sort_values(['year','quarter'])
    if q_summary.empty:
        st.write("Không có dữ liệu quý để so sánh.")
    else:
        q_pairs = q_summary[['year','quarter']].drop_duplicates().values.tolist()
        q_options = [quarter_label(y, q) for y, q in q_pairs]
        selected_qs = st.multiselect("Chọn Quý (tối đa 4)", options=q_options, default=q_options[-2:] if len(q_options)>=2 else q_options)
        chosen = []
        for s in selected_qs:
            try:
                parts = s.split()
                q = int(parts[0].replace('Q',''))
                y = int(parts[1])
                chosen.append((y,q))
            except Exception:
                continue
        if not chosen:
            st.write("Chưa chọn Quý để hiển thị.")
        else:
            summary_list = []
            q_compare_dict = {}  # Để xuất report sau
            for y,q in chosen:
                s,e = quarter_start_end(y,q)
                mask = (df_filtered['date'] >= s) & (df_filtered['date'] <= e)
                dfq = df_filtered[mask]
                revenue_q = dfq['revenue'].sum()
                orders_q = dfq.assign(order_key=dfq['customer_full'].astype(str) + '|' + dfq['date'].dt.strftime('%Y-%m-%d'))['order_key'].nunique()
                cust_q = dfq['customer_full'].nunique()
                prod_q = dfq['drug_full'].nunique()
                emp_q = dfq['employee'].nunique()
                summary_list.append({'Quý': quarter_label(y,q), 'Doanh thu': revenue_q, 'Số đơn hàng': orders_q, 'Số KH': cust_q, 'Số SP': prod_q, 'Số NV': emp_q})
            summary_df = pd.DataFrame(summary_list)
            st.dataframe(style_wide_doanhso(add_total_row(summary_df).rename(columns={'Doanh thu':'Doanh thu'})))

            st.markdown("### Tất cả tăng/giảm sản phẩm so với Quý trước (bảng)")
            for idx, (y,q) in enumerate(chosen):
                if q == 1:
                    py, pq = y-1, 4
                else:
                    py, pq = y, q-1
                s_cur, e_cur = quarter_start_end(y,q)
                s_prev, e_prev = quarter_start_end(py,pq)
                df_cur = df_filtered[(df_filtered['date'] >= s_cur) & (df_filtered['date'] <= e_cur)]
                df_prev_q = df_filtered[(df_filtered['date'] >= s_prev) & (df_filtered['date'] <= e_prev)]
                if df_prev_q.empty or df_cur.empty:
                    st.write(f"Không đủ dữ liệu cho {quarter_label(y,q)} hoặc quý trước.")
                    continue
                cur_prod = df_cur.groupby('drug_full')['revenue'].sum().reset_index().rename(columns={'revenue':'rev_cur'})
                prev_prod = df_prev_q.groupby('drug_full')['revenue'].sum().reset_index().rename(columns={'revenue':'rev_prev'})
                compare = pd.merge(cur_prod, prev_prod, on='drug_full', how='outer').fillna(0)
                compare['delta'] = compare['rev_cur'] - compare['rev_prev']
                compare[['Mã Thuốc','Tên Thuốc']] = compare['drug_full'].str.split(' - ', n=1, expand=True)
                compare_q = compare[['Mã Thuốc','Tên Thuốc','rev_prev','rev_cur','delta']].rename(columns={'rev_prev':'Quý trước','rev_cur':'Quý hiện tại','delta':'Chênh lệch'})
                q_compare_dict[f"SanPham_Q{quarter_label(y,q)}_vs_{quarter_label(py,pq)}"] = compare_q
                st.write(f"Quý {quarter_label(y,q)} vs {quarter_label(py,pq)} - TẤT CẢ tăng/giảm (sắp xếp theo chênh lệch)")
                styled_df = add_total_row(compare_q).style.format({
                    "Quý trước": "{:,.0f}",
                    "Quý hiện tại": "{:,.0f}",
                    "Chênh lệch": "{:,.0f}"
                })
                st.dataframe(style_wide_doanhso(styled_df))

            st.markdown("### Tất cả tăng/giảm khách hàng so với Quý trước (bảng)")
            for idx, (y,q) in enumerate(chosen):
                if q == 1:
                    py, pq = y-1, 4
                else:
                    py, pq = y, q-1
                s_cur, e_cur = quarter_start_end(y,q)
                s_prev, e_prev = quarter_start_end(py,pq)
                df_cur = df_filtered[(df_filtered['date'] >= s_cur) & (df_filtered['date'] <= e_cur)]
                df_prev_q = df_filtered[(df_filtered['date'] >= s_prev) & (df_filtered['date'] <= e_prev)]
                if df_prev_q.empty or df_cur.empty:
                    continue
                cur_c = df_cur.groupby('customer_full')['revenue'].sum().reset_index().rename(columns={'revenue':'rev_cur'})
                prev_c = df_prev_q.groupby('customer_full')['revenue'].sum().reset_index().rename(columns={'revenue':'rev_prev'})
                comp_c = pd.merge(cur_c, prev_c, on='customer_full', how='outer').fillna(0)
                comp_c['delta'] = comp_c['rev_cur'] - comp_c['rev_prev']
                comp_c[['Mã KH','Tên KH']] = comp_c['customer_full'].str.split(' - ', n=1, expand=True)
                comp_c_q = comp_c[['Mã KH','Tên KH','rev_prev','rev_cur','delta']].rename(columns={'rev_prev':'Quý trước','rev_cur':'Quý hiện tại','delta':'Chênh lệch'})
                q_compare_dict[f"KhachHang_Q{quarter_label(y,q)}_vs_{quarter_label(py,pq)}"] = comp_c_q
                st.write(f"Quý {quarter_label(y,q)} vs {quarter_label(py,pq)} - TẤT CẢ tăng/giảm (sắp xếp theo chênh lệch)")
                styled_df = add_total_row(compare_q).style.format({
                    "Quý trước": "{:,.0f}",
                    "Quý hiện tại": "{:,.0f}",
                    "Chênh lệch": "{:,.0f}"
                })
                st.dataframe(style_wide_doanhso(styled_df))

# TAB: So sánh năm trước (thay đổi: bổ sung metrics KH/giao dịch/mới/mất, sửa format bảng)
with tabs[4]:
    st.markdown("## So sánh cùng kỳ (năm trước)")
    if len(sel_date_range) == 2:
        start_now = pd.to_datetime(sel_date_range[0])
        end_now = pd.to_datetime(sel_date_range[1])
    else:
        end_now = df_filtered['date'].max() if not df_filtered.empty else pd.Timestamp.today()
        start_now = end_now - pd.DateOffset(months=11)
        start_now = start_now.replace(day=1)

    start_prev = start_now - pd.DateOffset(years=1)
    end_prev = end_now - pd.DateOffset(years=1)

    st.markdown(f"**Hiện tại**: {start_now.strftime('%d/%m/%Y')} → {end_now.strftime('%d/%m/%Y')}  \n**Cùng kỳ năm trước**: {start_prev.strftime('%d/%m/%Y')} → {end_prev.strftime('%d/%m/%Y')}")

    if df_prev.empty:
        st.info("Chưa có dữ liệu năm trước → Không thể so sánh đầy đủ.")

    df_filtered['cust_key'] = df_filtered['cust_code_raw'].astype(str).str.strip()
    df_filtered['drug_key'] = df_filtered['drug_code_raw'].astype(str).str.strip()

    if not df_prev.empty:
        df_prev['cust_key'] = df_prev['cust_code_raw'].astype(str).str.strip()
        df_prev['drug_key'] = df_prev['drug_code_raw'].astype(str).str.strip()

    mask_now = (df_filtered['date'] >= start_now) & (df_filtered['date'] <= end_now)
    df_now_period = df_filtered[mask_now]

    df_prev_period = pd.DataFrame()
    if not df_prev.empty:
        mask_prev = (df_prev['date'] >= start_prev) & (df_prev['date'] <= end_prev)
        df_prev_period = df_prev[mask_prev]

    if df_now_period.empty:
        st.warning("Không có dữ liệu trong khoảng thời gian hiện tại.")
    else:
        def apply_filters(df, sel_customers, sel_drugs, sel_emps, sel_channels):
            if df.empty:
                return df
            if sel_customers:
                selected_keys = df_filtered[df_filtered['customer_full'].isin(sel_customers)]['cust_key'].unique()
                df = df[df['cust_key'].isin(selected_keys)]
            if sel_drugs:
                selected_keys = df_filtered[df_filtered['drug_full'].isin(sel_drugs)]['drug_key'].unique()
                df = df[df['drug_key'].isin(selected_keys)]
            if sel_emps:
                df = df[df['employee'].isin(sel_emps)]
            if sel_channels:
                df = df[df['channel'].isin(sel_channels)]
            return df

        df_now_period = apply_filters(df_now_period, sel_customers, sel_drugs, sel_emps, sel_channels)
        df_prev_period = apply_filters(df_prev_period, sel_customers, sel_drugs, sel_emps, sel_channels)

        if df_now_period.empty:
            st.warning("Không có dữ liệu trong khoảng thời gian hiện tại với bộ lọc đã chọn.")
        else:
            cust_now = df_now_period.groupby('cust_key')['revenue'].sum().reset_index().rename(columns={'revenue':'rev_now'})
            cust_prev = df_prev_period.groupby('cust_key')['revenue'].sum().reset_index().rename(columns={'revenue':'rev_prev'}) if not df_prev_period.empty else pd.DataFrame(columns=['cust_key','rev_prev'])
            prod_now = df_now_period.groupby('drug_key')['revenue'].sum().reset_index().rename(columns={'revenue':'rev_now'})
            prod_prev = df_prev_period.groupby('drug_key')['revenue'].sum().reset_index().rename(columns={'revenue':'rev_prev'}) if not df_prev_period.empty else pd.DataFrame(columns=['drug_key','rev_prev'])

            cust_map_df = pd.concat([
                df_filtered[['cust_code_raw','customer_full']].drop_duplicates(),
                df_prev[['cust_code_raw','customer_full']].drop_duplicates() if not df_prev.empty else pd.DataFrame()
            ]).drop_duplicates(subset='cust_code_raw', keep='first')
            cust_map = dict(zip(cust_map_df['cust_code_raw'], cust_map_df['customer_full']))

            prod_map_df = pd.concat([
                df_filtered[['drug_code_raw','drug_full']].drop_duplicates(),
                df_prev[['drug_code_raw','drug_full']].drop_duplicates() if not df_prev.empty else pd.DataFrame()
            ]).drop_duplicates(subset='drug_code_raw', keep='first')
            prod_map = dict(zip(prod_map_df['drug_code_raw'], prod_map_df['drug_full']))

            compare_cust = pd.merge(cust_prev, cust_now, on='cust_key', how='outer').fillna(0)
            compare_cust['delta'] = compare_cust['rev_now'] - compare_cust['rev_prev']
            compare_cust['pct_change'] = np.where(compare_cust['rev_prev']==0,
                                                  np.where(compare_cust['rev_now']>0, np.inf, -np.inf),
                                                  compare_cust['delta']/compare_cust['rev_prev'])
            compare_cust['customer_full'] = compare_cust['cust_key'].map(cust_map).fillna(compare_cust['cust_key'])
            compare_cust[['Mã KH','Tên KH']] = compare_cust['customer_full'].str.split(' - ', n=1, expand=True)
            compare_cust = compare_cust[['Mã KH','Tên KH','rev_prev','rev_now','delta','pct_change']]

            compare_prod = pd.merge(prod_prev, prod_now, on='drug_key', how='outer').fillna(0)
            compare_prod['delta'] = compare_prod['rev_now'] - compare_prod['rev_prev']
            compare_prod['pct_change'] = np.where(compare_prod['rev_prev']==0,
                                                  np.where(compare_prod['rev_now']>0, np.inf, -np.inf),
                                                  compare_prod['delta']/compare_prod['rev_prev'])
            compare_prod['drug_full'] = compare_prod['drug_key'].map(prod_map).fillna(compare_prod['drug_key'])
            compare_prod[['Mã Thuốc','Tên Thuốc']] = compare_prod['drug_full'].str.split(' - ', n=1, expand=True)
            compare_prod = compare_prod[['Mã Thuốc','Tên Thuốc','rev_prev','rev_now','delta','pct_change']]

            total_now = df_now_period['revenue'].sum()
            total_prev = df_prev_period['revenue'].sum() if not df_prev_period.empty else 0
            delta_total = total_now - total_prev
            pct_total = (delta_total / total_prev) if total_prev > 0 else (np.inf if total_now > 0 else 0)

            # Bổ sung metrics
            total_customers_now = df_now_period['cust_key'].nunique()
            total_customers_prev = df_prev_period['cust_key'].nunique() if not df_prev_period.empty else 0
            lost_customers_count = len(set(df_prev_period['cust_key'].unique()) - set(df_now_period['cust_key'].unique())) if not df_prev_period.empty else 0
            new_customers_count = len(set(df_now_period['cust_key'].unique()) - set(df_prev_period['cust_key'].unique())) if not df_prev_period.empty else total_customers_now
            total_orders_now = df_now_period.assign(order_key=df_now_period['customer_full'].astype(str) + '|' + df_now_period['date'].dt.strftime('%Y-%m-%d'))['order_key'].nunique()
            total_orders_prev = df_prev_period.assign(order_key=df_prev_period['customer_full'].astype(str) + '|' + df_prev_period['date'].dt.strftime('%Y-%m-%d'))['order_key'].nunique() if not df_prev_period.empty else 0

            col1, col2, col3 = st.columns(3)
            col1.metric("Doanh số hiện tại", f"{int(total_now):,}")
            col2.metric("Doanh số năm trước", f"{int(total_prev):,}")
            arrow = "▲" if delta_total > 0 else "▼" if delta_total < 0 else "→"
            color = UP_COLOR if delta_total > 0 else DOWN_COLOR if delta_total < 0 else "gray"
            pct_str = f"{pct_total:+.1%}" if np.isfinite(pct_total) else ("Mới" if total_now > 0 else "Không có")
            col3.markdown(f"**Tăng trưởng doanh số**<br><span style='color:{color};font-size:20px'>{arrow} {pct_str}</span>", unsafe_allow_html=True)

            col4, col5, col6 = st.columns(3)
            col4.metric("Số khách hàng hiện tại", f"{total_customers_now:,}")
            col5.metric("Số khách hàng năm trước", f"{total_customers_prev:,}")
            col6.metric("KH mới / KH mất", f"{new_customers_count:,} / {lost_customers_count:,}")

            col7, col8 = st.columns(2)
            col7.metric("Số giao dịch hiện tại", f"{total_orders_now:,}")
            col8.metric("Số giao dịch năm trước", f"{total_orders_prev:,}")

            def format_pct(val, now_val):
                if np.isinf(val):
                    return "Mới" if val>0 else "Mất"
                elif pd.isna(val):
                    return "—"
                else:
                    return f"{val:+.1%}"

            st.markdown("#### Bảng so sánh Khách hàng")
            cust_display = compare_cust.rename(columns={'rev_prev':'Năm trước','rev_now':'Năm hiện tại','delta':'Chênh lệch','pct_change':'% Tăng trưởng'})
            cust_display['% Tăng trưởng'] = cust_display.apply(lambda row: format_pct(row['% Tăng trưởng'], row['Năm hiện tại']), axis=1)
            # Bảng khách hàng
            styled_cust = add_total_row(cust_display).style.format({
                "Năm trước": "{:,.0f}",
                "Năm hiện tại": "{:,.0f}",
                "Chênh lệch": "{:,.0f}"
            })
            st.dataframe(style_wide_doanhso(styled_cust))

            st.markdown("#### Bảng so sánh Sản phẩm")
            prod_display = compare_prod.rename(columns={'rev_prev':'Năm trước','rev_now':'Năm hiện tại','delta':'Chênh lệch','pct_change':'% Tăng trưởng'})
            prod_display['% Tăng trưởng'] = prod_display.apply(lambda row: format_pct(row['% Tăng trưởng'], row['Năm hiện tại']), axis=1)
            # Bảng sản phẩm
            styled_prod = add_total_row(prod_display).style.format({
                "Năm trước": "{:,.0f}",
                "Năm hiện tại": "{:,.0f}",
                "Chênh lệch": "{:,.0f}"
            })
            st.dataframe(style_wide_doanhso(styled_prod))

            st.markdown("#### Top 5 Khách hàng tăng/giảm mạnh")
            top_cust_up = cust_display[cust_display['% Tăng trưởng'].astype(str).str.contains(r'\+|Mới', na=False)].nlargest(5, 'Chênh lệch')
            top_cust_down = cust_display[cust_display['% Tăng trưởng'].astype(str).str.contains(r'\-|Mất', na=False)].nsmallest(5, 'Chênh lệch')

            col_up, col_down = st.columns(2)
            with col_up:
                st.markdown("**Tăng mạnh nhất**")
                if not top_cust_up.empty:
                    # Top tăng mạnh
                    if not top_cust_up.empty:
                        styled = top_cust_up[['Mã KH', 'Tên KH', 'Chênh lệch', '% Tăng trưởng']].style.format({
                            "Chênh lệch": "{:,.0f}"
                        })
                        st.dataframe(style_wide_doanhso(styled))
                    else:
                        st.write("—")   
                else:
                    st.write("—")
            with col_down:
                st.markdown("**Giảm mạnh nhất**")
                if not top_cust_down.empty:
                    # Top giảm mạnh
                    if not top_cust_down.empty:
                        styled = top_cust_down[['Mã KH', 'Tên KH', 'Chênh lệch', '% Tăng trưởng']].style.format({
                            "Chênh lệch": "{:,.0f}"
                        })
                        st.dataframe(style_wide_doanhso(styled))
                    else:
                        st.write("—")
                else:
                    st.write("—")

            st.markdown("#### Top 5 Sản phẩm tăng/giảm mạnh")
            top_prod_up = prod_display[prod_display['% Tăng trưởng'].astype(str).str.contains(r'\+|Mới', na=False)].nlargest(5, 'Chênh lệch')
            top_prod_down = prod_display[prod_display['% Tăng trưởng'].astype(str).str.contains(r'\-|Mất', na=False)].nsmallest(5, 'Chênh lệch')

            col_up, col_down = st.columns(2)
            with col_up:
                st.markdown("**Tăng mạnh nhất**")
                if not top_prod_up.empty:
                    # Top tăng mạnh (sản phẩm)
                    if not top_prod_up.empty:
                        styled = top_prod_up[['Mã Thuốc', 'Tên Thuốc', 'Chênh lệch', '% Tăng trưởng']].style.format({
                            "Chênh lệch": "{:,.0f}"
                        })
                        st.dataframe(style_wide_doanhso(styled))
                    else:
                        st.write("—")
                else:
                    st.write("—")
            with col_down:
                st.markdown("**Giảm mạnh nhất**")
                if not top_prod_down.empty:
                    # Top giảm mạnh (sản phẩm)
                    if not top_prod_down.empty:
                        styled = top_prod_down[['Mã Thuốc', 'Tên Thuốc', 'Chênh lệch', '% Tăng trưởng']].style.format({
                            "Chênh lệch": "{:,.0f}"
                        })
                        st.dataframe(style_wide_doanhso(styled))
                    else:
                        st.write("—")
                else:
                    st.write("—")

            st.markdown("#### Biểu đồ doanh số theo tháng (cùng kỳ)")
            df_now_period['month_num'] = df_now_period['date'].dt.month
            df_prev_period['month_num'] = df_prev_period['date'].dt.month if not df_prev_period.empty else pd.Series(dtype=int)

            monthly_now = df_now_period.groupby('month_num')['revenue'].sum().reset_index().rename(columns={'revenue': 'Hiện tại'})
            monthly_prev = df_prev_period.groupby('month_num')['revenue'].sum().reset_index().rename(columns={'revenue': 'Năm trước'}) if not df_prev_period.empty else pd.DataFrame(columns=['month_num','Năm trước'])

            monthly_compare = pd.merge(monthly_prev, monthly_now, on='month_num', how='outer').fillna(0)
            monthly_compare['Tháng'] = monthly_compare['month_num'].apply(lambda x: f"Tháng {int(x)}" if pd.notna(x) else "")

            if not monthly_compare.empty:
                plot_df = pd.melt(monthly_compare, id_vars=['Tháng'], value_vars=[c for c in ['Năm trước','Hiện tại'] if c in monthly_compare.columns], var_name='Năm', value_name='Doanh số')
                fig = px.bar(plot_df, x='Tháng', y='Doanh số', color='Năm', barmode='group', text='Doanh số')
                fig.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                fig.update_layout(xaxis={'categoryorder': 'array', 'categoryarray': [f"Tháng {i}" for i in range(1,13)]})
                st.plotly_chart(fig, use_container_width=True)

                st.markdown("#### Bảng doanh số theo tháng")
                monthly_table = monthly_compare[['Tháng'] + [c for c in ['Năm trước','Hiện tại'] if c in monthly_compare.columns]]
                monthly_table['Chênh lệch'] = monthly_table.get('Hiện tại',0) - monthly_table.get('Năm trước',0)
                def pct_row(r):
                    prev = r.get('Năm trước',0)
                    if prev == 0:
                        return "Mới" if r.get('Hiện tại',0) > 0 else "—"
                    return f"{(r.get('Hiện tại',0)-prev)/prev:+.1%}"
                monthly_table['% Tăng trưởng'] = monthly_table.apply(pct_row, axis=1)
                styled_monthly = add_total_row(monthly_table).style.format({
                    "Năm trước": "{:,.0f}",
                    "Hiện tại": "{:,.0f}",
                    "Chênh lệch": "{:,.0f}"
                })
                st.dataframe(style_wide_doanhso(styled_monthly))

# TAB: Khách hàng & Sản phẩm
with tabs[5]:
    st.markdown("## Khách hàng & Sản phẩm")
    col_left, col_right = st.columns([1,1])
    with col_left:
        st.markdown("### Khách hàng - Doanh số gộp ")
        cust_all = df_filtered.groupby('customer_full')['revenue'].sum().reset_index().sort_values('revenue', ascending=False)
        if not cust_all.empty:
            cust_all[['Mã KH','Tên KH']] = cust_all['customer_full'].str.split(' - ', n=1, expand=True)
            st.dataframe(style_wide_doanhso(add_total_row(cust_all[['Mã KH','Tên KH','revenue']].rename(columns={'revenue':'Doanh số'}))))
        else:
            st.write("Không có dữ liệu khách hàng.")

        st.markdown("### Khách hàng lâu chưa mua (không phát sinh > 60 ngày)")
        last_purchase = df_filtered.groupby('customer_full')['date'].max().reset_index().rename(columns={'date':'last_date'})
        last_purchase['days_since'] = (pd.to_datetime(df_filtered['date'].max()) - last_purchase['last_date']).dt.days
        dormant = last_purchase[last_purchase['days_since'] > 60].sort_values('days_since', ascending=False)
        if not dormant.empty:
            dormant[['Mã KH','Tên KH']] = dormant['customer_full'].str.split(' - ', n=1, expand=True)
            st.dataframe(add_total_row(dormant[['Mã KH','Tên KH','last_date','days_since']].rename(columns={'last_date':'Ngày mua cuối','days_since':'Số ngày'})).style.format({"Số ngày":"{:,}"}))
        else:
            st.write("Không có khách hàng lâu chưa mua (>60 ngày).")

    with col_right:
        st.markdown("### Sản phẩm - Doanh số gộp ")
        prod_all = df_filtered.groupby('drug_full')['revenue'].sum().reset_index().sort_values('revenue', ascending=False)
        if not prod_all.empty:
            prod_all[['Mã Thuốc','Tên Thuốc']] = prod_all['drug_full'].str.split(' - ', n=1, expand=True)
            st.dataframe(style_wide_doanhso(add_total_row(prod_all[['Mã Thuốc','Tên Thuốc','revenue']].rename(columns={'revenue':'Doanh số'}))))
        else:
            st.write("Không có dữ liệu sản phẩm.")

        st.markdown("### Sản phẩm lâu không bán (>90 ngày)")
        last_prod = df_filtered.groupby('drug_full')['date'].max().reset_index().rename(columns={'date':'last_date'})
        last_prod['days_since'] = (pd.to_datetime(df_filtered['date'].max()) - last_prod['last_date']).dt.days
        prod_dormant = last_prod[last_prod['days_since'] > 90].sort_values('days_since', ascending=False)
        if not prod_dormant.empty:
            prod_dormant[['Mã Thuốc','Tên Thuốc']] = prod_dormant['drug_full'].str.split(' - ', n=1, expand=True)
            st.dataframe(add_total_row(prod_dormant[['Mã Thuốc','Tên Thuốc','last_date','days_since']].rename(columns={'last_date':'Ngày bán cuối','days_since':'Số ngày'})).style.format({"Số ngày":"{:,}"}))
        else:
            st.write("Không có sản phẩm lâu không bán (>90 ngày).")

    st.markdown("### Biểu đồ & Bảng: Doanh số theo từng tháng ")
    monthly = ag.get('monthly', pd.DataFrame())
    if monthly.empty:
        st.write("Không có dữ liệu theo tháng.")
    else:
        monthly['dt'] = pd.to_datetime(monthly['year_month'] + "-01", errors='coerce')
        figm = px.bar(monthly, x='dt', y='revenue', title='Doanh số theo tháng ', labels={'dt':'Tháng','revenue':'Doanh số'})
        st.plotly_chart(figm, use_container_width=True)
        st.markdown("Bảng doanh số theo tháng")
        monthly_table = monthly[['year_month','revenue']].rename(columns={'year_month':'Tháng','revenue':'Doanh số'})
        st.dataframe(style_wide_doanhso(add_total_row(monthly_table)))

# TAB: Xuất báo cáo 
with tabs[6]:
    st.markdown("## Xuất báo cáo - Tải Excel tổng hợp tất cả bảng dữ liệu quan trọng")
    filtered_export = df_filtered[['date','cust_code_raw','cust_name_raw','drug_code_raw','drug_name_raw','quantity','revenue','employee','channel','year_month']].rename(columns={
        'cust_code_raw':'Mã KH','cust_name_raw':'Tên KH','drug_code_raw':'Mã Thuốc','drug_name_raw':'Tên Thuốc',
        'quantity':'Số lượng','revenue':'Doanh số','employee':'Nhân viên','channel':'Kênh','year_month':'Tháng'
    })
    prod_pareto_export = ag.get('prod_pareto', pd.DataFrame())
    if not prod_pareto_export.empty:
        prod_pareto_export[['Mã Thuốc','Tên Thuốc']] = prod_pareto_export['drug_full'].str.split(' - ', n=1, expand=True)
        prod_pareto_export = prod_pareto_export[['Mã Thuốc','Tên Thuốc','revenue','cum_pct']].rename(columns={'revenue':'Doanh số','cum_pct':'Tỷ lệ lũy kế'})
    cust_pareto_export = ag.get('cust_pareto', pd.DataFrame())
    if not cust_pareto_export.empty:
        cust_pareto_export[['Mã KH','Tên KH']] = cust_pareto_export['customer_full'].str.split(' - ', n=1, expand=True)
        cust_pareto_export = cust_pareto_export[['Mã KH','Tên KH','revenue','cum_pct']].rename(columns={'revenue':'Doanh số','cum_pct':'Tỷ lệ lũy kế'})
    top_prod_export = ag.get('top_products', pd.DataFrame())
    if not top_prod_export.empty:
        top_prod_export[['Mã Thuốc','Tên Thuốc']] = top_prod_export['drug_full'].str.split(' - ', n=1, expand=True)
        top_prod_export = top_prod_export[['Mã Thuốc','Tên Thuốc','revenue']].rename(columns={'revenue':'Doanh số'})
    top_cust_export = ag.get('top_customers', pd.DataFrame())
    if not top_cust_export.empty:
        top_cust_export[['Mã KH','Tên KH']] = top_cust_export['customer_full'].str.split(' - ', n=1, expand=True)
        top_cust_export = top_cust_export[['Mã KH','Tên KH','revenue']].rename(columns={'revenue':'Doanh số'})
    monthly_export = ag.get('monthly', pd.DataFrame()).rename(columns={'year_month':'Tháng','revenue':'Doanh số'})
    channel_export = ag.get('channel_summary', pd.DataFrame())

    last_purchase = df_filtered.groupby('customer_full')['date'].max().reset_index().rename(columns={'date':'last_date'})
    last_purchase['days_since'] = (pd.to_datetime(df_filtered['date'].max()) - last_purchase['last_date']).dt.days
    dormant_export = last_purchase[last_purchase['days_since'] > 60]
    if not dormant_export.empty:
        dormant_export[['Mã KH','Tên KH']] = dormant_export['customer_full'].str.split(' - ', n=1, expand=True)
        dormant_export = dormant_export[['Mã KH','Tên KH','last_date','days_since']].rename(columns={'last_date':'Ngày mua cuối','days_since':'Số ngày'})

    last_prod = df_filtered.groupby('drug_full')['date'].max().reset_index().rename(columns={'date':'last_date'})
    last_prod['days_since'] = (pd.to_datetime(df_filtered['date'].max()) - last_prod['last_date']).dt.days
    prod_dormant_export = last_prod[last_prod['days_since'] > 90]
    if not prod_dormant_export.empty:
        prod_dormant_export[['Mã Thuốc','Tên Thuốc']] = prod_dormant_export['drug_full'].str.split(' - ', n=1, expand=True)
        prod_dormant_export = prod_dormant_export[['Mã Thuốc','Tên Thuốc','last_date','days_since']].rename(columns={'last_date':'Ngày bán cuối','days_since':'Số ngày'})

    # Bổ sung từ So sánh Quý
    q_summary_export = summary_df if 'summary_df' in locals() else pd.DataFrame()

    # Bổ sung từ So sánh năm trước
    cust_compare_export = cust_display if 'cust_display' in locals() else pd.DataFrame()
    prod_compare_export = prod_display if 'prod_display' in locals() else pd.DataFrame()
    monthly_compare_export = monthly_table if 'monthly_table' in locals() else pd.DataFrame()

    # Thêm q_compare_dict từ So sánh Quý
    sheets = {
        "Filtered_Data": filtered_export,
        "Pareto_Thuoc": prod_pareto_export,
        "Pareto_KhachHang": cust_pareto_export,
        "Top_Thuoc": top_prod_export,
        "Top_KhachHang": top_cust_export,
        "Monthly": monthly_export,
        "Channel_Summary": channel_export,
        "KhachHang_Dormant": dormant_export,
        "SanPham_Dormant": prod_dormant_export,
        "SoSanh_Quy_Summary": q_summary_export,
        "SoSanh_NamTruoc_KH": cust_compare_export,
        "SoSanh_NamTruoc_SP": prod_compare_export,
        "SoSanh_NamTruoc_Thang": monthly_compare_export,
    }
    # Thêm các so sánh quý từ dict
    if 'q_compare_dict' in locals():
        sheets.update(q_compare_dict)

    excel_bytes = export_to_excel_bytes(**sheets)
    st.download_button("⬇️ Tải Excel tổng hợp (tất cả báo cáo)", data=excel_bytes, file_name="BaoCao_PhanTich_pharma_v6_full.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

