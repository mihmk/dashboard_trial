import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from pandas.tseries.offsets import DateOffset
import time

st.set_page_config(page_title="A350 Dashboard with COA POST Count", layout="wide")

# -------------------------------
# データ読み込み関数
# -------------------------------
@st.cache_data
def load_defect_data():
    df = pd.read_excel("Defects_by_Date.xlsx")
    df = df.rename(columns={
        'Tail': 'Tail',
        'Reported Date': 'Reported_Date',
        'ATA': 'ATA',
        'MOD-Description': 'MOD_Description',
        'P/N': 'PN',
        'Corrective Action': 'Corrective_Action'
    })
    df['Reported_Date'] = pd.to_datetime(df['Reported_Date'], errors='coerce')
    df.dropna(subset=['Reported_Date'], inplace=True)
    df['Reported_Date_Str'] = df['Reported_Date'].dt.strftime('%Y-%m-%d')
    df['Reported_Date_Only'] = df['Reported_Date'].dt.date
    df['YearMonth'] = pd.to_datetime(df['Reported_Date'], errors='coerce').dt.to_period('M').astype(str)
    df['ATA_Chapter'] = df['ATA'].astype(str).str.zfill(4).str[:2]
    df['ATA_SubChapter'] = df['ATA'].astype(str).str.zfill(4).str[:4]
    df['Aircraft_Type'] = df['Tail'].apply(lambda x:
        'A350-900' if x in [f"JA{str(i).zfill(2)}XJ" for i in range(1, 17)] else (
        'A350-1000' if x in [f"JA{str(i).zfill(2)}WJ" for i in range(1, 11)] else 'その他'))
    return df

@st.cache_data
def load_irregular_data():
    # W列（Delay Code）を含めて読み込む
    df_ir = pd.read_excel(
        "AIBTYO DLI.xlsx",
        sheet_name="EVENTS",
        header=2,
        usecols="A,B,D,E,H,I,J,K,L,M,P,Q,S,T,V,Y,W"
    )

    df_ir.columns = [
        "FLT_Number", "Date", "Tail", "Branch",
        "Delay_Flag", "Delay_Time",
        "Cancel_Flag", "ShipChange_Flag", "RTO_Flag", "ATB_Flag",
        "Diversion_Flag", "EngShutDown_Flag", "Description", "Work_Performed",
        "ATA_SubChapter", "Total_Maintenance_DownTime",
        "Delay_Code"  # ← W列
    ]

    df_ir["Date"] = pd.to_datetime(df_ir["Date"], format="%d-%b-%Y", errors="coerce")
    df_ir.dropna(subset=["Date"], inplace=True)
    df_ir["YearMonth"] = df_ir["Date"].dt.to_period("M").astype(str)

    df_ir["Aircraft_Type"] = df_ir["Tail"].apply(lambda x:
        "A350-900" if x in [f"JA{str(i).zfill(2)}XJ" for i in range(1, 17)] else (
        "A350-1000" if x in [f"JA{str(i).zfill(2)}WJ" for i in range(1, 11)] else "その他")
    )
    return df_ir


df = load_defect_data()
df_irregular = load_irregular_data()

# -------------------------------
# 関数
# -------------------------------
def is_seat_related(row):
    return row['ATA_Chapter'] == '00' and 'seat' in str(row['MOD_Description']).lower()

def filter_cabin_related(df):
    exclude_patterns = ["2520", "2521", "2528"] + [f"442{i}" for i in range(10)] + [f"443{i}" for i in range(10)]
    mask1 = ~df['ATA_SubChapter'].isin(exclude_patterns)
    mask2 = ~( (df['ATA_Chapter'] == '00') & df['MOD_Description'].str.lower().str.contains('seat') )
    return df[mask1 & mask2]

# -------------------------------
# 表示
# -------------------------------
st.title("A350 Monitoring Dashboard")

latest_date = df['Reported_Date'].max()
one_year_ago = latest_date - DateOffset(years=1)
df_recent_1y = df[df['Reported_Date'] >= one_year_ago]

# 不具合件数（機種別・月別）
monthly_by_type = (
    df_recent_1y.groupby(['YearMonth', 'Aircraft_Type'])
    .size()
    .reset_index(name='Defect_Count')
    .pivot(index='YearMonth', columns='Aircraft_Type', values='Defect_Count')
    .fillna(0)
    .reset_index()
)
monthly_by_type['Defect_Total'] = monthly_by_type[['A350-900', 'A350-1000']].sum(axis=1)

# 列名に "Defect_" プレフィックスを付ける
monthly_by_type = monthly_by_type.rename(columns={
    'A350-900': 'Defect_A350-900',
    'A350-1000': 'Defect_A350-1000',
    'Defect_Total': 'Defect_Total'
})

# イレギュラー件数（機種別・月別）
monthly_irregular = (
    df_irregular.groupby(['YearMonth', 'Aircraft_Type'])
    .size()
    .reset_index(name="Irreg_Count")
    .pivot(index="YearMonth", columns="Aircraft_Type", values="Irreg_Count")
    .fillna(0)
    .reset_index()
)
monthly_irregular['Irreg_Total'] = monthly_irregular[['A350-900', 'A350-1000']].sum(axis=1)

# 列名に "Irreg_" プレフィックスを付ける
monthly_irregular = monthly_irregular.rename(columns={
    'A350-900': 'Irreg_A350-900',
    'A350-1000': 'Irreg_A350-1000',
    'Irreg_Total': 'Irreg_Total'
})

# マージ（YearMonth をキーに結合）
monthly_combined = pd.merge(monthly_by_type, monthly_irregular, on="YearMonth", how="outer").fillna(0)
monthly_combined = monthly_combined.sort_values("YearMonth")


# -------------------------------
# 📊 月別推移グラフ（不具合 + イレギュラー）
# -------------------------------
st.subheader("📊 A350 Fleet Brief")

filter_exclude_graph = st.checkbox("Seat/IFE/WiFiを除く（グラフ適用）")

def filter_cabin_related_both(df_def, df_ir):
    exclude_patterns = ["2520", "2521", "2528"] + \
                       [f"442{i}" for i in range(10)] + \
                       [f"443{i}" for i in range(10)]
    
    # 不具合データフィルタ
    mask_def = ~df_def['ATA_SubChapter'].isin(exclude_patterns) & \
               ~( (df_def['ATA_Chapter'] == '00') &
                  df_def['MOD_Description'].astype(str).str.lower().str.contains('seat', na=False) )
    
    # イレギュラーデータフィルタ
    mask_ir = ~df_ir['ATA_SubChapter'].isin(exclude_patterns) & \
              ~( (df_ir['ATA_SubChapter'].astype(str).str[:2] == '00') &
                 df_ir['Description'].astype(str).str.lower().str.contains('seat', na=False) )
    
    return df_def[mask_def], df_ir[mask_ir]


if filter_exclude_graph:
    df_recent_1y_filtered, df_irregular_filtered = filter_cabin_related_both(df_recent_1y, df_irregular)
else:
    df_recent_1y_filtered, df_irregular_filtered = df_recent_1y, df_irregular

# 不具合（月別）
monthly_by_type = (
    df_recent_1y_filtered.groupby(['YearMonth', 'Aircraft_Type'])
    .size()
    .reset_index(name='Defect_Count')
    .pivot(index='YearMonth', columns='Aircraft_Type', values='Defect_Count')
    .fillna(0)
    .reset_index()
)
monthly_by_type['Defect_Total'] = monthly_by_type[['A350-900', 'A350-1000']].sum(axis=1)
monthly_by_type = monthly_by_type.rename(columns={
    'A350-900': 'Defect_A350-900',
    'A350-1000': 'Defect_A350-1000'
})

# イレギュラー（月別）
monthly_irregular = (
    df_irregular_filtered.groupby(['YearMonth', 'Aircraft_Type'])
    .size()
    .reset_index(name="Irreg_Count")
    .pivot(index="YearMonth", columns="Aircraft_Type", values="Irreg_Count")
    .fillna(0)
    .reset_index()
)
monthly_irregular['Irreg_Total'] = monthly_irregular[['A350-900', 'A350-1000']].sum(axis=1)
monthly_irregular = monthly_irregular.rename(columns={
    'A350-900': 'Irreg_A350-900',
    'A350-1000': 'Irreg_A350-1000'
})

# マージ
monthly_combined = pd.merge(monthly_by_type, monthly_irregular, on="YearMonth", how="outer").fillna(0)
monthly_combined = monthly_combined.sort_values("YearMonth")

# グラフ作成
fig_total = go.Figure()
# 折れ線（不具合）- 左軸
for col in ["Defect_A350-900", "Defect_A350-1000", "Defect_Total"]:
    fig_total.add_trace(go.Scatter(
        x=monthly_combined["YearMonth"],
        y=monthly_combined[col],
        mode="lines+markers",
        name=f"不具合 {col.replace('Defect_', '')}",
        yaxis="y1"
    ))
# 棒（イレギュラー）- 右軸
fig_total.add_trace(go.Bar(
    x=monthly_combined["YearMonth"],
    y=monthly_combined["Irreg_Total"],
    name="イレギュラー件数",
    yaxis="y2",
    opacity=0.5
))
fig_total.update_layout(
    title="A350全体・機種別 月別不具合件数 & イレギュラー件数",
    xaxis=dict(type="category", title="年月"),
    yaxis=dict(title="不具合件数", side="left"),
    yaxis2=dict(title="イレギュラー件数", overlaying="y", side="right"),
    barmode="overlay"
)
st.plotly_chart(fig_total, use_container_width=True)

# --- FCデータ読み込み関数 ---
@st.cache_data
def load_fc_data():
    import re

    file_path = "FHFC(Airbus).xlsx"
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    all_data = []

    for sheet in sheet_names:
        try:
            # 年月正規化
            match = re.match(r"(\d{4})([A-Z]{3})", sheet)
            if not match:
                continue
            year, mon_str = match.groups()
            month_map = {
                "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
                "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
                "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"
            }
            if mon_str not in month_map:
                continue
            yearmonth = f"{year}-{month_map[mon_str]}"

            df_sheet = pd.read_excel(file_path, sheet_name=sheet, header=None)

            # F列が "FCY" の行だけ抽出
            mask_fcy = df_sheet.iloc[:, 5].astype(str).str.strip().str.upper() == "FCY"
            df_fcy = df_sheet.loc[mask_fcy, [1, 3]].copy()
            df_fcy.columns = ["Tail", "FC"]

            # A350-900 or A350-1000 判定
            # → Tail 番号の位置から判断（JA01XJ〜JA39XJくらいがA350-900、それ以降はA350-1000）
            def get_type(tail):
                try:
                    num = int(str(tail)[2:4])
                    return "A350-900" if num <= 39 else "A350-1000"
                except:
                    return "Unknown"

            df_fcy["Aircraft_Type"] = df_fcy["Tail"].apply(get_type)
            df_fcy["YearMonth"] = yearmonth

            # 数値化
            df_fcy["FC"] = pd.to_numeric(df_fcy["FC"], errors="coerce")
            df_fcy = df_fcy.dropna(subset=["FC"])

            all_data.append(df_fcy)

        except Exception as e:
            st.warning(f"{sheet} 読み込み失敗: {e}")

    if all_data:
        return pd.concat(all_data, ignore_index=True)
    else:
        return pd.DataFrame(columns=["Tail", "FC", "Aircraft_Type", "YearMonth"])



# -------------------------------
# 📊 Reliability
# -------------------------------

st.subheader("📈 Reliability (機種別)")

# FC データ読み込み
df_fc = load_fc_data()

# Irregular データ（月別・機種別）
irreg_by_type = (
    df_irregular.groupby(["YearMonth", "Aircraft_Type"])
    .size()
    .reset_index(name="Irreg_Count")
)

# FC データ（月別・機種別）
fc_by_type = (
    df_fc.groupby(["YearMonth", "Aircraft_Type"], as_index=False)["FC"].sum()
    .rename(columns={"FC": "Total_FC"})
)

# マージ
rel_by_type = pd.merge(fc_by_type, irreg_by_type, on=["YearMonth", "Aircraft_Type"], how="left")

# Operational Reliability (%) 計算
rel_by_type["Operational_Reliability"] = (
    (rel_by_type["Total_FC"] - rel_by_type["Irreg_Count"]) / rel_by_type["Total_FC"]
) * 100

# 欠損補完
rel_by_type = rel_by_type.fillna({"Irreg_Count": 0, "Operational_Reliability": 100})

# グラフ
fig_rel_type = go.Figure()

for ac_type in ["A350-900", "A350-1000"]:
    df_plot = rel_by_type[rel_by_type["Aircraft_Type"] == ac_type]
    fig_rel_type.add_trace(go.Scatter(
        x=df_plot["YearMonth"],
        y=df_plot["Operational_Reliability"],
        mode="lines+markers+text",
        text=df_plot["Operational_Reliability"].round(2).astype(str) + "%",
        textposition="top center",
        textfont=dict(size=14, color="black", family="Arial Black"),  # ← サイズ・色・太さ
        name=f"{ac_type} Operational Reliability (%)",
        yaxis="y1"
    ))

# イレギュラー件数は棒グラフ（全機種合計）
irreg_total = (
    df_irregular.groupby("YearMonth")
    .size()
    .reset_index(name="Irreg_Total")
)
fig_rel_type.add_trace(go.Bar(
    x=irreg_total["YearMonth"],
    y=irreg_total["Irreg_Total"],
    name="イレギュラー件数（全機種）",
    yaxis="y2",
    opacity=0.5
))

fig_rel_type.update_layout(
    title="Operational Reliability (%)（機種別） & イレギュラー件数（月別）",
    xaxis=dict(type="category", title="年月"),
    yaxis=dict(title="Operational Reliability (%)", side="left", range=[95, 100]),
    yaxis2=dict(title="イレギュラー件数", overlaying="y", side="right"),
    barmode="overlay",
    hovermode="x unified"
)

st.plotly_chart(fig_rel_type, use_container_width=True)


# --- Reliability グラフの下にイレギュラー内容の表を追加 ---
st.subheader("✈ イレギュラー事象一覧")

# 表示列を変更（Delay_Flag → Delay_Code）
irreg_display_cols = [
    "Date", "FLT_Number", "Tail", "Branch",
    "Delay_Code", "Delay_Time",
    "ATA_SubChapter", "Description", "Work_Performed"
]

df_irregular_sorted = df_irregular[irreg_display_cols].sort_values("Date", ascending=False)

st.dataframe(df_irregular_sorted, use_container_width=True, height=500)



# -------------------------------
# 📊 不具合件数上位10のMOD_Description月次推移（機種別）
# -------------------------------
st.subheader("📊 上位10件の不具合内容（MOD_Description）の月次推移")
col5, col6 = st.columns(2)

for aircraft, col in zip(['A350-900', 'A350-1000'], [col5, col6]):
    with col:
        st.markdown(f"### ✈ {aircraft}")
        df_aircraft = df_filtered[df_filtered['Aircraft_Type'] == aircraft]
        top10_mods = (
            df_aircraft['MOD_Description']
            .value_counts()
            .nlargest(10)
            .index
        )
        trend_data = (
            df_aircraft[df_aircraft['MOD_Description'].isin(top10_mods)]
            .groupby(['YearMonth', 'MOD_Description'])
            .size()
            .reset_index(name='Count')
            .sort_values(by='YearMonth')
        )
        fig_top10 = px.line(
            trend_data,
            x='YearMonth',
            y='Count',
            color='MOD_Description',
            markers=True,
            title=f"{aircraft}：上位10不具合の月次件数推移（直近1年）",
            labels={'Count': '件数', 'MOD_Description': '不具合内容'}
        )
        fig_top10.update_layout(
            xaxis_title="年月",
            yaxis_title="件数",
            xaxis=dict(type='category'),
            hovermode='x unified'
        )
        st.plotly_chart(fig_top10, use_container_width=True)

# -------------------------------
# ① データ要約
# -------------------------------
st.header("① データ要約")
latest_month = df['YearMonth'].max()
prev_month = (pd.Period(latest_month, freq='M') - 1).strftime('%Y-%m')

st.subheader("📋 直近1か月の不具合内容（件数上位）・機種別")
filter_exclude = st.checkbox("📋 Seat/IFE/Wi-Fiを除く")

if filter_exclude:
    target_df = df[
        (~df['ATA_SubChapter'].isin(exclude_patterns)) &
        (~df.apply(is_seat_related, axis=1))
    ]
else:
    target_df = df

col_a, col_b = st.columns(2)
for col, aircraft_type in zip([col_a, col_b], ["A350-900", "A350-1000"]):
    with col:
        st.markdown(f"#### ✈ {aircraft_type}")
        filtered = target_df[(target_df['YearMonth'] == latest_month) & (target_df['Aircraft_Type'] == aircraft_type)]
        top_mod = (
            filtered.groupby(['MOD_Description', 'ATA_Chapter'])
            .size()
            .reset_index(name='件数')
            .sort_values(by='件数', ascending=False)
        )
        st.dataframe(top_mod, use_container_width=True, hide_index=True, height=350)

st.subheader("📈 ATAサブチャプターごとの不具合件数増加率・機種別")
st.markdown("#### 📉 長期トレンド（6か月移動平均）")
col1, col2 = st.columns(2)
for aircraft, col in zip(['A350-900', 'A350-1000'], [col1, col2]):
    with col:
        st.markdown(f"### ✈ {aircraft}")
        df_type = df[df['Aircraft_Type'] == aircraft]
        if filter_exclude:
            df_type = filter_cabin_related(df_type)
        ata_monthly = df_type.groupby(['YearMonth', 'ATA_SubChapter']).size().unstack(fill_value=0).sort_index()
        ata_ma12 = ata_monthly.rolling(window=6, min_periods=2).mean()
        if latest_month in ata_ma12.index and prev_month in ata_ma12.index:
            latest_ma = ata_ma12.loc[latest_month]
            prev_ma = ata_ma12.loc[prev_month]
            increase_rate = ((latest_ma - prev_ma) / prev_ma.replace(0, pd.NA)) * 100
            increase_rate = pd.to_numeric(increase_rate, errors='coerce').dropna()
            alert_df = pd.DataFrame({
                'ATA_SubChapter': increase_rate.index,
                '増加率(%)': increase_rate.round(1).values,
                '今月件数': [ata_monthly.loc[latest_month, ata] for ata in increase_rate.index]
            })
            mod_map = df_type[df_type['YearMonth'] == latest_month].groupby('ATA_SubChapter')['MOD_Description'].agg(lambda x: x.value_counts().idxmax()).to_dict()
            alert_df['代表的な不具合内容'] = alert_df['ATA_SubChapter'].map(mod_map).fillna("")
            alert_df = alert_df.sort_values(by='増加率(%)', ascending=False)
            st.dataframe(alert_df, use_container_width=True, hide_index=True, height=350)
        else:
            st.info(f"{aircraft} の移動平均を算出するのに十分な月次データがありません。")


st.markdown("#### 📈 短期トレンド（当月 vs 前月）")

col3, col4 = st.columns(2)
for aircraft, col in zip(['A350-900', 'A350-1000'], [col3, col4]):
    with col:
        st.markdown(f"### ✈ {aircraft}")

        df_type = df_recent_1y[df_recent_1y['Aircraft_Type'] == aircraft]

        if filter_exclude:
            df_type = filter_cabin_related(df_type)

        ata_monthly = df_type.groupby(['YearMonth', 'ATA_SubChapter']).size().unstack(fill_value=0).sort_index()

        if latest_month in ata_monthly.index and prev_month in ata_monthly.index:
            latest_counts = ata_monthly.loc[latest_month]
            prev_counts = ata_monthly.loc[prev_month]

            short_term_rate = ((latest_counts - prev_counts) / prev_counts.replace(0, pd.NA)) * 100
            short_term_rate = pd.to_numeric(short_term_rate, errors='coerce').dropna()
            
            short_df = pd.DataFrame({
                
                'ATA_SubChapter': short_term_rate.index,
                '増加率(%)': short_term_rate.round(1).values,
                '今月件数': latest_counts[short_term_rate.index].values
            })

            mod_map = df_type[df_type['YearMonth'] == latest_month] \
                .groupby('ATA_SubChapter')['MOD_Description'] \
                .agg(lambda x: x.value_counts().idxmax()).to_dict()

            short_df['代表的な不具合内容'] = short_df['ATA_SubChapter'].map(mod_map).fillna("")
            short_df = short_df.sort_values(by='増加率(%)', ascending=False)

            st.dataframe(short_df, use_container_width=True, hide_index=True, height=350)
        else:
            st.info(f"{aircraft} の短期比較を算出するのに十分な月次データがありません。")




# -------------------------------
# ② ATA別不具合件数（直近月）
# -------------------------------
st.header("② ATA別不具合件数（直近月）")

df_latest_month = df[df['YearMonth'] == latest_month]
df_prev_month = df[df['YearMonth'] == prev_month]

latest_counts = df_latest_month.groupby('ATA_Chapter').size().reset_index(name='Latest_Count')
prev_counts = df_prev_month.groupby('ATA_Chapter').size().reset_index(name='Prev_Count')

merged = pd.merge(latest_counts, prev_counts, on='ATA_Chapter', how='left').fillna(0)
merged = merged.sort_values(by='Latest_Count', ascending=False)

fig_ata = go.Figure(data=[
    go.Bar(
        name=f"{latest_month}",
        x=merged['ATA_Chapter'],
        y=merged['Latest_Count'],
        marker_color='steelblue',
        text=merged['Latest_Count'],
        textposition='outside'
    ),
    go.Bar(
        name=f"{prev_month}",
        x=merged['ATA_Chapter'],
        y=merged['Prev_Count'],
        marker_color='lightcoral',
        text=merged['Prev_Count'],
        textposition='outside'
    )
])

fig_ata.update_layout(
    barmode='group',
    title=f"ATA別不具合件数（{latest_month} と {prev_month}）",
    xaxis_title="ATA Chapter",
    yaxis_title="件数",
    xaxis=dict(type='category'),
    bargap=0.2
)

st.plotly_chart(fig_ata, use_container_width=True)

# -------------------------------
# ③ ATA別 月別不具合件数推移（直近1年）
# -------------------------------
st.header("③ ATA別 月別不具合件数推移（直近1年）")

latest_date = df['Reported_Date'].max()
one_year_ago = latest_date - DateOffset(years=1)
df_recent = df[df['Reported_Date'] >= one_year_ago]

ata_monthly = df_recent.groupby(['ATA_Chapter', 'YearMonth']).size().reset_index(name='Count')
ata_monthly_sum = ata_monthly.groupby('ATA_Chapter')['Count'].sum().reset_index()
ata_monthly_sorted = ata_monthly_sum.sort_values(by='Count', ascending=False)

selected_ata = st.selectbox(
    "📌 ATAチャプターを選択",
    ata_monthly_sorted['ATA_Chapter'].tolist(),
    index=0
)

ata_month = df_recent[df_recent['ATA_Chapter'] == selected_ata]

monthly_trend = ata_month.groupby('YearMonth').size().reset_index(name='Count')
fig_bar = px.bar(monthly_trend, x='YearMonth', y='Count', title=f"ATA{selected_ata} の月別不具合件数推移")
st.plotly_chart(fig_bar, use_container_width=True)

sub_trend = (
    ata_month.groupby(['YearMonth', 'ATA_SubChapter'])
    .size()
    .reset_index(name='Count')
)
fig_line = px.line(
    sub_trend,
    x='YearMonth',
    y='Count',
    color='ATA_SubChapter',
    markers=True,
    title=f"ATA{selected_ata} のサブチャプター別不具合件数推移（直近1年）"
)
fig_line.update_layout(
    xaxis_title="年月",
    yaxis_title="件数",
    hovermode="x unified"
)
st.plotly_chart(fig_line, use_container_width=True)

# --- サブチャプター選択と不具合詳細表示 ---
st.subheader("🔍 サブチャプターごとの不具合詳細")

subchapter_counts = ata_month['ATA_SubChapter'].value_counts().reset_index()
subchapter_counts.columns = ['ATA_SubChapter', 'Count']

selected_sub = st.selectbox("サブチャプターを選択（件数順）", subchapter_counts['ATA_SubChapter'].tolist())

sub_df = ata_month[ata_month['ATA_SubChapter'] == selected_sub].copy()

# Tailでフィルター可能なインターフェースを追加
unique_tails = sorted(sub_df['Tail'].dropna().unique())
tail_filter = st.selectbox("✈️ 表示する機体（Tail）を選択", options=["すべて"] + unique_tails)

if tail_filter != "すべて":
    sub_df = sub_df[sub_df['Tail'] == tail_filter]

sub_df_display = sub_df[['ATA_SubChapter', 'Reported_Date_Only', 'Tail', 'MOD_Description', 'Corrective_Action']]
sub_df_display = sub_df_display.sort_values(by='Reported_Date_Only', ascending=False)

st.dataframe(sub_df_display, use_container_width=True, hide_index=True)

# -------------------------------
# 🔢 サブチャプター内 不具合内容別件数推移（折れ線グラフ）
# -------------------------------
if not sub_df.empty:
    # 月単位へ変換
    sub_df['YearMonth'] = sub_df['Reported_Date'].dt.to_period('M').astype(str)

    # 件数上位5種類の不具合だけを表示（多すぎると見づらいため）
    top_faults = (
        sub_df['MOD_Description']
        .value_counts()
        .head(5)                       # 上位5件
        .index
    )

    trend_data = (
        sub_df[sub_df['MOD_Description'].isin(top_faults)]
        .groupby(['YearMonth', 'MOD_Description'])
        .size()
        .reset_index(name='Count')
        .sort_values(by='YearMonth')
    )

    if not trend_data.empty:
        fig_fault_trend = px.line(
            trend_data,
            x='YearMonth',
            y='Count',
            color='MOD_Description',
            markers=True,
            title=f"📈 サブチャプター {selected_sub} 内 不具合内容別 月次件数推移（上位5種類）",
            labels={'Count': '件数', 'MOD_Description': '不具合内容'}
        )
        fig_fault_trend.update_layout(
            xaxis_title="年月",
            yaxis_title="件数",
            hovermode="x unified"
        )
        st.plotly_chart(fig_fault_trend, use_container_width=True)
    else:
        st.info("このサブチャプターには表示できる不具合データがありません。")
else:
    st.info("選択された条件に合致するデータがありません。")

# -------------------------------
# ⑤ 部品（P/N）検索と履歴（履歴一覧表示 + 件数 + 日付絞り込み）
# -------------------------------
st.header("⑤ 部品（P/N）検索と履歴")

col1, col2 = st.columns(2)
with col1:
    pn_search = st.text_input("🔍 P/Nで検索（部分一致）")
with col2:
    ata_search = st.text_input("🔍 ATAチャプターで検索（2桁）")

# データ準備（PN・ATAが欠損していないもの）
pn_data = df[df['PN'].notna()].copy()
pn_data = pn_data[pn_data['ATA_Chapter'].notna()]

# 検索条件でフィルタリング
if pn_search:
    pn_data = pn_data[pn_data['PN'].astype(str).str.contains(pn_search, case=False, na=False)]
if ata_search:
    pn_data = pn_data[pn_data['ATA_Chapter'].astype(str).str.zfill(2).str.contains(ata_search.zfill(2))]

# 日付範囲指定（Reported_Date_Only）
if not pn_data.empty:
    min_date = pn_data['Reported_Date_Only'].min()
    max_date = pn_data['Reported_Date_Only'].max()
    start_date, end_date = st.slider(
        "📅 表示する日付範囲を選択",
        min_value=min_date,
        max_value=max_date,
        value=(min_date, max_date),
        format="YYYY-MM-DD"
    )
    pn_data = pn_data[
        (pn_data['Reported_Date_Only'] >= start_date) & (pn_data['Reported_Date_Only'] <= end_date)
    ]

# 表示用データ
history_table = pn_data[['PN', 'Reported_Date_Only', 'Tail', 'MOD_Description']]
history_table = history_table.sort_values(by='Reported_Date_Only', ascending=False)

# 件数表示
record_count = len(history_table)
st.markdown(f"🔢 **検索結果：{record_count} 件**")

# 表表示
st.markdown("📋 **交換履歴一覧**")
st.dataframe(history_table, use_container_width=True, hide_index=True)

# -------------------------------
# 📊 PN検索時の積み上げ棒グラフ
# -------------------------------
if pn_search and not pn_data.empty:
    # 月単位でグループ化（PN + Tail）
    pn_data['YearMonth'] = pd.to_datetime(pn_data['Reported_Date']).dt.to_period('M').astype(str)
    
    bar_data = (
        pn_data.groupby(['YearMonth', 'Tail'])
        .size()
        .reset_index(name='Count')
    )

    fig_pn_bar = px.bar(
        bar_data,
        x='YearMonth',
        y='Count',
        color='Tail',
        title=f"📊 P/N: {pn_search} の交換履歴（Tail別・月別 件数）",
        labels={'Count': '交換件数', 'Tail': '機番'},
    )

    fig_pn_bar.update_layout(
        barmode='stack',
        xaxis_title="年月",
        yaxis_title="件数",
        xaxis=dict(type='category'),
        hovermode='x unified',
        height=400
    )

    st.plotly_chart(fig_pn_bar, use_container_width=True)



# -------------------------------
# ① 入力フォーム
# -------------------------------
st.markdown("#### COA番号を入力してください（例：COA12-34567ER01）")

col1, col2, col3 = st.columns(3)
with col1:
    coa_xx = st.text_input("XX (2桁)", max_chars=2)
with col2:
    coa_yyyyy = st.text_input("YYYYY (5桁)", max_chars=5)
with col3:
    coa_z = st.text_input("Z (1桁)", max_chars=1)

full_coa_code = f"COA{coa_xx}{coa_yyyyy}ER0{coa_z}"

# -------------------------------
# ② 検索ボタン
# -------------------------------
if st.button("検索"):
    if len(coa_xx) == 2 and len(coa_yyyyy) == 5 and len(coa_z) == 1:
        if platform.system() == "Windows":
            try:
                # SAP接続処理（Windows環境限定）
                SapGuiAuto = win32com.client.GetObject("SAPGUI")
                application = SapGuiAuto.GetScriptingEngine
                connection = application.Children(0)
                session = connection.Children(0)

                session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZDMPM_VAR_TAB_DISP"
                session.findById("wnd[0]/tbar[0]/btn[0]").press()

                session.findById("wnd[0]/usr/radP_RBVT").Select()
                session.findById("wnd[0]/usr/ctxtP_VTAB").Text = "D_AC_350"
                session.findById("wnd[0]/usr/radP_RBCVD").Select()
                session.findById("wnd[0]/tbar[1]/btn[8]").press()

                alv = session.findById("wnd[0]/usr/cntlCONTAINER_ALV/shellcont/shell")
                row_count = alv.RowCount

                result = []
                for i in range(row_count):
                    chara = alv.GetCellValue(i, "CHARS")
                    if full_coa_code in chara:
                        for ship in [
                            "JA01XJ", "JA02XJ", "JA03XJ", "JA04XJ", "JA05XJ", "JA06XJ", "JA07XJ",
                            "JA08XJ", "JA09XJ", "JA10XJ", "JA11XJ", "JA12XJ", "JA14XJ", "JA15XJ", "JA16XJ",
                            "JA17XJ", "JA18XJ", "JA19XJ", "JA01WJ", "JA02WJ", "JA03WJ", "JA04WJ", "JA05WJ",
                            "JA06WJ", "JA07WJ", "JA08WJ", "JA09WJ", "JA10WJ", "JA11WJ", "JA12WJ", "JA13WJ"
                        ]:
                            try:
                                status = alv.GetCellValue(i, ship)
                                result.append({'Ship': ship, 'Status': status})
                            except:
                                continue

                df_result = pd.DataFrame(result)
                df_post = df_result[df_result['Status'] == 'C']
                post_count = df_post.shape[0]

                st.success(f"{full_coa_code} のPOST状態（C）の機番数： {post_count} 機")
                st.dataframe(df_post)

            except Exception as e:
                st.error(f"SAPアクセスエラー: {e}")
        else:
            st.warning("この機能はWindows環境（SAP GUIがインストールされている環境）でのみ利用できます。")
    else:
        st.warning("すべての入力欄（XX・YYYYY・Z）を正しく入力してください。")

















