import streamlit as st 
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from pandas.tseries.offsets import DateOffset
import time
import re


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
    # データ存在行をすべて読み込む（空白行含む）
    df_ir = pd.read_excel(
        "AIBTYO DLI.xlsx",
        sheet_name="EVENTS",
        skiprows=2,  # 3行目から読み込み（header=2と同じ効果）
        usecols="A,B,D,E,H,I,J,K,L,M,P,Q,S,T,V,W,Y"
    )

    df_ir.columns = [
        "FLT_Number", "Date", "Tail", "Branch",
        "Delay_Flag", "Delay_Time",
        "Cancel_Flag", "ShipChange_Flag", "RTO_Flag", "ATB_Flag",
        "Diversion_Flag", "EngShutDown_Flag", "Description", "Work_Performed",
        "ATA_SubChapter","Delay_Code", "Total_Maintenance_DownTime"
    ]

    # Date列を日付型に変換
    df_ir["Date"] = pd.to_datetime(df_ir["Date"], format="%d-%b-%Y", errors="coerce")

    # 空行削除（TailやDateがない行は不要）
    df_ir.dropna(subset=["Date", "Tail"], how="any", inplace=True)

    # YearMonth列作成
    df_ir["YearMonth"] = df_ir["Date"].dt.to_period("M").astype(str)

    # Aircraft_Type 判定
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
# 📊 月別推移グラフ（不具合 + DLY）
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

# DLY（月別）
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
# 棒（DLY）- 右軸
fig_total.add_trace(go.Bar(
    x=monthly_combined["YearMonth"],
    y=monthly_combined["Irreg_Total"],
    name="DLY件数",
    yaxis="y2",
    opacity=0.5
))
fig_total.update_layout(
    title="A350全体・機種別 月別FLT SQ件数 & DLY件数",
    xaxis=dict(type="category", title="年月"),
    yaxis=dict(title="FLT SQ件数", side="left"),
    yaxis2=dict(title="DLY件数", overlaying="y", side="right"),
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
# 📊 Reliability（修正版：年月を datetime に変換して昇順で表示）
# -------------------------------
import numpy as np
from pandas.tseries.offsets import DateOffset

st.subheader("Operational Reliability")

# FC データ読み込み（既存関数）
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
rel_by_type["Irreg_Count"] = rel_by_type["Irreg_Count"].fillna(0)

# Operational Reliability (%)（ゼロ除算対策）
rel_by_type["Operational_Reliability"] = np.where(
    rel_by_type["Total_FC"] > 0,
    ((rel_by_type["Total_FC"] - rel_by_type["Irreg_Count"]) / rel_by_type["Total_FC"]) * 100,
    np.nan
)

# Irregular：全機種合計（月別）
irreg_total = (
    df_irregular.groupby("YearMonth")
    .size()
    .reset_index(name="Irreg_Total")
)

# YearMonth を datetime に変換
rel_by_type["YearMonth_dt"] = pd.to_datetime(rel_by_type["YearMonth"], format="%Y-%m", errors="coerce")
irreg_total["YearMonth_dt"] = pd.to_datetime(irreg_total["YearMonth"], format="%Y-%m", errors="coerce")

# 最新日付と直近12か月の範囲
available_months = pd.concat([rel_by_type["YearMonth_dt"].dropna(), irreg_total["YearMonth_dt"].dropna()])
if available_months.empty:
    st.info("Operational Reliability 表示のための年月データが不足しています。")
else:
    max_dt = available_months.max()
    min_dt = max_dt - DateOffset(months=11)

    # データを直近12か月に絞る
    rel_by_type_12 = rel_by_type[
        (rel_by_type["YearMonth_dt"] >= min_dt) &
        (rel_by_type["YearMonth_dt"] <= max_dt)
    ].copy()
    irreg_total_12 = irreg_total[
        (irreg_total["YearMonth_dt"] >= min_dt) &
        (irreg_total["YearMonth_dt"] <= max_dt)
    ].copy()

    # 昇順ソート
    rel_by_type_12 = rel_by_type_12.sort_values("YearMonth_dt")
    irreg_total_12 = irreg_total_12.sort_values("YearMonth_dt")

    # NaN を埋める
    rel_by_type_12["Operational_Reliability"] = rel_by_type_12["Operational_Reliability"].fillna(100)

    # グラフ作成
    fig_rel_type = go.Figure()

    # 機種別折れ線
    for ac_type, color in zip(["A350-900", "A350-1000"], ["royalblue", "crimson"]):
        df_plot = rel_by_type_12[rel_by_type_12["Aircraft_Type"] == ac_type]
        if df_plot.empty:
            continue
        fig_rel_type.add_trace(go.Scatter(
            x=df_plot["YearMonth_dt"],
            y=df_plot["Operational_Reliability"],
            mode="lines+markers+text",
            text=df_plot["Operational_Reliability"].round(2).astype(str) + "%",
            textposition="top center",
            textfont=dict(size=12, color="black", family="Arial Black"),
            name=f"{ac_type} Operational Reliability (%)",
            line=dict(color=color),
            yaxis="y1"
        ))

    # DLY件数（棒グラフ）
    if not irreg_total_12.empty:
        fig_rel_type.add_trace(go.Bar(
            x=irreg_total_12["YearMonth_dt"],
            y=irreg_total_12["Irreg_Total"],
            name="DLY件数（全機種）",
            yaxis="y2",
            marker=dict(color="lightgrey"),
            opacity=0.6
        ))

    # 縦軸レンジを動的調整
    min_rel = rel_by_type_12["Operational_Reliability"].min()
    y_lower = 0 if pd.isna(min_rel) else max(0, min(95, (min_rel - 1)))

    # レイアウト
    fig_rel_type.update_layout(
        title="Operational Reliability (%)（機種別） & DLY件数（月別・直近12か月）",
        xaxis=dict(type="date", title="年月", tickformat="%Y-%m"),
        yaxis=dict(title="Operational Reliability (%)", side="left", range=[y_lower, 100]),
        yaxis2=dict(title="DLY件数", overlaying="y", side="right"),
        barmode="overlay",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0)
    )

    st.plotly_chart(fig_rel_type, use_container_width=True)


# -------------------------------
# 📊 DLY件数（ATA別・上位50位） 機種別（左右並び） + 円グラフ + フィルタ
# -------------------------------
st.subheader("DLY件数（ATA別・上位50位） 機種別 + 比率")

# 期間選択
min_date = df_irregular["Date"].min().date()
max_date = df_irregular["Date"].max().date()
start_date, end_date = st.slider(
    "期間を選択してください",
    min_value=min_date,
    max_value=max_date,
    value=(min_date, max_date),
    format="YYYY-MM-DD",
    key="slider_ata_chart"
)

# Count / OI Rate 切替
display_mode = st.radio("表示形式を選択してください", ["Count", "OI Rate (100 TO)"])

# Seat/IFE/Wi-Fi除外チェック
exclude_seat = st.checkbox(
    "Seat/IFE/Wi-Fi以外のみ表示（2500, 2520, 2528, 2521, 4400, 442X, 443X を除外）"
)

# 選択期間でフィルタ
df_period = df_irregular[
    (df_irregular["Date"].dt.date >= start_date) &
    (df_irregular["Date"].dt.date <= end_date)
].copy()

# Seat/IFE/Wi-Fi除外処理
if exclude_seat:
    exclude_patterns = [r"2500", r"2520", r"2528", r"2521", r"4400", r"442\d", r"443\d"]
    pattern = re.compile("|".join(exclude_patterns))
    df_period = df_period[~df_period["ATA_SubChapter"].astype(str).str.match(pattern)]

if df_period.empty:
    st.info("選択期間のデータがありません。期間を変更してください。")
else:
    col_900, col_1000 = st.columns(2)

    for ac_type, col in zip(["A350-900", "A350-1000"], [col_900, col_1000]):
        with col:
            df_ac = df_period[df_period["Aircraft_Type"] == ac_type]
            if df_ac.empty:
                st.info(f"{ac_type} のデータが選択期間にありません。")
                continue

            # ATA SubChapter別集計
            ata_counts = df_ac.groupby("ATA_SubChapter").size().reset_index(name="Count").sort_values("Count", ascending=False).head(50)
            ata_counts["ATA_SubChapter"] = ata_counts["ATA_SubChapter"].astype(str)
            categories = ata_counts["ATA_SubChapter"].tolist()

            # OI Rate計算
            if display_mode == "OI Rate (100 TO)":
                fc_sum = df_fc[df_fc["Aircraft_Type"] == ac_type]["FC"].sum()
                ata_counts["Count"] = np.where(fc_sum > 0, (ata_counts["Count"] / fc_sum) * 100, 0)
                yaxis_title = "OI Rate (100 TO)"
            else:
                yaxis_title = "件数"

            # 縦棒グラフ
            fig_bar = go.Figure(go.Bar(
                x=ata_counts["ATA_SubChapter"],
                y=ata_counts["Count"],
                marker_color="steelblue",
                text=ata_counts["Count"].round(2) if display_mode=="OI Rate (100 TO)" else ata_counts["Count"],
                textposition="outside"
            ))

            fig_bar.update_layout(
                title=f"DLY件数（ATA別・上位50位） - {ac_type}  {start_date} 〜 {end_date}",
                xaxis_title="ATA SubChapter",
                yaxis_title=yaxis_title,
                xaxis=dict(
                    type="category",
                    categoryorder="array",
                    categoryarray=categories,
                    tickangle=-45
                ),
                bargap=0.2,
                margin=dict(t=60, b=120, l=50, r=20),
                height=min(max(400, len(categories) * 30), 1200)
            )

            st.plotly_chart(fig_bar, use_container_width=True)

            # 円グラフ
            fig_pie = px.pie(
                ata_counts,
                names="ATA_SubChapter",
                values="Count",
                title=f"サブチャプター別 不具合比率 - {ac_type}",
                hole=0.3
            )
            fig_pie.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig_pie, use_container_width=True)

# --- Reliability グラフの下にDLY内容の表を追加 ---
st.subheader("✈Data")

# 表示列
irreg_display_cols = [
    "Date", "FLT_Number", "Tail", "Branch",
    "Delay_Code", "Delay_Time",
    "ATA_SubChapter", "Description", "Work_Performed"
]

# 表用に日付フォーマットを変更（YYYY-MM-DDのみ）
df_irregular_display = df_irregular.copy()
df_irregular_display["Date"] = df_irregular_display["Date"].dt.strftime("%Y-%m-%d")

# ATAチャプター（最初の2桁）用のフィルター選択
df_irregular_display["ATA_Chapter"] = df_irregular_display["ATA_SubChapter"].astype(str).str[:2]
ata_chapter_options = ["All"] + sorted(df_irregular_display["ATA_Chapter"].unique().tolist())
selected_ata_chapter = st.selectbox("表示する ATA チャプターを選択してください", ata_chapter_options)

# 選択に応じてフィルタリング
if selected_ata_chapter != "All":
    df_irregular_display = df_irregular_display[df_irregular_display["ATA_Chapter"] == selected_ata_chapter]

# 表示（インデックス削除）
df_irregular_sorted = df_irregular_display[irreg_display_cols] \
    .sort_values("Date", ascending=False) \
    .reset_index(drop=True)

# 表示（高さ調整のみ）
st.dataframe(df_irregular_sorted, use_container_width=True, height=500)



# ==== Data 表の下：Data表で選択した ATA_Chapter に連動した月別推移（Count / OI Rate） ====
st.subheader("📊 Selected ATA Monthly (Data表の選択に連動)")

metric_choice_data = st.radio(
    "表示指標を選択してください",
    ("Count", "OI Rate (100 TO)"),
    horizontal=True,
    key="metric_choice_data_section"
)

# ベースデータ準備（ATA_Chapter を2桁文字列で付与 / 月を日時で用意）
df_ir_base = df_irregular.copy()
df_ir_base["ATA_Chapter"] = df_ir_base["ATA_SubChapter"].astype(str).str[:2]
df_ir_base["Month"] = pd.to_datetime(df_ir_base["YearMonth"], format="%Y-%m", errors="coerce")

# 「Data」表の選択に応じてフィルタ（All の場合は全て）
if selected_ata_chapter != "All":
    df_ir_base = df_ir_base[df_ir_base["ATA_Chapter"] == selected_ata_chapter]

# 直近12か月の範囲を決定
if df_ir_base["Month"].dropna().empty:
    st.info("選択された条件に該当するデータがありません。")
else:
    max_month = df_ir_base["Month"].dropna().max()
    min_month = (max_month - DateOffset(months=11)).to_period("M").to_timestamp()
    months_range = pd.period_range(min_month, max_month, freq="M").to_timestamp()

    # 月別・機種別 Count を作成（欠月は 0 で埋める）
    def monthly_counts_for(ac_type: str) -> pd.DataFrame:
        s = (df_ir_base[df_ir_base["Aircraft_Type"] == ac_type]
             .groupby("Month").size())
        s = s.reindex(months_range, fill_value=0)
        return s.reset_index().rename(columns={"index": "Month", 0: "Count"})

    # FC（月別・機種別）を用意（OI Rate 用）
    df_fc_base = df_fc.copy()
    df_fc_base["Month"] = pd.to_datetime(df_fc_base["YearMonth"], format="%Y-%m", errors="coerce")

    def monthly_fc_for(ac_type: str) -> pd.DataFrame:
        s = (df_fc_base[df_fc_base["Aircraft_Type"] == ac_type]
             .groupby("Month")["FC"].sum())
        s = s.reindex(months_range, fill_value=0)
        return s.reset_index().rename(columns={"index": "Month", "FC": "FC"})

    col_900, col_1000 = st.columns(2)
    for ac_type, col in zip(["A350-900", "A350-1000"], [col_900, col_1000]):
        with col:
            df_cnt = monthly_counts_for(ac_type)

            if metric_choice_data == "OI Rate (100 TO)":
                df_fc_m = monthly_fc_for(ac_type)
                df_m = pd.merge(df_cnt, df_fc_m, on="Month", how="left")
                # OI Rate = Count / FC * 100（FC=0は0に）
                df_m["Value"] = np.where(df_m["FC"] > 0, (df_m["Count"] / df_m["FC"]) * 100, 0.0)
                y_vals = df_m["Value"].fillna(0)
                y_title = "OI Rate (100 TO)"
                text_vals = df_m["Value"].round(2).astype(str)
            else:
                y_vals = df_cnt["Count"].fillna(0)
                y_title = "件数"
                text_vals = df_cnt["Count"].astype(str)

            # —— ここが追加のポイント：極小値でも棒が消えないようY軸上限の最小値を固定 ——
            y_max = float(pd.Series(y_vals).max()) if len(y_vals) else 0.0
            if metric_choice_data == "OI Rate (100 TO)":
                y_upper = max(0.1, y_max * 1.2)   # 最低でも 0.1 を確保
            else:
                y_upper = max(1.0, y_max * 1.2)   # 最低でも 1 件を確保

            fig = go.Figure(go.Bar(
                x=months_range.strftime("%Y-%m"),
                y=y_vals,
                text=text_vals,
                textposition="outside",
                marker_color="steelblue",
                cliponaxis=False  # 低い棒でもラベルが軸外に出せるように
            ))

            ata_label = selected_ata_chapter if selected_ata_chapter != "All" else "All"
            fig.update_layout(
                title=f"{ac_type} - ATA {ata_label} 月別 {metric_choice_data}",
                xaxis_title="年月",
                yaxis_title=y_title,
                xaxis=dict(type="category"),
                margin=dict(t=60, b=100, l=50, r=20)
            )
            fig.update_yaxes(range=[0, y_upper])  # 縦軸を固定
            st.plotly_chart(fig, use_container_width=True)


# --- Selected ATA Monthly Count の下に月別積み上げグラフ追加 ---
st.subheader("📊 Selected ATA Monthly Count - Monthly Trend (Tail別積み上げ)")

# 横軸: 月ごと、縦軸: Count / OI Rate
metric_choice_selected_ata = st.radio(
    "表示指標を選択してください",
    ("Count", "OI Rate (100 TO)"),
    horizontal=True,
    key="metric_choice_selected_ata"
)

# 選択されたATAに紐づくサブチャプター一覧を作成（文字列化・空白除去）
df_selected_ata_filtered = df_ir_base.copy()  # df_ir_base は既に選択されたATA_Chapterでフィルタ済み
ata_subchapter_list = (
    df_selected_ata_filtered["ATA_SubChapter"]
    .dropna()
    .astype(str)
    .str.strip()
    .value_counts()
    .index.tolist()
)

# デフォルトは件数が最も多いサブチャプター
default_subchapter = ata_subchapter_list[0] if ata_subchapter_list else None

# サブチャプター選択窓
selected_ata_subchapter = st.selectbox(
    "サブチャプターを選択してください",
    options=ata_subchapter_list,
    index=0 if default_subchapter else -1,
    key="selected_ata_subchapter"
)

# 選択 ATA_SubChapter データ抽出（文字列型・空白除去で一致）
df_selected_ata = df_selected_ata_filtered[
    df_selected_ata_filtered["ATA_SubChapter"].astype(str).str.strip() == selected_ata_subchapter
].copy()

if df_selected_ata.empty:
    st.warning(f"{selected_ata_subchapter} のデータは存在しません。")
else:
    # 直近12か月の範囲を決定
    df_selected_ata["Month"] = pd.to_datetime(df_selected_ata["YearMonth"], format="%Y-%m", errors="coerce")
    max_month = df_selected_ata["Month"].max()
    min_month = (max_month - DateOffset(months=11)).to_period("M").to_timestamp()
    months_range = pd.period_range(min_month, max_month, freq="M").to_timestamp()

    # Tailごとの月別件数集計
    def monthly_tail_counts(ac_type: str) -> pd.DataFrame:
        df_ac = df_selected_ata[df_selected_ata["Aircraft_Type"] == ac_type].copy()
        # Tailごとに月別件数集計
        df_ac_grouped = (
            df_ac.groupby(["Month", "Tail"]).size().reset_index(name="Count")
        )
        # 欠月・欠Tailの補完
        df_ac_grouped = df_ac_grouped.pivot(index="Month", columns="Tail", values="Count").reindex(months_range, fill_value=0)
        df_ac_grouped = df_ac_grouped.fillna(0)
        return df_ac_grouped

    # OI Rate 用の月別FCも準備
    df_fc_base_sub = df_fc.copy()
    df_fc_base_sub["Month"] = pd.to_datetime(df_fc_base_sub["YearMonth"], format="%Y-%m", errors="coerce")

    def monthly_tail_oi_rate(ac_type: str, df_count: pd.DataFrame) -> pd.DataFrame:
        df_fc_ac = df_fc_base_sub[df_fc_base_sub["Aircraft_Type"] == ac_type].copy()
        # TailごとのFC（OI Rate計算用）
        df_fc_grouped = df_fc_ac.groupby(["Month", "Tail"])["FC"].sum().reset_index()
        df_fc_grouped = df_fc_grouped.pivot(index="Month", columns="Tail", values="FC").reindex(df_count.index, fill_value=0)
        # OI Rate = Count / FC * 100
        df_oi = df_count.divide(df_fc_grouped).multiply(100).fillna(0)
        return df_oi

    # 左右カラムにグラフ表示
    col_900, col_1000 = st.columns(2)
    for ac_type, col in zip(["A350-900", "A350-1000"], [col_900, col_1000]):
        with col:
            df_count_tail = monthly_tail_counts(ac_type)
            if metric_choice_selected_ata == "OI Rate (100 TO)":
                df_plot = monthly_tail_oi_rate(ac_type, df_count_tail)
                y_title = "OI Rate (100 TO)"
            else:
                df_plot = df_count_tail
                y_title = "件数"

            # 積み上げ棒グラフ
            fig = go.Figure()
            for tail in df_plot.columns:
                fig.add_trace(go.Bar(
                    x=df_plot.index.strftime("%Y-%m"),
                    y=df_plot[tail],
                    name=tail,
                    text=df_plot[tail].astype(int),
                    textposition="inside"
                ))

            fig.update_layout(
                barmode="stack",
                title=f"{ac_type} - {selected_ata_subchapter} 月別 {metric_choice_selected_ata} (Tail別積み上げ)",
                xaxis_title="年月",
                yaxis_title=y_title,
                xaxis=dict(type="category"),
                margin=dict(t=60, b=100, l=50, r=20)
            )

            # Y軸上限
            y_max = df_plot.sum(axis=1).max() if not df_plot.empty else 0
            fig.update_yaxes(range=[0, max(1.0, y_max * 1.2)])

            st.plotly_chart(fig, use_container_width=True)



# --- Selected ATA Monthly Count 下に Tail別積み上げグラフに続き、イレギュラー表を追加 ---
st.subheader(f"📋 {selected_ata_subchapter} のDLY詳細（Tail別）")

# 表示したい列（不要列削除・追加）
irreg_display_cols_sub = [
    "Date",
    "Tail",
    "Branch",
    "Delay_Time",
    "Description",
    "Work_Performed"
]

# 左右カラムに A350-900 / A350-1000 表を表示
col_900, col_1000 = st.columns(2)

for ac_type, col in zip(["A350-900", "A350-1000"], [col_900, col_1000]):
    with col:
        # 選択サブチャプター＆機種でフィルタ
        df_table = df_selected_ata[
            (df_selected_ata["Aircraft_Type"] == ac_type)
        ].copy()

        if df_table.empty:
            st.info(f"{ac_type} に該当するデータはありません。")
        else:
            # 存在する列だけ抽出
            existing_cols = [c for c in irreg_display_cols_sub if c in df_table.columns]
            df_table_display = df_table[existing_cols].copy()

            # 日付列を整形（存在する場合のみ）
            if "Date" in df_table_display.columns:
                df_table_display["Date"] = pd.to_datetime(
                    df_table_display["Date"], errors="coerce"
                ).dt.strftime("%Y-%m-%d")
            
            st.dataframe(df_table_display, use_container_width=True)





# ================================
# ✈ FLT SQ / Pilot Report
# ================================
st.subheader("FLT SQ / Pilot Report")

latest_month = df['YearMonth'].max()
prev_month = (pd.Period(latest_month, freq='M') - 1).strftime('%Y-%m')

ata_orders = {}  # ATA並び順を保存

col_left, col_right = st.columns(2)

for aircraft, col in zip(['A350-900', 'A350-1000'], [col_left, col_right]):
    with col:
        st.markdown(f"### ✈ {aircraft}")
        
        df_type = df[df['Aircraft_Type'] == aircraft]

        # 件数集計
        latest_counts = df_type[df_type['YearMonth'] == latest_month].groupby('ATA_Chapter').size().reset_index(name='Latest_Count')
        prev_counts = df_type[df_type['YearMonth'] == prev_month].groupby('ATA_Chapter').size().reset_index(name='Prev_Count')
        merged = pd.merge(latest_counts, prev_counts, on='ATA_Chapter', how='left').fillna(0)
        merged = merged.sort_values(by='Latest_Count', ascending=False)
        ata_orders[aircraft] = merged['ATA_Chapter'].astype(str).tolist()


# ================================
# Top Driver（月別件数推移、過去1年間総件数ベース）
# ================================
exclude_patterns = ["2520", "2521", "2528"] + \
                   [f"442{i}" for i in range(10)] + \
                   [f"443{i}" for i in range(10)]

def is_seat_related(row):
    return (row['ATA_Chapter'] == "0" and "seat" in str(row['MOD_Description']).lower())

filter_exclude_top_driver = st.checkbox("Seat/IFE/WiFi以外（Top Driverのみ適用）", value=False)

one_year_ago = (pd.Period(latest_month, freq='M') - 11).strftime('%Y-%m')
df_recent_1y_top = df[df['YearMonth'] >= one_year_ago]

if filter_exclude_top_driver:
    df_recent_1y_top = df_recent_1y_top[
        (~df_recent_1y_top['ATA_SubChapter'].isin(exclude_patterns)) &
        (~df_recent_1y_top.apply(is_seat_related, axis=1))
    ]

col_a, col_b = st.columns(2)
for col, aircraft_type in zip([col_a, col_b], ["A350-900", "A350-1000"]):
    with col:
        df_td = df_recent_1y_top[df_recent_1y_top['Aircraft_Type'] == aircraft_type]

        # 過去1年間総件数でTop10
        top_mod_list = (
            df_td.groupby('MOD_Description')
            .size()
            .sort_values(ascending=False)
            .head(10)
            .index.tolist()
        )

        df_top = df_td[df_td['MOD_Description'].isin(top_mod_list)]
        monthly_counts = (
            df_top.groupby(['YearMonth', 'MOD_Description'])
            .size()
            .reset_index(name='件数')
        )

        fig_top = px.line(
            monthly_counts,
            x='YearMonth',
            y='件数',
            color='MOD_Description',
            markers=True
        )
        fig_top.update_layout(
            title=f"{aircraft_type} Top Driver (Top10)",
            xaxis_title="月",
            yaxis_title="件数",
            legend_title="不具合内容",
            margin=dict(t=50)
        )
        st.plotly_chart(fig_top, use_container_width=True)

# ================================
# 円グラフ → 件数棒グラフ → 増加率グラフ
# ================================
col_left, col_right = st.columns(2)
for aircraft, col in zip(['A350-900', 'A350-1000'], [col_left, col_right]):
    with col:
        
        df_type_ata = df[df['Aircraft_Type'] == aircraft]

        latest_counts = df_type_ata[df_type_ata['YearMonth'] == latest_month].groupby('ATA_Chapter').size().reset_index(name='Latest_Count')
        prev_counts = df_type_ata[df_type_ata['YearMonth'] == prev_month].groupby('ATA_Chapter').size().reset_index(name='Prev_Count')
        merged = pd.merge(latest_counts, prev_counts, on='ATA_Chapter', how='left').fillna(0)
        merged = merged.sort_values(by='Latest_Count', ascending=False)
        ata_orders[aircraft] = merged['ATA_Chapter'].astype(str).tolist()

        # 円グラフ
        counts = df_type_ata[df_type_ata['YearMonth'] == latest_month].groupby('ATA_Chapter').size().reset_index(name='Count')
        fig_pie = go.Figure(go.Pie(
            labels=counts['ATA_Chapter'],
            values=counts['Count'],
            textinfo='label',
            hole=0.3
        ))
        fig_pie.update_layout(
            title=f"{aircraft} ATA別比率（{latest_month}）",
            height=400,
            margin=dict(t=40, b=0, l=0, r=0)
        )
        st.plotly_chart(fig_pie, use_container_width=True)

        # 棒グラフ（件数）
        fig_count = go.Figure(data=[
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
        fig_count.update_layout(
            barmode='group',
            title=f"ATA別不具合件数（{latest_month} と {prev_month}）",
            xaxis_title="ATA Chapter",
            yaxis_title="件数",
            xaxis=dict(type='category'),
            bargap=0.2,
            margin=dict(t=50)
        )
        st.plotly_chart(fig_count, use_container_width=True)

        # 増加率グラフ
        ata_monthly = df_type_ata.groupby(['YearMonth', 'ATA_Chapter']).size().unstack(fill_value=0).sort_index()

        if latest_month in ata_monthly.index and prev_month in ata_monthly.index:
            latest_c = ata_monthly.loc[latest_month]
            prev_c = ata_monthly.loc[prev_month]
            short_term_rate = ((latest_c - prev_c) / prev_c.replace(0, pd.NA)) * 100
            short_term_rate = pd.to_numeric(short_term_rate, errors='coerce').fillna(0)
        else:
            short_term_rate = pd.Series(0, index=ata_monthly.columns)

        ata_ma6 = ata_monthly.rolling(window=6, min_periods=2).mean()
        if latest_month in ata_ma6.index and prev_month in ata_ma6.index:
            latest_ma = ata_ma6.loc[latest_month]
            prev_ma = ata_ma6.loc[prev_month]
            long_term_rate = ((latest_ma - prev_ma) / prev_ma.replace(0, pd.NA)) * 100
            long_term_rate = pd.to_numeric(long_term_rate, errors='coerce').fillna(0)
        else:
            long_term_rate = pd.Series(0, index=ata_monthly.columns)

        rate_df = pd.DataFrame({
            'ATA_Chapter': ata_monthly.columns.astype(str),
            '短期増加率(%)': short_term_rate.round(1),
            '長期増加率(%)': long_term_rate.round(1)
        }).reset_index(drop=True)

        if aircraft in ata_orders:
            rate_df['ATA_Chapter'] = pd.Categorical(
                rate_df['ATA_Chapter'],
                categories=ata_orders[aircraft],
                ordered=True
            )
            rate_df = rate_df.sort_values('ATA_Chapter').reset_index(drop=True)

        fig_rate = go.Figure(data=[
            go.Bar(
                name='短期増加率(%)',
                x=rate_df['ATA_Chapter'],
                y=rate_df['短期増加率(%)'],
                marker_color='orange'
            ),
            go.Bar(
                name='長期増加率(%)',
                x=rate_df['ATA_Chapter'],
                y=rate_df['長期増加率(%)'],
                marker_color='green'
            )
        ])
        fig_rate.update_layout(
            barmode='group',
            title=f"増加率 (%)（{latest_month}）",
            xaxis_title="ATA Chapter",
            yaxis_title="増加率(%)",
            xaxis=dict(type='category'),
            bargap=0.2,
            margin=dict(t=30)
        )
        st.plotly_chart(fig_rate, use_container_width=True)



# -------------------------------
# ① データ要約
# -------------------------------
#st.header("① データ要約")
#latest_month = df['YearMonth'].max()
#prev_month = (pd.Period(latest_month, freq='M') - 1).strftime('%Y-%m')

#st.subheader("📋 直近1か月の不具合内容（件数上位）・機種別")
#filter_exclude = st.checkbox("📋 Seat/IFE/Wi-Fiを除く")

#if filter_exclude:
    #target_df = df[
       # (~df['ATA_SubChapter'].isin(exclude_patterns)) &
        #(~df.apply(is_seat_related, axis=1))
   # ]
#else:
    #target_df = df

#col_a, col_b = st.columns(2)
#for col, aircraft_type in zip([col_a, col_b], ["A350-900", "A350-1000"]):
   # with col:
      #  st.markdown(f"#### ✈ {aircraft_type}")
       # filtered = target_df[(target_df['YearMonth'] == latest_month) & (target_df['Aircraft_Type'] == aircraft_type)]
      #  top_mod = (
        #    filtered.groupby(['MOD_Description', 'ATA_Chapter'])
         #   .size()
         #   .reset_index(name='件数')
         #   .sort_values(by='件数', ascending=False)
       # )
       # st.dataframe(top_mod, use_container_width=True, hide_index=True, height=350)


# -------------------------------
# ATA別 月別不具合件数 + FC比 推移（左右比較）
# -------------------------------
st.header("Data by ATA chapter")

latest_date = df['Reported_Date'].max()
one_year_ago = latest_date - DateOffset(years=1)

# 不具合データ（直近1年間）
df_recent = df[df['Reported_Date'] >= one_year_ago]

# 月別・ATA別件数
ata_monthly = df_recent.groupby(['ATA_Chapter', 'YearMonth']).size().reset_index(name='Count')
ata_monthly_sum = ata_monthly.groupby('ATA_Chapter')['Count'].sum().reset_index()
ata_monthly_sorted = ata_monthly_sum.sort_values(by='Count', ascending=False)

# ATA選択
selected_ata = st.selectbox(
    "📌 ATA Chapter",
    ata_monthly_sorted['ATA_Chapter'].tolist(),
    index=0
)

# ==== 左右共通のサブチャプター順序と色を作成 ====
all_subchapters = sorted(df_recent[df_recent['ATA_Chapter'] == selected_ata]['ATA_SubChapter'].unique())
base_colors = px.colors.qualitative.Plotly
color_map = {sub: base_colors[i % len(base_colors)] for i, sub in enumerate(all_subchapters)}

col_900, col_1000 = st.columns(2)

for aircraft, col in zip(["A350-900", "A350-1000"], [col_900, col_1000]):
    with col:
        # 該当ATA & 機種データ
        ata_month = df_recent[(df_recent['ATA_Chapter'] == selected_ata) &
                              (df_recent['Aircraft_Type'] == aircraft)]

        # 月別不具合件数（1年分）
        monthly_trend = ata_month.groupby('YearMonth').size().reset_index(name='Count')

        # FCデータ（FC比は存在する月だけ計算）
        fc_monthly = df_fc[df_fc['Aircraft_Type'] == aircraft].groupby('YearMonth')['FC'].sum().reset_index()
        merged = pd.merge(monthly_trend, fc_monthly, on='YearMonth', how='left')
        merged['FC比'] = merged.apply(lambda r: r['Count'] / r['FC'] if pd.notna(r['FC']) else None, axis=1)

        # 件数＋FC比グラフ
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=merged['YearMonth'],
            y=merged['Count'],
            name='件数',
            marker_color='steelblue'
        ))
        fig.add_trace(go.Scatter(
            x=merged['YearMonth'],
            y=merged['FC比'],
            name='FC比',
            mode='lines+markers',
            yaxis='y2',
            marker_color='orange'
        ))
        fig.update_layout(
            title=f"{aircraft} ATA{selected_ata} 月別件数 & FC比",
            xaxis_title="年月",
            yaxis=dict(title="件数"),
            yaxis2=dict(title="FC比", overlaying="y", side="right"),
            hovermode="x unified",
            margin=dict(t=50)
        )
        st.plotly_chart(fig, use_container_width=True)

        # ==== サブチャプター別月別件数 ====
        sub_trend = ata_month.groupby(['YearMonth', 'ATA_SubChapter']).size().reset_index(name='Count')

        # 順序固定（左右で同じ順序）
        sub_trend['ATA_SubChapter'] = pd.Categorical(
            sub_trend['ATA_SubChapter'],
            categories=all_subchapters,
            ordered=True
        )

        fig_sub = px.line(
            sub_trend,
            x='YearMonth',
            y='Count',
            color='ATA_SubChapter',
            markers=True,
            title=f"{aircraft} ATA{selected_ata} サブチャプター別月別件数",
            color_discrete_map=color_map
        )
        fig_sub.update_layout(
            xaxis_title="年月",
            yaxis_title="件数",
            hovermode="x unified",
            margin=dict(t=50)
        )
        st.plotly_chart(fig_sub, use_container_width=True)



# --- サブチャプター選択と不具合詳細表示（機種別左右表示） ---
st.subheader("📊 Selected ATA Monthly Count")

# 選択されたATAでフィルタ
ata_filtered = df[df["ATA"] == selected_ata]

# 両機種を含めたサブチャプター件数集計
sub_counts = ata_filtered.groupby("ATA_SubChapter").size().reset_index(name="count")
sub_counts = sub_counts.sort_values("count", ascending=False)

# 両機種に共通するサブチャプターを抽出
common_subchapters = (
    set(df[df["AC_Type"] == "A350-900"]["ATA_SubChapter"]) &
    set(df[df["AC_Type"] == "A350-1000"]["ATA_SubChapter"])
)

# デフォルト値を決定（両機種に共通する中で最多件数）
if not common_subchapters:
    default_sub = sub_counts["ATA_SubChapter"].iloc[0]
else:
    sub_counts_common = sub_counts[sub_counts["ATA_SubChapter"].isin(common_subchapters)]
    if not sub_counts_common.empty:
        default_sub = sub_counts_common["ATA_SubChapter"].iloc[0]
    else:
        default_sub = sub_counts["ATA_SubChapter"].iloc[0]

# サブチャプター選択ウィジェット
selected_subchapter = st.selectbox(
    "🔍 Select Subchapter",
    options=sub_counts["ATA_SubChapter"].tolist(),
    index=sub_counts["ATA_SubChapter"].tolist().index(default_sub)
)

# 機種別の不具合詳細を左右に表示
col1, col2 = st.columns(2)

with col1:
    st.markdown("**A350-900**")
    df_900 = ata_filtered[
        (ata_filtered["AC_Type"] == "A350-900") &
        (ata_filtered["ATA_SubChapter"] == selected_subchapter)
    ]
    st.dataframe(df_900)

with col2:
    st.markdown("**A350-1000**")
    df_1000 = ata_filtered[
        (ata_filtered["AC_Type"] == "A350-1000") &
        (ata_filtered["ATA_SubChapter"] == selected_subchapter)
    ]
    st.dataframe(df_1000)


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
# サブチャプター別 機番ごとの積み上げ棒グラフ
# -------------------------------
st.markdown("#### サブチャプター別 機番ごとの積み上げ棒グラフ")

col_a, col_b = st.columns(2)

for aircraft, col in zip(["A350-900", "A350-1000"], [col_a, col_b]):
    with col:
        # 選択されたサブチャプター＆機種のデータ抽出
        df_sub_tail = df_recent[
            (df_recent['ATA_SubChapter'] == selected_sub) &
            (df_recent['Aircraft_Type'] == aircraft)
        ]

        # 月別・機番ごとの件数集計
        tail_monthly = (
            df_sub_tail.groupby(['YearMonth', 'Tail']).size().reset_index(name='Count')
        )

        # 積み上げ棒グラフ作成
        fig_tail = px.bar(
            tail_monthly,
            x='YearMonth',
            y='Count',
            color='Tail',
            title=f"{aircraft} ATA Subchapter {selected_sub} 月別件数（Tail別）",
            barmode='stack'
        )
        fig_tail.update_layout(
            xaxis_title="年月",
            yaxis_title="件数",
            hovermode="x unified",
            margin=dict(t=50)
        )
        st.plotly_chart(fig_tail, use_container_width=True)


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



































































































































