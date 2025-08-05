import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from pandas.tseries.offsets import DateOffset
import time
import sys

# win32com.client はローカルのみ
try:
    import win32com.client
    SAP_AVAILABLE = True
except ImportError:
    SAP_AVAILABLE = False

st.set_page_config(page_title="A350 Dashboard with COA POST Count", layout="wide")

@st.cache_data
def load_data():
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
    df['YearMonth'] = pd.to_datetime(df['Reported_Date'], errors='coerce')
    df['YearMonth'] = df['YearMonth'].dt.to_period('M').astype(str)
    df['ATA_Chapter'] = df['ATA'].astype(str).str.zfill(4).str[:2]
    df['ATA_SubChapter'] = df['ATA'].astype(str).str.zfill(4).str[:4]
    df['Aircraft_Type'] = df['Tail'].apply(lambda x:
        'A350-900' if x in [f"JA{str(i).zfill(2)}XJ" for i in range(1, 17)] else (
        'A350-1000' if x in [f"JA{str(i).zfill(2)}WJ" for i in range(1, 11)] else 'その他'))
    return df

def is_seat_related(row):
    return row['ATA_Chapter'] == '00' and 'seat' in str(row['MOD_Description']).lower()

def filter_cabin_related(df):
    exclude_patterns = ["2520", "2521", "2528"] + [f"442{i}" for i in range(10)] + [f"443{i}" for i in range(10)]
    mask1 = ~df['ATA_SubChapter'].isin(exclude_patterns)
    mask2 = ~( (df['ATA_Chapter'] == '00') & df['MOD_Description'].str.lower().str.contains('seat') )
    return df[mask1 & mask2]

df = load_data()
st.title("🛫 A350 不具合モニタリングダッシュボード")

latest_date = df['Reported_Date'].max()
one_year_ago = latest_date - DateOffset(years=1)
df_recent_1y = df[df['Reported_Date'] >= one_year_ago]

# 月別件数（全体＋機種別）
monthly_by_type = (
    df_recent_1y.groupby(['YearMonth', 'Aircraft_Type'])
    .size()
    .reset_index(name='Count')
    .pivot(index='YearMonth', columns='Aircraft_Type', values='Count')
    .fillna(0)
    .reset_index()
)
monthly_by_type['Total_Count'] = monthly_by_type[['A350-900', 'A350-1000']].sum(axis=1)

st.subheader("📊 A350全体・機種別 月別不具合件数推移")
exclude_seat = st.checkbox("Seat/IFE/Wi-Fiを除く")

exclude_patterns = ["2520", "2521", "2528"] + [f"442{i}" for i in range(10)] + [f"443{i}" for i in range(10)]

if exclude_seat:
    df_filtered = df_recent_1y[
        (~df_recent_1y['ATA_SubChapter'].isin(exclude_patterns)) &
        (~df_recent_1y.apply(is_seat_related, axis=1))
    ]
else:
    df_filtered = df_recent_1y.copy()

monthly_excl = (
    df_filtered.groupby(['YearMonth', 'Aircraft_Type'])
    .size()
    .reset_index(name='Count')
    .pivot(index='YearMonth', columns='Aircraft_Type', values='Count')
    .fillna(0)
    .reset_index()
)
monthly_excl['Total_Count'] = monthly_excl[['A350-900', 'A350-1000']].sum(axis=1)

fig_total = px.line(
    monthly_by_type,
    x='YearMonth',
    y=['A350-900', 'A350-1000', 'Total_Count'],
    markers=True,
    title="A350全体、A350-900、A350-1000の月別不具合件数推移",
    labels={'value': '件数', 'YearMonth': '年月', 'variable': '機種'}
)
fig_total.update_layout(xaxis=dict(type='category'))

if exclude_seat:
    for col in ['A350-900', 'A350-1000', 'Total_Count']:
        fig_total.add_trace(go.Scatter(
            x=monthly_excl['YearMonth'],
            y=monthly_excl[col],
            mode='lines+markers',
            name=f"{col}（除くSeat/IFE）",
            line=dict(dash='dot')
        ))

st.plotly_chart(fig_total, use_container_width=True)

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
            short_term_rate = short_term_rate.dropna()

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

