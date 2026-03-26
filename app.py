"""
進捗管理＆AI挽回プラン提案ダッシュボード
医療・福祉機関向け Streamlit アプリケーション
"""

import io
import os
import re
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# 環境変数の読み込み
# ---------------------------------------------------------------------------
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# ---------------------------------------------------------------------------
# ページ設定
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="進捗管理ダッシュボード",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# カスタムCSS
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* ---------- フォント ---------- */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Noto Sans JP', 'Hiragino Sans', 'Hiragino Kaku Gothic ProN',
                     'Meiryo', sans-serif;
    }

    /* ---------- ハンバーガーメニュー & フッター非表示 ---------- */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* ---------- メインコンテナ ---------- */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 1rem;
    }

    /* ---------- KPIカード ---------- */
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.06);
        text-align: center;
    }
    .kpi-card .kpi-label {
        font-size: 0.85rem;
        color: #64748B;
        font-weight: 500;
        margin-bottom: 0.3rem;
    }
    .kpi-card .kpi-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0047AB;
    }
    .kpi-card .kpi-sub {
        font-size: 0.8rem;
        color: #94A3B8;
        margin-top: 0.2rem;
    }

    /* ---------- セクションヘッダー ---------- */
    .section-header {
        font-size: 1.1rem;
        font-weight: 700;
        color: #0047AB;
        border-left: 4px solid #007BBB;
        padding-left: 0.8rem;
        margin-top: 1.5rem;
        margin-bottom: 0.8rem;
    }

    /* ---------- アラートバナー ---------- */
    .alert-banner {
        background: #FEF3C7;
        border: 1px solid #F59E0B;
        border-radius: 8px;
        padding: 0.8rem 1.2rem;
        color: #92400E;
        font-weight: 500;
        margin-bottom: 1rem;
    }

    /* ---------- サイドバー ---------- */
    [data-testid="stSidebar"] {
        background: #F8FAFC;
    }
    [data-testid="stSidebar"] .block-container {
        padding-top: 1rem;
    }

    /* ---------- ボタン ---------- */
    .stButton > button {
        background: #0047AB;
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.6rem 1.5rem;
        transition: background 0.2s;
    }
    .stButton > button:hover {
        background: #003580;
        color: #FFFFFF;
    }

    /* ---------- データフレーム ---------- */
    .stDataFrame {
        border-radius: 8px;
    }

    /* ---------- タイトル ---------- */
    .main-title {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0047AB;
        margin-bottom: 0.2rem;
    }
    .main-subtitle {
        font-size: 0.9rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# ヘッダー
# ---------------------------------------------------------------------------
st.markdown('<div class="main-title">進捗管理ダッシュボード</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="main-subtitle">医療・福祉機関向け 進捗管理＆AI挽回プラン提案システム</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# サイドバー
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### データ設定")

    data_source = st.radio(
        "データソースを選択",
        ["Google スプレッドシート", "ファイルアップロード"],
        horizontal=True,
    )

    sheet_url = None
    uploaded_file = None

    if data_source == "Google スプレッドシート":
        sheet_url = st.text_input(
            "スプレッドシートのURL",
            placeholder="https://docs.google.com/spreadsheets/d/...",
            help="共有設定で『リンクを知っている全員が閲覧可』にしてください",
        )
        sheet_name = st.text_input(
            "シート名（空欄で先頭シート）",
            placeholder="シート1",
            help="読み込むシート名を指定。空欄の場合は最初のシートを使用します",
        )
    else:
        uploaded_file = st.file_uploader(
            "進捗データファイル（CSV / Excel）",
            type=["csv", "xlsx"],
            help="必須カラム: 法人名, 取引先名, フェーズ, メールアドレス, 契約確度, 最終連絡日, フォロー状況, ご意向, 現状, 電話番号, 担当者, サブメールアドレス",
        )

    st.markdown("---")
    st.markdown("### 目標設定")
    target_count = st.number_input("目標件数（件）", min_value=1, value=50, step=1)
    target_amount = st.number_input(
        "目標金額（円）", min_value=10000, value=1_000_000, step=10000, format="%d"
    )

# ---------------------------------------------------------------------------
# 定数: 単価推測ルール
# ---------------------------------------------------------------------------
HIGH_PRICE_KEYWORDS = [
    "特別養護老人ホーム", "特養", "病院", "クリニック",
    "有料老人ホーム", "老人保健施設", "老健",
]
MID_PRICE_KEYWORDS = ["デイサービス", "デイ", "ケアハウス"]

REQUIRED_COLUMNS = [
    "法人名", "取引先名", "フェーズ", "メールアドレス", "契約確度",
    "最終連絡日", "フォロー状況", "ご意向", "現状", "電話番号",
    "担当者", "サブメールアドレス",
]

# 完了とみなすフェーズ
COMPLETED_PHASES = ["契約", "契約済み", "契約完了", "受注"]

# 放置アラート除外（完了系）フェーズ
ALERT_EXCLUDE_PHRASES = COMPLETED_PHASES + ["失注", "辞退", "キャンセル"]

# ---------------------------------------------------------------------------
# ABCヨミ判定キーワード
# ---------------------------------------------------------------------------
YOMI_A_KEYWORDS = [
    "稟議承認", "契約日", "開始日", "契約書", "最終段階",
    "理事長", "決裁", "契約調整", "締結", "内定",
]
YOMI_B_KEYWORDS = [
    "導入準備", "合意", "稟議を上げる", "契約意思", "会議",
    "検討予定", "見積", "予算確保", "稟議中", "稟議",
]
YOMI_C_KEYWORDS = [
    "前向き", "導入前提", "相談", "興味あり", "検討中",
]

YOMI_RATE_MAP = {"A": 0.8, "B": 0.5, "C": 0.2, "その他": 0.0}


# ---------------------------------------------------------------------------
# ユーティリティ関数
# ---------------------------------------------------------------------------
def extract_spreadsheet_id(url: str) -> str | None:
    """Google スプレッドシートのURLからスプレッドシートIDを抽出する。"""
    patterns = [
        r"/spreadsheets/d/([a-zA-Z0-9-_]+)",
        r"^([a-zA-Z0-9-_]{20,})$",
    ]
    for pat in patterns:
        m = re.search(pat, url.strip())
        if m:
            return m.group(1)
    return None


def load_from_google_sheets(url: str, sheet_name: str = "") -> pd.DataFrame:
    """公開/共有されたGoogleスプレッドシートからDataFrameを読み込む。"""
    sid = extract_spreadsheet_id(url)
    if not sid:
        raise ValueError(
            "スプレッドシートのURLを正しく認識できませんでした。"
            "URLが https://docs.google.com/spreadsheets/d/... の形式か確認してください。"
        )
    gid_param = ""
    if sheet_name:
        # シート名が指定された場合、gidではなくCSVエクスポートのsheet引数を使う
        export_url = (
            f"https://docs.google.com/spreadsheets/d/{sid}"
            f"/gviz/tq?tqx=out:csv&sheet={sheet_name}"
        )
    else:
        export_url = (
            f"https://docs.google.com/spreadsheets/d/{sid}"
            f"/gviz/tq?tqx=out:csv"
        )
    try:
        df = pd.read_csv(export_url)
    except Exception as e:
        raise ConnectionError(
            f"スプレッドシートの読み込みに失敗しました。"
            f"共有設定が『リンクを知っている全員が閲覧可』になっているか確認してください。\n"
            f"詳細: {e}"
        )
    return df


def estimate_unit_price(facility_name: str) -> int:
    """取引先名から想定単価を推測する。"""
    for kw in HIGH_PRICE_KEYWORDS:
        if kw in facility_name:
            return 30_000
    for kw in MID_PRICE_KEYWORDS:
        if kw in facility_name:
            return 20_000
    return 20_000  # デフォルト


def classify_yomi(row: pd.Series) -> str:
    """フェーズ・現状・フォロー状況・ご意向のテキストからABCヨミを判定する。"""
    # 契約済みの場合は判定しない（確定扱い）
    phase_str = str(row.get("フェーズ", ""))
    if is_completed(phase_str):
        return "確定"
    # 判定対象テキストを結合
    text = " ".join(
        str(row.get(col, ""))
        for col in ["フェーズ", "現状", "フォロー状況", "ご意向", "契約確度"]
    )
    for kw in YOMI_A_KEYWORDS:
        if kw in text:
            return "A"
    for kw in YOMI_B_KEYWORDS:
        if kw in text:
            return "B"
    for kw in YOMI_C_KEYWORDS:
        if kw in text:
            return "C"
    return "その他"


def is_completed(phase: str) -> bool:
    """フェーズが完了状態かどうかを判定する。"""
    return any(cp in str(phase) for cp in COMPLETED_PHASES)


def is_alert_target(phase: str) -> bool:
    """フェーズがアラート対象（未完了かつ非離脱）かどうかを判定する。"""
    return not any(ex in str(phase) for ex in ALERT_EXCLUDE_PHRASES)


def make_gauge(value: float, title: str, color: str = "#0047AB") -> go.Figure:
    """Plotly ゲージチャートを生成する。"""
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=value,
            number={"suffix": "%", "font": {"size": 36, "color": "#334155"}},
            title={"text": title, "font": {"size": 14, "color": "#64748B"}},
            gauge={
                "axis": {"range": [0, 100], "tickwidth": 1, "tickcolor": "#CBD5E1"},
                "bar": {"color": color},
                "bgcolor": "#F1F5F9",
                "borderwidth": 0,
                "steps": [
                    {"range": [0, 50], "color": "#E2E8F0"},
                    {"range": [50, 80], "color": "#BFDBFE"},
                    {"range": [80, 100], "color": "#93C5FD"},
                ],
                "threshold": {
                    "line": {"color": "#DC2626", "width": 3},
                    "thickness": 0.8,
                    "value": 100,
                },
            },
        )
    )
    fig.update_layout(
        height=250,
        margin=dict(l=30, r=30, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        font={"family": "Noto Sans JP, sans-serif"},
    )
    return fig


# ---------------------------------------------------------------------------
# メインコンテンツ
# ---------------------------------------------------------------------------
df = None

if data_source == "Google スプレッドシート":
    if not sheet_url:
        st.markdown("---")
        st.info("左側のメニューにGoogleスプレッドシートのURLを入力してください。")
        st.stop()
    try:
        with st.spinner("Googleスプレッドシートからデータを読み込み中..."):
            df = load_from_google_sheets(sheet_url, sheet_name if sheet_name else "")
    except Exception as e:
        st.error(f"【エラー】スプレッドシートの読み込みに失敗しました: {e}")
        st.stop()
else:
    if uploaded_file is None:
        st.markdown("---")
        st.info("左側のメニューから進捗データファイルをアップロードしてください。")
        st.stop()
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"【エラー】ファイルの読み込みに失敗しました: {e}")
        st.stop()

# ----- 必須カラム チェック -----
missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
if missing:
    st.error(f"【エラー】必須カラムが不足しています: {', '.join(missing)}")
    st.stop()

# ----- データ前処理 -----
df["最終連絡日"] = pd.to_datetime(df["最終連絡日"], errors="coerce")
df["想定単価"] = df["取引先名"].astype(str).apply(estimate_unit_price)
df["想定金額"] = df["想定単価"]  # 1件あたりの金額

# ----- ABCヨミ判定 -----
df["確度ランク"] = df.apply(classify_yomi, axis=1)
df["見込み率"] = df["確度ランク"].map({"A": 0.8, "B": 0.5, "C": 0.2, "その他": 0.0, "確定": 1.0})
df["見込み金額"] = (df["想定単価"] * df["見込み率"]).astype(int)

# ----- KPI集計 -----
completed_mask = df["フェーズ"].astype(str).apply(is_completed)
completed_df = df[completed_mask]
achieved_count = len(completed_df)
achieved_amount = int(completed_df["想定金額"].sum())

count_rate = min(round(achieved_count / target_count * 100, 1), 100) if target_count > 0 else 0
amount_rate = min(round(achieved_amount / target_amount * 100, 1), 100) if target_amount > 0 else 0

# ----- KPI カード -----
st.markdown('<div class="section-header">KPI サマリー</div>', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(
        f"""<div class="kpi-card">
            <div class="kpi-label">達成件数</div>
            <div class="kpi-value">{achieved_count}</div>
            <div class="kpi-sub">目標 {target_count} 件</div>
        </div>""",
        unsafe_allow_html=True,
    )
with col2:
    st.markdown(
        f"""<div class="kpi-card">
            <div class="kpi-label">件数達成率</div>
            <div class="kpi-value">{count_rate}%</div>
            <div class="kpi-sub">残り {max(target_count - achieved_count, 0)} 件</div>
        </div>""",
        unsafe_allow_html=True,
    )
with col3:
    st.markdown(
        f"""<div class="kpi-card">
            <div class="kpi-label">達成金額</div>
            <div class="kpi-value">{achieved_amount:,}</div>
            <div class="kpi-sub">目標 {target_amount:,} 円</div>
        </div>""",
        unsafe_allow_html=True,
    )
with col4:
    st.markdown(
        f"""<div class="kpi-card">
            <div class="kpi-label">金額達成率</div>
            <div class="kpi-value">{amount_rate}%</div>
            <div class="kpi-sub">残り {max(target_amount - achieved_amount, 0):,} 円</div>
        </div>""",
        unsafe_allow_html=True,
    )

# ----- パイプライン積み上げ棒グラフ -----
st.markdown(
    '<div class="section-header">パイプライン (目標金額に対する見込み構成)</div>',
    unsafe_allow_html=True,
)

# 確度別の見込み金額を集計
amount_a = int(df[df["確度ランク"] == "A"]["見込み金額"].sum())
amount_b = int(df[df["確度ランク"] == "B"]["見込み金額"].sum())
amount_c = int(df[df["確度ランク"] == "C"]["見込み金額"].sum())
confirmed_amount = achieved_amount  # 確定済み = 契約済みの想定金額合計

pipeline_total = confirmed_amount + amount_a + amount_b + amount_c
gap = max(target_amount - pipeline_total, 0)

pipeline_fig = go.Figure()
pipeline_fig.add_trace(go.Bar(
    name=f"確定済み ({confirmed_amount:,}円)",
    x=[confirmed_amount], y=["目標到達度"],
    orientation="h",
    marker_color="#0047AB",
    text=[f"{confirmed_amount:,}"], textposition="inside",
    textfont={"color": "white", "size": 12},
))
pipeline_fig.add_trace(go.Bar(
    name=f"確度A ({amount_a:,}円)",
    x=[amount_a], y=["目標到達度"],
    orientation="h",
    marker_color="#3B82F6",
    text=[f"{amount_a:,}"] if amount_a > 0 else [""],
    textposition="inside",
    textfont={"color": "white", "size": 12},
))
pipeline_fig.add_trace(go.Bar(
    name=f"確度B ({amount_b:,}円)",
    x=[amount_b], y=["目標到達度"],
    orientation="h",
    marker_color="#60A5FA",
    text=[f"{amount_b:,}"] if amount_b > 0 else [""],
    textposition="inside",
    textfont={"color": "white", "size": 12},
))
pipeline_fig.add_trace(go.Bar(
    name=f"確度C ({amount_c:,}円)",
    x=[amount_c], y=["目標到達度"],
    orientation="h",
    marker_color="#93C5FD",
    text=[f"{amount_c:,}"] if amount_c > 0 else [""],
    textposition="inside",
    textfont={"color": "#334155", "size": 12},
))
if gap > 0:
    pipeline_fig.add_trace(go.Bar(
        name=f"不足分 ({gap:,}円)",
        x=[gap], y=["目標到達度"],
        orientation="h",
        marker_color="#E2E8F0",
        text=[f"{gap:,}"], textposition="inside",
        textfont={"color": "#64748B", "size": 12},
    ))

pipeline_fig.update_layout(
    barmode="stack",
    height=120,
    margin=dict(l=10, r=10, t=10, b=10),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font={"family": "Noto Sans JP, sans-serif"},
    legend=dict(
        orientation="h", yanchor="bottom", y=1.02,
        xanchor="left", x=0, font={"size": 11},
    ),
    xaxis=dict(showticklabels=False, showgrid=False, zeroline=False),
    yaxis=dict(showticklabels=False, showgrid=False, zeroline=False),
)
st.plotly_chart(pipeline_fig, use_container_width=True)

# ----- 確度別件数カード -----
yomi_col1, yomi_col2, yomi_col3, yomi_col4 = st.columns(4)
count_a = len(df[df["確度ランク"] == "A"])
count_b = len(df[df["確度ランク"] == "B"])
count_c = len(df[df["確度ランク"] == "C"])
count_other = len(df[df["確度ランク"] == "その他"])
with yomi_col1:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-label">確度A (80%)</div>'
        f'<div class="kpi-value" style="color:#0047AB">{count_a}件</div>'
        f'<div class="kpi-sub">見込み {amount_a:,}円</div></div>',
        unsafe_allow_html=True,
    )
with yomi_col2:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-label">確度B (50%)</div>'
        f'<div class="kpi-value" style="color:#3B82F6">{count_b}件</div>'
        f'<div class="kpi-sub">見込み {amount_b:,}円</div></div>',
        unsafe_allow_html=True,
    )
with yomi_col3:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-label">確度C (20%)</div>'
        f'<div class="kpi-value" style="color:#60A5FA">{count_c}件</div>'
        f'<div class="kpi-sub">見込み {amount_c:,}円</div></div>',
        unsafe_allow_html=True,
    )
with yomi_col4:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-label">その他</div>'
        f'<div class="kpi-value" style="color:#94A3B8">{count_other}件</div>'
        f'<div class="kpi-sub">見込み対象外</div></div>',
        unsafe_allow_html=True,
    )

# ----- ゲージチャート -----
gauge_col1, gauge_col2 = st.columns(2)
with gauge_col1:
    st.plotly_chart(
        make_gauge(count_rate, "件数 達成率", "#0047AB"),
        use_container_width=True,
    )
with gauge_col2:
    st.plotly_chart(
        make_gauge(amount_rate, "金額 達成率", "#007BBB"),
        use_container_width=True,
    )

# ---------------------------------------------------------------------------
# フォロー推奨案件アラート
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="section-header">【要確認】フォロー推奨案件</div>',
    unsafe_allow_html=True,
)

today = pd.Timestamp(datetime.now().date())
alert_mask = (
    df["最終連絡日"].notna()
    & ((today - df["最終連絡日"]).dt.days >= 7)
    & df["フェーズ"].astype(str).apply(is_alert_target)
)
alert_df = df[alert_mask].copy()

if alert_df.empty:
    st.success("【情報】現在、フォロー推奨案件はありません。")
else:
    alert_df["経過日数"] = (today - alert_df["最終連絡日"]).dt.days
    display_cols = ["法人名", "取引先名", "フェーズ", "契約確度", "最終連絡日", "経過日数", "現状", "フォロー状況", "担当者"]
    available_cols = [c for c in display_cols if c in alert_df.columns]

    st.markdown(
        f'<div class="alert-banner">【重要】最終連絡から7日以上経過している案件が {len(alert_df)} 件あります。早急なフォローをご検討ください。</div>',
        unsafe_allow_html=True,
    )
    st.dataframe(
        alert_df[available_cols].sort_values("経過日数", ascending=False),
        use_container_width=True,
        hide_index=True,
    )

    # ------------------------------------------------------------------
    # AI 挽回プラン提案
    # ------------------------------------------------------------------
    st.markdown(
        '<div class="section-header">AI 挽回プラン提案</div>',
        unsafe_allow_html=True,
    )

    if st.button("AIに挽回プランを相談する"):
        if not GEMINI_API_KEY or GEMINI_API_KEY == "your_api_key_here":
            st.warning(
                "【注意】Gemini APIキーが設定されていません。`.env` ファイルに `GEMINI_API_KEY` を設定してください。"
            )
        else:
            try:
                import google.generativeai as genai

                genai.configure(api_key=GEMINI_API_KEY)
                model = genai.GenerativeModel("gemini-2.0-flash")

                # プロンプト作成
                alert_summary_lines = []
                for _, row in alert_df.head(10).iterrows():
                    alert_summary_lines.append(
                        f"- 法人名: {row.get('法人名', '不明')} / 取引先名: {row.get('取引先名', '不明')} / "
                        f"フェーズ: {row['フェーズ']} / 契約確度: {row.get('契約確度', '不明')} / "
                        f"経過日数: {row['経過日数']}日 / 現状: {row.get('現状', '不明')} / "
                        f"フォロー状況: {row.get('フォロー状況', '不明')} / ご意向: {row.get('ご意向', '不明')}"
                    )
                alert_text = "\n".join(alert_summary_lines)

                prompt = f"""あなたは医療・福祉機関向けの営業コンサルタントです。
以下のKPI達成状況とフォローが滞っている案件情報をもとに、
目標を達成するための具体的な次のアクション（挽回プラン）を提案してください。

【KPI達成状況】
- 件数目標: {target_count}件 / 達成: {achieved_count}件（達成率 {count_rate}%）
- 金額目標: {target_amount:,}円 / 達成: {achieved_amount:,}円（達成率 {amount_rate}%）
- 残り必要件数: {max(target_count - achieved_count, 0)}件
- 残り必要金額: {max(target_amount - achieved_amount, 0):,}円

【フォロー推奨案件（最終連絡から7日以上経過）】
{alert_text}

以下の観点で具体的なアクションプランを提案してください:
1. 各案件への具体的なアプローチ方法（電話、訪問、資料送付など）
2. 優先順位の提案（どの案件から着手すべきか）
3. トークスクリプトやメール文案のヒント
4. 目標達成に向けた全体戦略
"""

                with st.spinner("AIが挽回プランを作成中です..."):
                    response = model.generate_content(prompt)

                st.markdown("---")
                st.markdown("#### 【AI提案】挽回プラン")
                st.markdown(response.text)

            except Exception as e:
                st.error(f"【エラー】AI連携に失敗しました: {e}")

# ---------------------------------------------------------------------------
# 今週の行動計画
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="section-header">今週の行動計画</div>',
    unsafe_allow_html=True,
)

if st.button("今週の行動計画を生成する"):
    if not GEMINI_API_KEY or GEMINI_API_KEY == "your_api_key_here":
        st.warning(
            "【注意】Gemini APIキーが設定されていません。`.env` ファイルに `GEMINI_API_KEY` を設定してください。"
        )
    else:
        try:
            import google.generativeai as genai

            genai.configure(api_key=GEMINI_API_KEY)
            model = genai.GenerativeModel("gemini-2.0-flash")

            # 確度別案件リストの作成
            def _build_case_list(rank: str, max_items: int = 5) -> str:
                subset = df[df["確度ランク"] == rank].head(max_items)
                if subset.empty:
                    return "  該当案件なし"
                lines = []
                for _, r in subset.iterrows():
                    lines.append(
                        f"  - {r.get('法人名', '不明')} / {r.get('取引先名', '不明')} / "
                        f"現状: {r.get('現状', '不明')} / フォロー状況: {r.get('フォロー状況', '不明')} / "
                        f"想定単価: {r.get('想定単価', 0):,}円"
                    )
                return "\n".join(lines)

            cases_a = _build_case_list("A")
            cases_b = _build_case_list("B")
            cases_c = _build_case_list("C")

            action_prompt = f"""あなたは医療・福祉機関向けの営業戦略コンサルタントです。
以下のパイプライン状況を分析し、「今週の具体的な行動計画」を提案してください。

【目標】
- 目標金額: {target_amount:,}円
- 確定済み金額: {confirmed_amount:,}円
- 不足金額: {gap:,}円

【パイプライン構成】
- 確度A (見込み率80%): {count_a}件 / 見込み金額 {amount_a:,}円
{cases_a}
- 確度B (見込み率50%): {count_b}件 / 見込み金額 {amount_b:,}円
{cases_b}
- 確度C (見込み率20%): {count_c}件 / 見込み金額 {amount_c:,}円
{cases_c}

以下の構成で、今週の行動計画を提案してください:
1. 【最優先アクション】不足分を埋めるために今週最も注力すべき確度層（A/B/C）とその理由
2. 【確度A案件への今週のアクション】契約確定に向けた具体的な行動（各案件への電話・訪問・書類対応 等）
3. 【確度B案件への今週のアクション】確度Aに引き上げるための具体的なアプローチ
4. 【確度C案件への今週のアクション】関係構築・情報提供のための行動
5. 【週間スケジュール案】月〜金の曜日別に、どの案件にどうアプローチするかの具体的なスケジュール

箇条書きと具体的な表現で、実行しやすい計画を作成してください。
"""

            with st.spinner("AIが今週の行動計画を作成中です..."):
                action_response = model.generate_content(action_prompt)

            st.markdown("---")
            st.markdown("#### 【AI提案】今週の行動計画")
            st.markdown(action_response.text)

        except Exception as e:
            st.error(f"【エラー】AI連携に失敗しました: {e}")

# ---------------------------------------------------------------------------
# 全データ一覧
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="section-header">全案件データ一覧</div>',
    unsafe_allow_html=True,
)

with st.expander("データを表示", expanded=False):
    st.dataframe(df, use_container_width=True, hide_index=True)
