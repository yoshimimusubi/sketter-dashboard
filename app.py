"""
進捗管理＆AI挽回プラン提案ダッシュボード
医療・福祉機関向け Streamlit アプリケーション
"""

import hashlib
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
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# カスタムCSS
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap');
        html, body, [class*="css"] {
            font-family: 'Noto Sans JP', 'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Meiryo', sans-serif;
        }
        #MainMenu, footer, header { visibility: hidden; }
        .block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
        .kpi-card {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px;
            padding: 1.2rem 1.5rem; box-shadow: 0 1px 4px rgba(0, 0, 0, 0.06); text-align: center;
        }
        .kpi-card .kpi-label { font-size: 0.85rem; color: #64748B; font-weight: 500; margin-bottom: 0.3rem; }
        .kpi-card .kpi-value { font-size: 1.8rem; font-weight: 700; color: #0047AB; }
        .kpi-card .kpi-sub { font-size: 0.8rem; color: #94A3B8; margin-top: 0.2rem; }
        .section-header {
            font-size: 1.1rem; font-weight: 700; color: #0047AB; border-left: 4px solid #007BBB;
            padding-left: 0.8rem; margin-top: 1.5rem; margin-bottom: 0.8rem;
        }
        .alert-banner {
            background: #FEF3C7; border: 1px solid #F59E0B; border-radius: 8px;
            padding: 0.8rem 1.2rem; color: #92400E; font-weight: 500; margin-bottom: 1rem;
        }
        [data-testid="stSidebar"] { background: #F8FAFC; }
        [data-testid="stSidebar"] .block-container { padding-top: 1rem; }
        .stButton>button {
            background: #0047AB; color: #FFFFFF; border: none; border-radius: 8px;
            font-weight: 600; padding: 0.6rem 1.5rem; transition: background 0.2s;
        }
        .stButton>button:hover { background: #003580; color: #FFFFFF; }
        .main-title { font-size: 1.6rem; font-weight: 700; color: #0047AB; margin-bottom: 0.2rem; }
        .main-subtitle { font-size: 0.9rem; color: #64748B; margin-bottom: 1.5rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# ヘッダー
# ---------------------------------------------------------------------------
st.markdown('<div class="main-title">進捗管理ダッシュボード</div>', unsafe_allow_html=True)
st.markdown('<div class="main-subtitle">医療・福祉機関向け 進捗管理＆AI挽回プラン提案システム</div>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# サイドバー
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### データ設定")
    data_source = st.radio("データソースを選択", ["Google スプレッドシート", "ファイルアップロード"], horizontal=True)

    sheet_url = None
    uploaded_file = None

    if data_source == "Google スプレッドシート":
        sheet_url = st.text_input("スプレッドシートのURL", placeholder="https://docs.google.com/spreadsheets/d/...")
        sheet_name = st.text_input("シート名（空欄で先頭シート）", placeholder="シート1")
    else:
        uploaded_file = st.file_uploader("進捗データファイル（CSV / Excel）", type=["csv", "xlsx"])
        excel_sheet_name = st.text_input("Excelシート名（空欄で先頭シート）", placeholder="Sheet1")

    st.markdown("---")

    # データ再読み込みボタン — キャッシュをクリアして最新データを取得
    if st.button("🔄 データを再読み込み"):
        st.cache_data.clear()
        # セッションに保存された古いデータハッシュをリセット
        if "last_data_hash" in st.session_state:
            del st.session_state["last_data_hash"]
        st.rerun()

    st.markdown("---")
    st.markdown("### 目標設定")
    target_count = st.number_input("目標件数（件）", min_value=1, value=50, step=1)
    target_amount = st.number_input("目標金額（円）", min_value=10000, value=1_000_000, step=10000, format="%d")

# ---------------------------------------------------------------------------
# 定数とロジック
# ---------------------------------------------------------------------------
HIGH_PRICE_KEYWORDS = ["特別養護老人ホーム", "特養", "病院", "クリニック", "有料老人ホーム", "老人保健施設", "老健"]
MID_PRICE_KEYWORDS = ["デイサービス", "デイ", "ケアハウス"]
COMPLETED_PHASES = ["契約", "契約済み", "契約完了", "受注"]

def extract_spreadsheet_id(url: str) -> str | None:
    patterns = [r"/spreadsheets/d/([a-zA-Z0-9-_]+)", r"^([a-zA-Z0-9-_]{20,})$"]
    for pat in patterns:
        m = re.search(pat, url.strip())
        if m: return m.group(1)
    return None

def load_from_google_sheets(url: str, sheet_name: str = "") -> pd.DataFrame:
    sid = extract_spreadsheet_id(url)
    if not sid:
        raise ValueError("URLを正しく認識できませんでした。")
    export_url = f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv"
    if sheet_name: export_url += f"&sheet={sheet_name}"
    # スプレッドシート側のキャッシュを回避して常に最新データを取得する
    export_url += f"&_={int(datetime.now().timestamp())}"
    import io, requests
    resp = requests.get(export_url)
    resp.raise_for_status()
    return pd.read_csv(io.StringIO(resp.text))

def estimate_unit_price(row) -> int:
    name = str(row.get("取引先名", ""))
    if not name or name == "nan":
        name = str(row.get("法人名", ""))
    for kw in HIGH_PRICE_KEYWORDS:
        if kw in name: return 30_000
    for kw in MID_PRICE_KEYWORDS:
        if kw in name: return 20_000
    return 20_000

def is_completed(phase: str) -> bool:
    return any(cp in str(phase) for cp in COMPLETED_PHASES)

def calculate_probability(row) -> tuple:
    if is_completed(str(row.get("フェーズ", ""))):
        return "確定", 1.0

    kakudo_raw = row.get("契約確度", "")
    kakudo_col = "" if pd.isna(kakudo_raw) else str(kakudo_raw).strip().upper()
    kakudo_col = kakudo_col.translate(str.maketrans('ＡＢＣ', 'ABC'))
    if kakudo_col == "A": return "A", 0.8
    if kakudo_col == "B": return "B", 0.5
    if kakudo_col == "C": return "C", 0.2

    text = f"{row.get('フェーズ','')} {row.get('現状','')} {row.get('フォロー状況','')} {row.get('ご意向','')}"
    if any(k in text for k in ["稟議承認", "契約日", "開始日", "契約書", "最終段階", "理事長", "決裁"]): return "A", 0.8
    if any(k in text for k in ["導入準備", "合意", "稟議を上げる", "契約意思", "会議", "検討予定"]): return "B", 0.5
    if any(k in text for k in ["前向き", "導入前提", "相談"]): return "C", 0.2

    return "その他", 0.0

# ---------------------------------------------------------------------------
# メイン処理
# ---------------------------------------------------------------------------
df = None

if data_source == "Google スプレッドシート":
    if not sheet_url:
        st.info("左側のメニューにGoogleスプレッドシートのURLを入力してください。")
        st.stop()
    try:
        with st.spinner("データ読み込み中..."):
            df = load_from_google_sheets(sheet_url, sheet_name)
    except Exception as e:
        st.error(f"【エラー】読み込み失敗: {e}")
        st.stop()
else:
    if uploaded_file is None:
        st.info("左側のメニューから進捗データファイルをアップロードしてください。")
        st.stop()
    try:
        # ファイルポインタを先頭にリセット（再読み込み時に古いデータが残る問題を修正）
        uploaded_file.seek(0)
        # ファイルの内容をバイト列として読み込み、ハッシュで変更検知
        raw_bytes = uploaded_file.read()
        file_hash = hashlib.md5(raw_bytes).hexdigest()
        # 前回と同じファイルかどうかを検知（デバッグ用にセッションに記録）
        st.session_state["last_data_hash"] = file_hash
        
        import io
        if uploaded_file.name.endswith(".csv"): 
            df = pd.read_csv(io.BytesIO(raw_bytes))
        else: 
            if "excel_sheet_name" in locals() and excel_sheet_name:
                df = pd.read_excel(io.BytesIO(raw_bytes), engine="openpyxl", sheet_name=excel_sheet_name)
            else:
                df = pd.read_excel(io.BytesIO(raw_bytes), engine="openpyxl")
    except Exception as e:
        st.error(f"【エラー】ファイル読み込み失敗: {e}")
        st.stop()

# カラム名の空白や改行を自動で削除してエラーを防ぐ
df.columns = df.columns.astype(str).str.strip().str.replace('\n', '').str.replace('\r', '')

# 前処理（カラムが存在しなくてもエラーにならないように get を使用）
df["最終連絡日"] = pd.to_datetime(df.get("最終連絡日"), errors="coerce")
df["想定単価"] = df.apply(estimate_unit_price, axis=1)

# ABCヨミの計算
df[["ヨミランク", "見込み率"]] = df.apply(calculate_probability, axis=1, result_type="expand")
df["見込み金額"] = df["想定単価"] * df["見込み率"]

# KPI集計
achieved_df = df[df["ヨミランク"] == "確定"]
achieved_count = len(achieved_df)
achieved_amount = int(achieved_df["見込み金額"].sum()) # 確定分

count_A = len(df[df["ヨミランク"] == "A"])
count_B = len(df[df["ヨミランク"] == "B"])
count_C = len(df[df["ヨミランク"] == "C"])
count_gap = len(df[df["ヨミランク"] == "その他"])

amount_A = int(df[df["ヨミランク"] == "A"]["見込み金額"].sum())
amount_B = int(df[df["ヨミランク"] == "B"]["見込み金額"].sum())
amount_C = int(df[df["ヨミランク"] == "C"]["見込み金額"].sum())

total_pipeline = achieved_amount + amount_A + amount_B + amount_C
gap_amount = max(target_amount - total_pipeline, 0)

count_rate = min(round(achieved_count / target_count * 100, 1), 100) if target_count > 0 else 0
amount_rate = min(round(achieved_amount / target_amount * 100, 1), 100) if target_amount > 0 else 0

# ----- KPI カード -----
st.markdown('<div class="section-header">KPI サマリー</div>', unsafe_allow_html=True)
col1, col2, col3, col4 = st.columns(4)
with col1: st.markdown(f'<div class="kpi-card"><div class="kpi-label">達成件数</div><div class="kpi-value">{achieved_count}</div><div class="kpi-sub">目標 {target_count} 件</div></div>', unsafe_allow_html=True)
with col2: st.markdown(f'<div class="kpi-card"><div class="kpi-label">件数達成率</div><div class="kpi-value">{count_rate}%</div><div class="kpi-sub">残り {max(target_count - achieved_count, 0)} 件</div></div>', unsafe_allow_html=True)
with col3: st.markdown(f'<div class="kpi-card"><div class="kpi-label">達成金額</div><div class="kpi-value">{achieved_amount:,}</div><div class="kpi-sub">目標 {target_amount:,} 円</div></div>', unsafe_allow_html=True)
with col4: st.markdown(f'<div class="kpi-card"><div class="kpi-label">金額達成率</div><div class="kpi-value">{amount_rate}%</div><div class="kpi-sub">残り {max(target_amount - achieved_amount, 0):,} 円</div></div>', unsafe_allow_html=True)

# ----- パイプライン積み上げグラフ -----
st.markdown('<div class="section-header">目標に対するパイプライン（見込み到達度）</div>', unsafe_allow_html=True)
fig = go.Figure()
fig.add_trace(go.Bar(y=["金額ベース"], x=[achieved_amount], name=f"確定済　{achieved_count}事業所", marker_color="#0047AB", orientation='h'))
fig.add_trace(go.Bar(y=["金額ベース"], x=[amount_A], name=f"確度A (80%)　{count_A}事業所", marker_color="#007BBB", orientation='h'))
fig.add_trace(go.Bar(y=["金額ベース"], x=[amount_B], name=f"確度B (50%)　{count_B}事業所", marker_color="#4FA8D1", orientation='h'))
fig.add_trace(go.Bar(y=["金額ベース"], x=[amount_C], name=f"確度C (20%)　{count_C}事業所", marker_color="#93C5FD", orientation='h'))
fig.add_trace(go.Bar(y=["金額ベース"], x=[gap_amount], name="不足分 (ギャップ)", marker_color="#E2E8F0", orientation='h'))
fig.update_layout(barmode='stack', height=200, margin=dict(t=30, b=30, l=10, r=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
st.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------------------------
# 今週の行動計画 (Weekly Action Plan)
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">今週の行動計画（AIアクションプラン）</div>', unsafe_allow_html=True)

if st.button("今週の行動計画を生成する"):
    if not GEMINI_API_KEY:
        st.warning("【注意】Gemini APIキーが設定されていません。")
    else:
        try:
            import google.generativeai as genai
            genai.configure(api_key=GEMINI_API_KEY)
            model = genai.GenerativeModel("gemini-2.0-flash")

            pipeline_data = ""
            for rank in ["A", "B", "C"]:
                rank_df = df[df["ヨミランク"] == rank]
                pipeline_data += f"\n【確度 {rank} の案件】計 {len(rank_df)} 件\n"
                for _, row in rank_df.head(5).iterrows():
                    name = str(row.get("取引先名", "")) if str(row.get("取引先名", "")) != "nan" else str(row.get("法人名", ""))
                    pipeline_data += f"- {name} / ご意向: {row.get('ご意向','')} / 現状: {row.get('現状','')} / フォロー状況: {row.get('フォロー状況','')} / チェック項目: {row.get('契約に向けてのチェック項目','')}\n"

            prompt = f"""あなたは医療・福祉機関向けの敏腕営業マネージャーです。
以下の目標ギャップと現在のパイプライン（ヨミ案件）をもとに、「今週チームとしてどう行動すべきか」の具体的な週間アクションプランを提示してください。

【現在の状況】
- 目標金額までの不足分（ギャップ）: {gap_amount:,}円
{pipeline_data}

以下の観点で具体的なプランを生成してください:
1. 今週最も優先して動くべき案件（Aヨミのクロージングなど）とその具体策
2. Bヨミ、Cヨミを次のフェーズに引き上げるためのアプローチ
3. 目標達成に向けたチームへのメッセージ
※絵文字は使わず、プロフェッショナルなトーンで出力してください。
"""
            with st.spinner("今週の行動計画を策定中..."):
                response = model.generate_content(prompt)
            st.markdown("---")
            st.markdown(response.text)
        except Exception as e:
            st.error(f"【エラー】AI連携に失敗しました: {e}")

# ---------------------------------------------------------------------------
# フォロー推奨案件アラート
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">【要確認】フォロー推奨案件</div>', unsafe_allow_html=True)

today = pd.Timestamp(datetime.now().date())
# 【修正箇所】フェーズが存在しない場合でも安全に計算する処理
phase_series = df["フェーズ"] if "フェーズ" in df.columns else pd.Series([""] * len(df), index=df.index)
alert_mask = (df["最終連絡日"].notna() & ((today - df["最終連絡日"]).dt.days >= 7) & ~phase_series.astype(str).apply(is_completed))
alert_df = df[alert_mask].copy()

if alert_df.empty:
    st.info("【情報】現在、7日以上放置されているフォロー推奨案件はありません。")
else:
    alert_df["経過日数"] = (today - alert_df["最終連絡日"]).dt.days
    st.markdown(f'<div class="alert-banner">【重要】最終連絡から7日以上経過している案件が {len(alert_df)} 件あります。早急なフォローをご検討ください。</div>', unsafe_allow_html=True)
    display_cols = ["法人名", "取引先名", "ヨミランク", "フェーズ", "最終連絡日", "経過日数", "現状", "フォロー状況"]
    available_cols = [c for c in display_cols if c in alert_df.columns]
    st.dataframe(alert_df[available_cols].sort_values("経過日数", ascending=False), use_container_width=True, hide_index=True)

st.markdown('<div class="section-header">全案件データ一覧</div>', unsafe_allow_html=True)
with st.expander("データを表示", expanded=False):
    st.dataframe(df, use_container_width=True, hide_index=True)
