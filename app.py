import streamlit as st
import yaml
import os
import json
import re
import gspread
from google.oauth2.service_account import Credentials
from typing import Optional

# ============================================================
# ページ設定 (必ず最初に呼ぶ)
# ============================================================
st.set_page_config(
    page_title="富士ミネラル向けタリフ",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# カスタム CSS (スマホ・PC 両対応)
# ============================================================
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }

    .fare-result-box {
        font-size: 3.5rem;
        font-weight: bold;
        color: #0d6efd;
        text-align: center;
        padding: 28px 20px;
        background: linear-gradient(135deg, #e8f4fd, #f0f8ff);
        border-radius: 16px;
        border: 2px solid #0d6efd;
        margin: 16px 0;
        box-shadow: 0 4px 12px rgba(13,110,253,0.15);
        letter-spacing: 0.04em;
    }

    .ref-box {
        background-color: #f8f9fa;
        padding: 14px 18px;
        border-radius: 10px;
        border-left: 5px solid #20c997;
        margin: 6px 0;
        font-size: 1rem;
    }
    .ref-box b { color: #495057; font-size: 0.85rem; display: block; margin-bottom: 4px; }
    .ref-value { font-size: 1.25rem; font-weight: 600; color: #212529; }

    .year-badge {
        display: inline-block;
        background-color: #0d6efd;
        color: white;
        padding: 4px 14px;
        border-radius: 20px;
        font-size: 0.9rem;
        font-weight: 600;
        margin-bottom: 8px;
    }

    @media (max-width: 768px) {
        .fare-result-box { font-size: 2.4rem; padding: 20px 12px; }
        .ref-value { font-size: 1.1rem; }
        h1 { font-size: 1.5rem !important; }
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# 設定読み込み
# ============================================================
@st.cache_data(show_spinner=False)
def load_config() -> dict:
    config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


# ============================================================
# Google Sheets 認証 (st.cache_resource でシングルトン)
# ============================================================
@st.cache_resource(show_spinner=False)
def get_gspread_client() -> gspread.Client:
    """
    環境変数 GOOGLE_CREDENTIALS (JSON文字列) または
    credentials.json ファイルから認証する。
    """
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    creds_json_str = os.environ.get("GOOGLE_CREDENTIALS", "")
    if creds_json_str:
        try:
            creds_dict = json.loads(creds_json_str)
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            return gspread.authorize(creds)
        except json.JSONDecodeError as e:
            st.error(f"GOOGLE_CREDENTIALS の JSON パースに失敗しました: {e}")
            st.stop()

    creds_file = os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")
    if os.path.exists(creds_file):
        creds = Credentials.from_service_account_file(creds_file, scopes=scopes)
        return gspread.authorize(creds)

    st.error(
        "Google API 認証情報が見つかりません。\n\n"
        "以下のいずれかを設定してください:\n"
        "- 環境変数 `GOOGLE_CREDENTIALS` にサービスアカウントJSONを文字列で設定\n"
        "- 環境変数 `GOOGLE_CREDENTIALS_FILE` にJSONファイルのパスを設定\n"
        "- プロジェクトルートに `credentials.json` を配置"
    )
    st.stop()


# ============================================================
# ユーティリティ
# ============================================================
def _normalize_city(name: str) -> str:
    """全角・半角スペースを全て除去して比較用に正規化する。"""
    return name.replace(" ", "").replace("　", "")


def _parse_number(text: str) -> Optional[float]:
    """
    "20 kg"、"1,500" など単位・記号混じりの文字列から数値を抽出する。
    変換不能な場合は None を返す。
    """
    if not text or not text.strip():
        return None
    normalized = text.strip().translate(
        str.maketrans("０１２３４５６７８９，．", "0123456789,.")
    )
    match = re.search(r"-?[\d,]+\.?\d*", normalized)
    if not match:
        return None
    try:
        return float(match.group().replace(",", ""))
    except ValueError:
        return None


# ============================================================
# データ取得 (OKTable: A=都市名, B=重量, C=運賃 のフラットリスト)
# ============================================================
@st.cache_data(ttl=3600, show_spinner="スプレッドシートからデータを取得しています...")
def load_fare_data(spreadsheet_id: str, sheet_name: str) -> tuple:
    """
    OKTable シートを読み込み、運賃テーブルを構築する。

    Returns:
        unique_cities : ユニークな都市名リスト（正規化済み、出現順）
        weights       : 重量リスト（昇順・重複なし）
        fare_table    : dict[正規化都市名][重量(float)] = 運賃(float)
    """
    gc = get_gspread_client()

    # --- シート取得 ---
    try:
        ss = gc.open_by_key(spreadsheet_id)
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(
            f"スプレッドシートが見つかりません。\n"
            f"ID: `{spreadsheet_id}`\n\n"
            "config.yaml のIDとサービスアカウントの共有設定を確認してください。"
        )
        st.stop()

    try:
        ws = ss.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        available = [w.title for w in ss.worksheets()]
        st.error(
            f"シート「{sheet_name}」が見つかりません。\n\n"
            f"利用可能なシート: {', '.join(available)}"
        )
        st.stop()

    # --- 全行取得 ---
    all_rows = ws.get_all_values()

    if not all_rows:
        st.error(f"シート「{sheet_name}」にデータがありません。")
        st.stop()

    # --- ヘッダー行をスキップ (B列が数値に変換できない行を飛ばす) ---
    fare_table: dict[str, dict[float, float]] = {}
    seen_cities: list[str] = []
    seen_set: set[str] = set()

    for row in all_rows:
        # 列数が 3 未満の行はスキップ
        if len(row) < 3:
            continue

        city_raw  = row[0].strip()
        weight_raw = row[1].strip()
        fare_raw   = row[2].strip()

        # 都市名が空 → スキップ
        if not city_raw:
            continue

        # 重量・運賃が数値でない → ヘッダー行などとみなしてスキップ
        weight = _parse_number(weight_raw)
        fare   = _parse_number(fare_raw)
        if weight is None or fare is None:
            continue

        # 都市名を正規化（スペース除去）してキーとして使用
        city = _normalize_city(city_raw)

        if city not in seen_set:
            seen_cities.append(city)
            seen_set.add(city)
            fare_table[city] = {}

        fare_table[city][weight] = fare

    if not fare_table:
        st.error(
            f"シート「{sheet_name}」から有効なデータを読み込めませんでした。\n\n"
            "A列=都市名、B列=重量(数値)、C列=運賃(数値) の形式を確認してください。"
        )
        st.stop()

    # 全都市共通の重量リストを昇順で作成
    all_weights = sorted({w for city_data in fare_table.values() for w in city_data})

    return seen_cities, all_weights, fare_table


# ============================================================
# 検索ロジック
# ============================================================
def find_weight_ceiling(input_weight: float, weights: list[float]) -> Optional[float]:
    """入力重量以上の最小値（切り上げ）を返す。"""
    candidates = [w for w in weights if w >= input_weight]
    return min(candidates) if candidates else None


def match_city(normalized_input: str, city_list: list[str]) -> Optional[str]:
    """
    都市名を以下の優先順位で照合して返す。difflib は使用しない。
      1. 完全一致
      2. 前方一致（入力が都市名の先頭と一致するもの、最初の1件）
    どちらも該当しなければ None を返す。
    """
    # 1. 完全一致
    if normalized_input in city_list:
        return normalized_input
    # 2. 前方一致（入力文字列で始まる都市名を探す）
    for city in city_list:
        if city.startswith(normalized_input):
            return city
    return None


# ============================================================
# サイドバー
# ============================================================
config = load_config()
spreadsheets_cfg = config.get("spreadsheets", [])

with st.sidebar:
    st.title("⚙️ 設定")
    st.markdown("---")

    if not spreadsheets_cfg:
        st.error("config.yaml にスプレッドシートが設定されていません。")
        st.stop()

    year_names = [s["name"] for s in spreadsheets_cfg]
    selected_year_name = st.selectbox("📅 参照する年度", year_names)

    selected_cfg = next(s for s in spreadsheets_cfg if s["name"] == selected_year_name)
    spreadsheet_id = selected_cfg["id"]
    sheet_name_cfg = selected_cfg.get("sheet_name", "OKTable")

    st.markdown("---")
    st.markdown("**現在の参照年度**")
    st.markdown(f'<div class="year-badge">📋 {selected_year_name}</div>', unsafe_allow_html=True)
    st.caption("データは取得後 1 時間キャッシュされます。")

    if st.button("🔄 キャッシュを更新", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ============================================================
# データ取得
# ============================================================
city_list, weights, fare_table = load_fare_data(
    spreadsheet_id=spreadsheet_id,
    sheet_name=sheet_name_cfg,
)

# ============================================================
# メイン UI
# ============================================================
st.title("💧 富士ミネラル向けタリフ")
st.markdown(
    f'<span class="year-badge">参照タリフ: {selected_year_name}</span>',
    unsafe_allow_html=True,
)
st.markdown("---")

# --- 入力フォーム ---
col_city, col_weight, col_btn = st.columns([3, 2, 1])

with col_city:
    city_input = st.text_input(
        "🌏 行先（都市名）",
        placeholder="例: 上海、バンコク、ロサンゼルス",
        help="都市名を正確に入力してください（全角・半角スペースは無視されます）。",
    )

with col_weight:
    weight_input = st.number_input(
        "⚖️ 重量 (kg)",
        min_value=0.0,
        max_value=99999.0,
        value=0.0,
        step=0.5,
        format="%.1f",
        help="入力値以上の最小重量区分を自動的に参照します。",
    )

with col_btn:
    st.markdown("<br>", unsafe_allow_html=True)
    search_clicked = st.button("🔍 検索", type="primary", use_container_width=True)

# --- 検索実行 ---
if search_clicked:
    if not city_input or weight_input <= 0:
        st.warning("行先と重量（0より大きい値）を入力してください。")
    else:
        normalized_input = _normalize_city(city_input)
        matched_city = match_city(normalized_input, city_list)

        if matched_city is None:
            st.error(f"「{city_input}」はタリフに存在しません。都市名を正確に入力してください。")

        else:
            # 前方一致で補完された場合は通知
            if matched_city != normalized_input:
                st.caption(f"「{city_input}」を「{matched_city}」として検索しました。")

            matched_weight = find_weight_ceiling(weight_input, weights)

            if matched_weight is None:
                st.error(
                    f"重量 **{weight_input} kg** 以上の運賃データがありません。\n\n"
                    f"このタリフの最大重量: **{max(weights):g} kg**"
                )

            else:
                fare = fare_table[matched_city].get(matched_weight)

                if fare is None:
                    st.error(
                        f"「{matched_city}」× **{matched_weight:g} kg** の運賃データが見つかりません。\n\n"
                        "OKTable のデータを確認してください。"
                    )
                else:
                    st.markdown("---")
                    st.subheader("📊 検索結果")

                    ref_col1, ref_col2 = st.columns(2)
                    with ref_col1:
                        st.markdown(
                            f'<div class="ref-box">'
                            f'<b>参照した都市名</b>'
                            f'<span class="ref-value">{matched_city}</span>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                    with ref_col2:
                        weight_note = f"（入力: {weight_input:g} kg → 切り上げ）" \
                                      if weight_input != matched_weight else ""
                        st.markdown(
                            f'<div class="ref-box">'
                            f'<b>参照した重量 {weight_note}</b>'
                            f'<span class="ref-value">{matched_weight:g} kg</span>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )

                    st.markdown(
                        f'<div class="fare-result-box">¥ {int(fare):,}</div>',
                        unsafe_allow_html=True,
                    )

                    st.metric(
                        label=f"{normalized_input}  |  {matched_weight:g} kg  |  {selected_year_name}",
                        value=f"¥{int(fare):,}",
                    )

st.caption(f"© 富士ミネラル向けタリフ | 参照タリフ: {selected_year_name}")
