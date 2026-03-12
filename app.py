import streamlit as st
import pandas as pd
import yaml
import os
import json
import re
import difflib
import gspread
from google.oauth2.service_account import Credentials
from typing import Optional

# ============================================================
# ページ設定 (必ず最初に呼ぶ)
# ============================================================
st.set_page_config(
    page_title="運賃検索システム",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# カスタム CSS (スマホ・PC 両対応)
# ============================================================
st.markdown("""
<style>
    /* ---- 全体 ---- */
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }

    /* ---- 運賃結果ボックス ---- */
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

    /* ---- 参照情報ボックス ---- */
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

    /* ---- 警告/候補ボックス ---- */
    .fuzzy-notice {
        background-color: #fff3cd;
        border-left: 5px solid #ffc107;
        padding: 12px 16px;
        border-radius: 8px;
        margin: 8px 0;
    }

    /* ---- 年度バッジ ---- */
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

    /* ---- スマホ対応 ---- */
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
    ファイルパス GOOGLE_CREDENTIALS_FILE からサービスアカウント認証を行う。
    """
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    # 方法1: 環境変数に JSON 文字列が直接入っている場合
    creds_json_str = os.environ.get("GOOGLE_CREDENTIALS", "")
    if creds_json_str:
        try:
            creds_dict = json.loads(creds_json_str)
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            return gspread.authorize(creds)
        except json.JSONDecodeError as e:
            st.error(f"GOOGLE_CREDENTIALS の JSON パースに失敗しました: {e}")
            st.stop()

    # 方法2: 環境変数にファイルパスが入っている場合
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
# データ取得・解析ロジック
# ============================================================
def _parse_number(text: str) -> Optional[float]:
    """
    "20 kg", "30KG", "1,500", "1500.5" など単位・記号混じりの文字列から
    数値部分だけを正規表現で抽出して float に変換する。
    空文字・ヘッダー文字列・変換不能な値はすべて None を返す（例外を起こさない）。
    """
    if not text or not text.strip():
        return None
    # 全角数字・カンマを半角に正規化
    normalized = text.strip()
    normalized = normalized.translate(str.maketrans("０１２３４５６７８９，．", "0123456789,."))
    # 数字・ドット・カンマ・先頭マイナスのみを抽出（最初にマッチした塊を使用）
    match = re.search(r"-?[\d,]+\.?\d*", normalized)
    if not match:
        return None
    num_str = match.group().replace(",", "")
    try:
        return float(num_str)
    except ValueError:
        return None


@st.cache_data(ttl=3600, show_spinner="スプレッドシートからデータを取得しています...")
def load_fare_data(spreadsheet_id: str, sheet_name: str, sheet_index: int,
                   city_row_start: int, city_row_end: int,
                   weight_row_start: int) -> tuple:
    """
    スプレッドシートから運賃テーブルを構築して返す。

    Returns:
        unique_cities   : ユニークな都市名リスト (列順)
        weights         : 重量リスト (昇順)
        fare_table      : dict[city_name][weight] = fare (float)
        col_to_city     : dict[col_index(0始まり)] = city_name
    """
    gc = get_gspread_client()

    try:
        ss = gc.open_by_key(spreadsheet_id)
        if sheet_name:
            ws = ss.worksheet(sheet_name)
        else:
            ws = ss.get_worksheet(sheet_index)
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"スプレッドシートID `{spreadsheet_id}` が見つかりません。config.yaml を確認してください。")
        st.stop()
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"シート `{sheet_name or sheet_index}` が見つかりません。")
        st.stop()

    # 全セルを取得 (list of list of str)
    all_values: list[list[str]] = ws.get_all_values()

    # ------------------------------------------------------------------
    # 1. 都市名の抽出 (rows city_row_start ~ city_row_end, 1始まり)
    # ------------------------------------------------------------------
    # 0始まりインデックスに変換
    r_start = city_row_start - 1  # e.g. 5-1=4
    r_end = city_row_end          # e.g. 14 (sliceは exclusive なので14のまま)

    city_header_rows = all_values[r_start:r_end]

    # 結合セル対応: 各列について「最初に現れた非空値」を採用
    # col 0 = A列 (重量列) なので col 1 以降が対象
    col_to_city: dict[int, str] = {}
    max_col = max((len(row) for row in city_header_rows), default=0)

    for col_idx in range(1, max_col):
        for row in city_header_rows:
            if col_idx < len(row) and row[col_idx].strip():
                col_to_city[col_idx] = row[col_idx].strip()
                break  # その列の最初の非空値を採用

    # ユニーク都市リスト (列順を保持、重複除去)
    seen: set[str] = set()
    unique_cities: list[str] = []
    for col_idx in sorted(col_to_city.keys()):
        city = col_to_city[col_idx]
        if city not in seen:
            unique_cities.append(city)
            seen.add(city)

    if not unique_cities:
        st.error(
            f"都市名が取得できませんでした。\n"
            f"スプレッドシートの {city_row_start}〜{city_row_end} 行目にデータが存在するか確認してください。"
        )
        st.stop()

    # ------------------------------------------------------------------
    # 2. 重量リストの抽出 (A列, row weight_row_start 以降)
    # ------------------------------------------------------------------
    w_start = weight_row_start - 1  # 0始まり
    weights: list[float] = []
    weight_row_map: list[int] = []  # 各重量が何行目(0始まり)にあるか

    for row_idx in range(w_start, len(all_values)):
        row = all_values[row_idx]
        if not row or not row[0].strip():
            continue
        val = _parse_number(row[0])
        if val is not None:
            weights.append(val)
            weight_row_map.append(row_idx)

    if not weights:
        st.error(
            f"重量データが取得できませんでした。\n"
            f"A列の {weight_row_start} 行目以降にデータが存在するか確認してください。"
        )
        st.stop()

    # ------------------------------------------------------------------
    # 3. 運賃テーブルの構築
    # ------------------------------------------------------------------
    fare_table: dict[str, dict[float, float]] = {city: {} for city in unique_cities}

    for row_idx, weight in zip(weight_row_map, weights):
        row = all_values[row_idx]
        for col_idx, city in col_to_city.items():
            if col_idx < len(row) and row[col_idx].strip():
                fare_val = _parse_number(row[col_idx])
                if fare_val is not None:
                    fare_table[city][weight] = fare_val

    return unique_cities, sorted(weights), fare_table, col_to_city


# ============================================================
# 検索ロジック
# ============================================================
def fuzzy_city_match(input_city: str, city_list: list[str],
                     cutoff: float = 0.4, n: int = 3
                     ) -> tuple[Optional[str], list[str]]:
    """
    都市名の曖昧検索。
    Returns:
        best_match  : 最も近い都市名 (完全一致含む)。見つからなければ None。
        candidates  : 上位候補リスト (best_match を含む)
    """
    if input_city in city_list:
        return input_city, [input_city]

    # 部分一致を先にチェック (前方一致や包含)
    partial = [c for c in city_list if input_city in c or c in input_city]
    if partial:
        return partial[0], partial

    # difflib による類似度検索
    matches = difflib.get_close_matches(input_city, city_list, n=n, cutoff=cutoff)
    if matches:
        return matches[0], matches

    return None, []


def find_weight_ceiling(input_weight: float, weights: list[float]) -> Optional[float]:
    """入力重量以上の最小値（切り上げ）を返す。"""
    candidates = [w for w in weights if w >= input_weight]
    return min(candidates) if candidates else None


# ============================================================
# サイドバー
# ============================================================
config = load_config()
spreadsheets_cfg = config.get("spreadsheets", [])
data_structure = config.get("data_structure", {})

CITY_ROW_START = data_structure.get("city_row_start", 5)
CITY_ROW_END = data_structure.get("city_row_end", 14)
WEIGHT_ROW_START = data_structure.get("weight_row_start", 16)

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
    sheet_name_cfg = selected_cfg.get("sheet_name", "") or ""
    sheet_index_cfg = selected_cfg.get("sheet_index", 0)

    st.markdown("---")
    st.markdown(f"**現在の参照年度**")
    st.markdown(f'<div class="year-badge">📋 {selected_year_name}</div>', unsafe_allow_html=True)
    st.caption("データは取得後 1 時間キャッシュされます。")

    if st.button("🔄 キャッシュを更新", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ============================================================
# データ取得
# ============================================================
city_list, weights, fare_table, col_to_city = load_fare_data(
    spreadsheet_id=spreadsheet_id,
    sheet_name=sheet_name_cfg,
    sheet_index=sheet_index_cfg,
    city_row_start=CITY_ROW_START,
    city_row_end=CITY_ROW_END,
    weight_row_start=WEIGHT_ROW_START,
)

# ============================================================
# メイン UI
# ============================================================
st.title("✈️ 運賃検索システム")
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
        help="部分一致・曖昧検索対応。入力ミスがあっても自動補正します。",
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
    st.markdown("<br>", unsafe_allow_html=True)  # ラベル分の余白
    search_clicked = st.button("🔍 検索", type="primary", use_container_width=True)

# --- 検索実行 ---
if search_clicked:
    if not city_input or weight_input <= 0:
        st.warning("行先と重量（0より大きい値）を入力してください。")
    else:
        matched_city, candidates = fuzzy_city_match(city_input, city_list)
        matched_weight = find_weight_ceiling(weight_input, weights)

        # --- 都市名マッチ結果 ---
        if matched_city is None:
            st.error(f"「{city_input}」に近い都市が見つかりませんでした。")
            with st.expander("利用可能な都市一覧"):
                st.write("、".join(city_list))

        # --- 重量範囲外 ---
        elif matched_weight is None:
            st.error(
                f"重量 **{weight_input} kg** 以上の運賃データがありません。\n\n"
                f"このタリフの最大重量: **{max(weights)} kg**"
            )

        else:
            # --- 曖昧マッチの通知 ---
            if city_input != matched_city:
                other_candidates = [c for c in candidates if c != matched_city]
                notice_html = (
                    f'<div class="fuzzy-notice">'
                    f'「<b>{city_input}</b>」→ <b>{matched_city}</b> に自動補正しました。'
                )
                if other_candidates:
                    notice_html += f"<br>他の候補: {', '.join(other_candidates)}"
                notice_html += "</div>"
                st.markdown(notice_html, unsafe_allow_html=True)

            # --- 運賃取得 ---
            fare = fare_table.get(matched_city, {}).get(matched_weight)

            if fare is None:
                st.error(
                    f"「{matched_city}」× **{matched_weight} kg** の運賃データが見つかりません。\n\n"
                    "スプレッドシートのデータを確認してください。"
                )
            else:
                st.markdown("---")
                st.subheader("📊 検索結果")

                # 参照情報の表示
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
                    weight_note = ""
                    if weight_input != matched_weight:
                        weight_note = f"（入力: {weight_input} kg → 切り上げ）"
                    st.markdown(
                        f'<div class="ref-box">'
                        f'<b>参照した重量 {weight_note}</b>'
                        f'<span class="ref-value">{matched_weight} kg</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

                # 運賃の大きな表示
                st.markdown(
                    f'<div class="fare-result-box">¥ {int(fare):,}</div>',
                    unsafe_allow_html=True,
                )

                # st.metric でも補足表示
                st.metric(
                    label=f"{matched_city}  |  {matched_weight} kg  |  {selected_year_name}",
                    value=f"¥{int(fare):,}",
                )

# ============================================================
# フッター: 都市・重量情報 (折りたたみ)
# ============================================================
st.markdown("---")
col_exp1, col_exp2 = st.columns(2)

with col_exp1:
    with st.expander(f"🌍 利用可能な都市一覧 ({len(city_list)}件)"):
        # 5列グリッドで表示
        chunk_size = max(1, (len(city_list) + 4) // 5)
        cols = st.columns(5)
        for i, city in enumerate(city_list):
            cols[i % 5].write(city)

with col_exp2:
    with st.expander(f"⚖️ 重量区分一覧 ({len(weights)}段階)"):
        w_df = pd.DataFrame({"重量 (kg)": [f"{w:g}" for w in weights]})
        st.dataframe(w_df, hide_index=True, use_container_width=True)

st.caption(f"© 運賃検索システム | 参照タリフ: {selected_year_name}")
