# fare-search-app 仕様メモ

> このファイルは、AIエージェント（Antigravity / Claude Code など）や
> 将来の自分が、短時間でこのアプリの全体像を把握するための申し送り書です。
> **コードを改修する前に、まずこのファイルと `app.py` を読んでください。**

最終更新: 2026-10-19

---

## 1. このアプリは何か

運送料金表（タリフ）の検索アプリ。**行先（都市名）と重量(kg)を入力すると運賃が表示される。**

- 画面タイトル: 💧 富士ミネラル向けタリフ
- 利用者: 業務での運賃見積もり
- 実装: Python / Streamlit の単一ファイル（`app.py`、約694行）
- 料金データはアプリ内に持たず、**Google スプレッドシートを参照する**

### 起点住所（固定値）

```python
ORIGIN_ADDRESS = "〒409-3611 山梨県西八代郡市川三郷町大塚1125"
```

`app.py` 冒頭のこの定数が、Google Maps で距離を測るときの出発地点。
変更する場合はここ1箇所を修正する。

---

## 2. ファイル構成

| ファイル | 役割 |
|---|---|
| `app.py` | 全ロジック（UI・データ取得・検索・ログ書き込み） |
| `config.yaml` | 参照するスプレッドシートの年度別定義 |
| `requirements.txt` | 依存パッケージ |
| `Procfile` | `web: streamlit run app.py` |
| `.streamlit/config.toml` | Streamlit のサーバ設定 |
| `.gitignore` | キャッシュ・認証情報の除外設定 |

---

## 3. データソース

### 3-1. 料金表シート（`OKTable`）

1行 = 1つの料金セル、という縦持ち形式。

| 列 | 内容 | 必須 |
|---|---|---|
| A | 都市名 | ○ |
| B | 重量(kg) | ○ |
| C | 運賃(円) | ○ |
| D | 距離(km) | 任意 |

読み込み時（`load_fare_data`）に、以下の2つの辞書へ展開される。

```python
fare_table[都市名][重量] = 運賃          # 都市名で引く用
distance_fare_table[距離][重量] = 運賃   # 距離で引く用（D列がある行のみ）
distance_map[都市名] = 距離の生文字列     # 画面表示用
```

- A列が空の行、B列またはC列が数値に変換できない行はスキップされる
- 都市名は `_normalize_city()` で半角・全角スペースを除去して正規化
- 数値は `_parse_number()` で全角数字・カンマを吸収して float 化

### 3-2. 検索ログシート（`検索結果`）

検索するたびに1行追記される。存在しなければ自動作成（1000行×10列）。

| 列 | ヘッダー |
|---|---|
| A | 検索日時 |
| B | 参照タリフ年度 |
| C | 行先 |
| D | マッチ種別 |
| E | 入力重量(kg) |
| F | 適用重量(kg) |
| G | 実走行距離(km) |
| H | 適用タリフ距離(km) |
| I | 運賃(円) |

**マッチ種別に入る値は3種類**: `都市名` / `ログキャッシュ` / `距離タリフ(Google Maps)`

---

## 4. 検索ロジック（最重要）

### 4-1. 重量の決め方（全ケース共通）

**入力値以上で、最も小さい重量区分に切り上げる**（`find_weight_ceiling`）。

例: 区分が 100 / 150 / 200 kg のとき、入力 130kg → **150kg区分**を適用。

該当区分がない（入力が最大区分を超える）場合はエラー表示。

### 4-2. 行先の引き方（3段階フォールバック）

上から順に試し、最初に成功したものを使う。

```
① 都市名マッチ        OKTable の A列に都市名があるか
       ↓ なければ
② ログキャッシュ      「検索結果」シートに過去の距離が残っているか
       ↓ なければ
③ Google Maps API    Distance Matrix で実走行距離を取得
```

**① 都市名マッチ** — `match_city()`
1. 完全一致を探す
2. なければ「入力文字列で始まる都市名」を探す（**シートの出現順で最初にヒットしたもの**）

**② ログキャッシュ** — `load_search_log_cache()`
「検索結果」シートから、**マッチ種別が `距離タリフ(Google Maps)` の行だけ**を読み、
`都市名 → 実走行距離(km)` の辞書を作る。同じ都市名が複数あれば**後勝ち（最新優先）**。
→ 一度 Google Maps を叩いた地名は、2回目以降 API を使わない。API課金の節約が目的。

**③ Google Maps API** — `get_road_distance_km()`
Distance Matrix API（`language=ja`, `units=metric`, タイムアウト10秒）で
`ORIGIN_ADDRESS` から目的地までの道路距離を取得。

### 4-3. 距離の決め方（②③の場合）

重量と同じく、**実距離以上で最も小さいタリフ距離区分に切り上げる**（`find_distance_ceiling`）。
その距離 × 適用重量で `distance_fare_table` から運賃を引く。

---

## 5. キャッシュ構成

| 対象 | 仕組み | 有効期間 |
|---|---|---|
| `config.yaml` の読み込み | `@st.cache_data` | セッション中 |
| gspread クライアント | `@st.cache_resource` | シングルトン |
| 料金表（OKTable） | `@st.cache_data` | **1時間** |
| 検索ログキャッシュ | `@st.cache_data` | **5分** |
| Google Maps の距離 | `@st.cache_data` | **24時間** |

- サイドバーの「🔄 キャッシュを更新」ボタンで全クリア
- ログ書き込み後は `load_search_log_cache.clear()` が自動実行される

---

## 6. 設定・環境変数

### `config.yaml`

```yaml
spreadsheets:
  - name: "2022年度版"
    id: "1-V8rDl_HlAOE03FESWsKwYCBuZC2VmLC66mS01_Yw4c"
    sheet_name: "OKTable"
```

サイドバーのプルダウンで年度を切り替える。

### 環境変数

| 変数名 | 用途 |
|---|---|
| `GOOGLE_CREDENTIALS` | サービスアカウントJSONを**文字列として**設定 |
| `GOOGLE_CREDENTIALS_FILE` | 上記の代わりにJSONファイルのパスを指定 |
| `GOOGLE_MAPS_API_KEY` | Distance Matrix API のキー |

認証の探索順: `GOOGLE_CREDENTIALS` → `GOOGLE_CREDENTIALS_FILE` → `./credentials.json`

必要なスコープ:
- `https://spreadsheets.google.com/feeds`
- `https://www.googleapis.com/auth/drive`

> **注意**: 認証情報ファイルは `.gitignore` で除外済み。絶対にコミットしないこと。

### デプロイ

`Procfile` があるため Railway / Heroku 系を想定。
`.streamlit/config.toml` でポート `8080`、`headless = true` を指定。

---

## 7. 既知の課題

改修時の検討材料。**修正は依頼があったものだけ行うこと。**

| # | 内容 |
|---|---|
| 1 | **前方一致の誤爆リスク** — `match_city()` は入力で始まる都市名を拾うため、短い地名を入れると意図しない都市にヒットする可能性がある（シート出現順で先勝ち） |
| 2 | **ログキャッシュの参照範囲** — `距離タリフ(Google Maps)` 行のみ読むため、`ログキャッシュ` 種別で記録された行は再利用されない。元データを汚さない意図なら正しい挙動 |
| 3 | **config.yaml のプレースホルダ** — 「2025年度版」「テスト」の `id` が `YOUR_SPREADSHEET_ID_...` のまま。選択すると必ずエラーになる |
| 4 | **ポート固定** — `config.toml` で 8080 固定。`$PORT` を渡すホスティングではズレる可能性 |
| 5 | **セキュリティ設定** — `enableCORS = false` / `enableXsrfProtection = false`。公開URLで運用する場合は要検討 |
| 6 | **未使用の依存** — `requirements.txt` に `pandas` があるが `app.py` では import していない |
| 7 | **デバッグコードの残存** — `write_search_log()` の except 節で `traceback` を画面表示している（コミット `a208680` の名残） |
| 8 | **運賃の二重表示** — 大きな `.fare-result-box` と `st.metric` で同じ金額を2回出している |
| 9 | **距離の重複** — 異なる都市が同じ距離値を持つ場合、`distance_fare_table` では後勝ちで上書きされる |

---

## 8. 改修時の注意（AIエージェント向け）

- **3段階フォールバックの順序（都市名 → ログキャッシュ → Google Maps）を変更しないこと。** API課金を抑える設計意図がある
- **重量・距離の「切り上げ」ルールを変更しないこと。** タリフの読み方としてこれが正。四捨五入や切り捨てにしてはいけない
- **`ORIGIN_ADDRESS` を勝手に変更しないこと**
- 「検索結果」シートの**列順とヘッダー名を変更しないこと。** 変更すると `load_search_log_cache()` が列番号で読んでいるため壊れる
- 認証情報・APIキーをコードに直接書き込まないこと
- 作業後は、変更したファイルと実行したコマンドを箇条書きで報告すること
