# Access ↔ Notion CSV変換ツール

Microsoft Access（Shift-JIS）と Notion（UTF-8）のデータベース間でCSVを相互変換するツールです。
Accessでは正規化のためにIDで管理している参照列を、Notionインポート用に名前へ展開し、
Notionエクスポート時には名前をIDに戻してAccessに取り込める形式に変換します。

## ディレクトリ構成

```
client-ledger/
├── Dockerfile
├── docker-compose.yml
├── convert.py        # 変換スクリプト本体
├── config.yml        # 設定ファイル（FK列の対応定義）
└── data/             # CSVファイル置き場
    ├── マスターテーブル.csv
    ├── 種別テーブル.csv
    ├── フロアテーブル.csv
    ├── 建屋テーブル.csv
    ├── 所属テーブル.csv
    ├── 機種テーブル.csv
    ├── メーカーテーブル.csv
    ├── ベンダーテーブル.csv
    ├── 導入時期テーブル.csv
    ├── OS起動種別テーブル.csv
    ├── 4階層図用種別テーブル.csv
    └── ホスト登録テーブル.csv
```

## 前提条件

- Docker および Docker Compose がインストールされていること
- ローカルへのPythonインストールは不要

## セットアップ（初回のみ）

```bash
docker compose build
```

## 使い方

### 設定の確認

変換前にFK列の対応と参照テーブルのエントリ数を確認できます。

```bash
docker compose run --rm converter info
```

出力例：
```
=== 設定情報 ===
  マスターテーブル: data/マスターテーブル.csv
  Accessエンコーディング: cp932
  Notionエンコーディング: utf-8-sig

=== FK列とルックアップテーブル ===
  [種別] → data/種別テーブル.csv (種別)
    エントリ数: 27件  例: 1→情報処理教室パソコン, 2→情報処理教室ノートパソコン, ...
  ...
```

---

### Access → Notion（IDを名前に変換）

AccessのCSVをNotionへインポートできる形式に変換します。

- **入力**: `data/マスターテーブル.csv`（Shift-JIS）
- **出力**: `data/マスターテーブル_notion.csv`（UTF-8）
- FK列のIDが参照先テーブルの名前に置き換わります

```bash
docker compose run --rm converter to_notion
```

ファイルを明示する場合：

```bash
docker compose run --rm converter to_notion \
  --input data/マスターテーブル.csv \
  --output data/出力.csv
```

---

### Notion → Access（名前をIDに変換）

NotionからエクスポートしたCSVをAccessへ取り込める形式に変換します。

- **入力**: `data/マスターテーブル_notion.csv`（UTF-8）
- **出力**: `data/マスターテーブル_access.csv`（Shift-JIS）
- 名前列がIDに戻されます
- **Notionで追加した独自列はAccessに存在しないため自動的に除外されます**

```bash
docker compose run --rm converter to_access
```

ファイルを明示する場合：

```bash
docker compose run --rm converter to_access \
  --input data/Notionエクスポート.csv \
  --output data/Access戻し.csv
```

Notionで「タグ」「担当者」などの列を追加していた場合、変換時に次のようなメッセージが表示されます：

```
[除外] Notion独自列（Accessに存在しないため除外）: ['タグ', '担当者']
```

---

## 設定ファイル（config.yml）

FK列の追加・変更はすべて `config.yml` で管理します。スクリプト本体は変更不要です。

```yaml
# エンコーディング
access_encoding: cp932       # Shift-JIS（変更不要）
notion_encoding: utf-8-sig   # BOM付きUTF-8（Excel互換）。BOMなしにしたい場合は utf-8

# ファイルパス
master_table:  data/マスターテーブル.csv         # to_notion の入力
output_notion: data/マスターテーブル_notion.csv  # to_notion の出力
input_notion:  data/マスターテーブル_notion.csv  # to_access の入力
output_access: data/マスターテーブル_access.csv  # to_access の出力

# FK列の定義
foreign_keys:
  - column: 種別                  # マスターテーブルの列名
    table: data/種別テーブル.csv  # 参照先テーブル
    id_col: ID                    # 参照先のID列名
    name_col: 種別                # 参照先の名前列名
  ...
```

### FK列を追加する場合

新しい参照テーブルが増えた場合は `foreign_keys` にエントリを追加するだけです。

```yaml
foreign_keys:
  # 既存の定義 ...

  # 追加例：「設置室」列が「室テーブル.csv」を参照する場合
  - column: 設置室
    table: data/室テーブル.csv
    id_col: ID
    name_col: 室名
```

## 変換フロー

```
【Access】マスターテーブル.csv（Shift-JIS, IDで管理）
        ↓  docker compose run --rm converter to_notion
【Notion】マスターテーブル_notion.csv（UTF-8, 名前で管理）
        ↓  Notionへインポート・編集・エクスポート
【Notion】エクスポートCSV（UTF-8, Notion独自列含む）
        ↓  docker compose run --rm converter to_access
【Access】マスターテーブル_access.csv（Shift-JIS, IDで管理, 独自列除外済）
```

## 注意事項

- Notionエクスポート時に列名が変わっていないことを確認してください。列名が変わると名前→ID変換が機能しません。
- 参照テーブル側（種別テーブルなど）への変更はAccessで行い、再エクスポートしてから `data/` に上書きしてください。
- IDに変換できなかった値は警告として表示され、元の値がそのまま残ります。
