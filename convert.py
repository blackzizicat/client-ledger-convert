#!/usr/bin/env python3
"""
Access ↔ Notion CSV変換ツール

使い方:
  to_notion : マスターテーブルのIDを名前に変換（Shift-JIS → UTF-8）
  to_access : NotionエクスポートのCSVを名前からIDに変換（UTF-8 → Shift-JIS）
  info      : 設定ファイルの内容と参照テーブルの確認
"""

import csv
import sys
import argparse
from pathlib import Path

try:
    import yaml
except ImportError:
    print("ERROR: pyyamlが必要です。Dockerfileから実行してください。")
    sys.exit(1)


def load_lookup_table(table_path: str, id_col: str, name_col: str, encoding: str):
    """
    ルックアップテーブルを読み込み、ID→名前・名前→IDの辞書を返す。
    """
    id_to_name = {}
    name_to_id = {}
    path = Path(table_path)
    if not path.exists():
        print(f"WARNING: テーブルファイルが見つかりません: {table_path}")
        return id_to_name, name_to_id

    with open(path, encoding=encoding, newline="") as f:
        reader = csv.DictReader(f)
        if id_col not in (reader.fieldnames or []):
            print(f"WARNING: {table_path} にID列「{id_col}」が存在しません。")
            return id_to_name, name_to_id
        if name_col not in (reader.fieldnames or []):
            print(f"WARNING: {table_path} に名前列「{name_col}」が存在しません。")
            return id_to_name, name_to_id

        for row in reader:
            id_val = row[id_col].strip()
            name_val = row[name_col].strip()
            if id_val:
                id_to_name[id_val] = name_val
            if name_val:
                name_to_id[name_val] = id_val

    return id_to_name, name_to_id


def cmd_to_notion(config: dict, input_override: str = None, output_override: str = None):
    """
    IDを名前に変換してUTF-8 CSVを出力する（Access → Notion方向）。
    """
    input_enc = config.get("access_encoding", "cp932")
    output_enc = config.get("notion_encoding", "utf-8-sig")  # BOM付きUTF-8（Excel互換）
    master = input_override or config.get("master_table", "マスターテーブル.csv")
    output = output_override or config.get("output_notion", "マスターテーブル_notion.csv")

    # 各FK列のID→名前辞書を構築
    id_to_name_map = {}
    for fk in config.get("foreign_keys", []):
        col = fk["column"]
        id_to_name, _ = load_lookup_table(
            fk["table"], fk["id_col"], fk["name_col"], input_enc
        )
        id_to_name_map[col] = id_to_name

    # 変換実行
    converted = 0
    skipped = 0
    with open(master, encoding=input_enc, newline="") as fin, \
         open(output, "w", encoding=output_enc, newline="") as fout:

        reader = csv.DictReader(fin)
        writer = csv.DictWriter(fout, fieldnames=reader.fieldnames)
        writer.writeheader()

        for row in reader:
            for col, id_to_name in id_to_name_map.items():
                if col not in row:
                    continue
                val = row[col].strip()
                if not val:
                    continue
                if val in id_to_name:
                    row[col] = id_to_name[val]
                    converted += 1
                else:
                    # IDが見つからない場合はそのまま（数値でなければ既に名前の可能性）
                    skipped += 1
            writer.writerow(row)

    print(f"[完了] Notion用CSV出力: {output}")
    print(f"  変換: {converted}件  / 変換不可（そのまま）: {skipped}件")
    print(f"  エンコーディング: {input_enc} → {output_enc}")


def get_access_columns(master_table_path: str, encoding: str) -> list[str]:
    """
    AccessのマスターテーブルCSVから列名リストを取得する。
    to_access変換時にNotion独自列を除外するために使用する。
    """
    with open(master_table_path, encoding=encoding, newline="") as f:
        reader = csv.reader(f)
        return next(reader)


def _coerce_value(val: str):
    """
    文字列値を数値に変換できる場合は int/float に変換する。
    csv.QUOTE_NONNUMERIC と組み合わせて、数値フィールドをクォートなし、
    文字列フィールドをダブルクォートで囲んで出力するために使用する。
    """
    if val == "":
        return val
    try:
        return int(val)
    except ValueError:
        try:
            return float(val)
        except ValueError:
            return val


def cmd_to_access(config: dict, input_override: str = None, output_override: str = None):
    """
    名前をIDに変換してShift-JIS CSVを出力する（Notion → Access方向）。
    Notionで追加されたAccess側に存在しない列は自動的に除外する。
    """
    input_enc = config.get("notion_encoding", "utf-8-sig")  # BOM付き/なし両対応
    output_enc = config.get("access_encoding", "cp932")
    lookup_enc = output_enc  # ルックアップテーブルはShift-JIS
    access_master = config.get("master_table", "data/マスターテーブル.csv")
    master = input_override or config.get("input_notion", "data/マスターテーブル_notion.csv")
    output = output_override or config.get("output_access", "data/マスターテーブル_access.csv")

    # Accessの列リストを正とし、Notion独自列を除外するために読み取る
    access_columns = get_access_columns(access_master, output_enc)

    # 各FK列の名前→ID辞書を構築
    name_to_id_map = {}
    for fk in config.get("foreign_keys", []):
        col = fk["column"]
        _, name_to_id = load_lookup_table(
            fk["table"], fk["id_col"], fk["name_col"], lookup_enc
        )
        name_to_id_map[col] = name_to_id

    # 変換実行
    converted = 0
    warnings = []

    with open(master, encoding=input_enc, newline="") as fin, \
         open(output, "w", encoding=output_enc, newline="") as fout:

        reader = csv.DictReader(fin)

        # Notionにしかない列を検出してログ表示
        notion_only_cols = [c for c in (reader.fieldnames or []) if c not in access_columns]
        if notion_only_cols:
            print(f"[除外] Notion独自列（Accessに存在しないため除外）: {notion_only_cols}")

        writer = csv.DictWriter(
            fout, fieldnames=access_columns, extrasaction="ignore",
            quoting=csv.QUOTE_NONNUMERIC
        )
        writer.writeheader()

        for row_num, row in enumerate(reader, start=2):
            for col, name_to_id in name_to_id_map.items():
                if col not in row:
                    continue
                val = row[col].strip()
                if not val:
                    continue
                if val in name_to_id:
                    row[col] = name_to_id[val]
                    converted += 1
                elif val.isdigit():
                    # 既にIDの場合はそのまま
                    pass
                else:
                    warnings.append(f"行{row_num} 列「{col}」: 「{val}」に対応するIDが見つかりません")
            writer.writerow({k: _coerce_value(v) for k, v in row.items()})

    if warnings:
        print("[警告] 以下の値はIDに変換できませんでした:")
        for w in warnings:
            print(f"  {w}")
        print()

    print(f"[完了] Access用CSV出力: {output}")
    print(f"  変換: {converted}件 / 警告: {len(warnings)}件")
    print(f"  エンコーディング: {input_enc} → {output_enc}")


def cmd_info(config: dict):
    """
    設定ファイルの内容と参照テーブルのエントリ数を表示する。
    """
    input_enc = config.get("access_encoding", "cp932")
    print("=== 設定情報 ===")
    print(f"  マスターテーブル: {config.get('master_table')}")
    print(f"  Accessエンコーディング: {input_enc}")
    print(f"  Notionエンコーディング: {config.get('notion_encoding', 'utf-8-sig')}")
    print()
    print("=== FK列とルックアップテーブル ===")
    for fk in config.get("foreign_keys", []):
        id_to_name, _ = load_lookup_table(
            fk["table"], fk["id_col"], fk["name_col"], input_enc
        )
        sample = list(id_to_name.items())[:3]
        sample_str = ", ".join(f"{k}→{v}" for k, v in sample)
        print(f"  [{fk['column']}] → {fk['table']} ({fk['name_col']})")
        print(f"    エントリ数: {len(id_to_name)}件  例: {sample_str}")
    print()


def main():
    parser = argparse.ArgumentParser(
        description="Access（Shift-JIS）↔ Notion（UTF-8）CSV変換ツール",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # AccessのマスターテーブルをNotion用に変換（ID→名前, Shift-JIS→UTF-8）
  python convert.py to_notion

  # NotionエクスポートをAccess用に変換（名前→ID, UTF-8→Shift-JIS）
  python convert.py to_access

  # 入出力ファイルを明示する場合
  python convert.py to_notion --input マスターテーブル.csv --output 出力.csv
  python convert.py to_access --input Notionエクスポート.csv --output Access用.csv

  # 設定内容と参照テーブルの確認
  python convert.py info
        """,
    )
    parser.add_argument(
        "mode",
        choices=["to_notion", "to_access", "info"],
        help="to_notion: IDを名前に変換 | to_access: 名前をIDに変換 | info: 設定確認",
    )
    parser.add_argument(
        "--config",
        default="config.yml",
        help="設定ファイルのパス（デフォルト: config.yml）",
    )
    parser.add_argument("--input", help="入力CSVファイルパス（config.ymlの設定を上書き）")
    parser.add_argument("--output", help="出力CSVファイルパス（config.ymlの設定を上書き）")
    args = parser.parse_args()

    config_path = Path(args.config)
    if not config_path.exists():
        print(f"ERROR: 設定ファイルが見つかりません: {args.config}")
        sys.exit(1)

    with open(config_path, encoding="utf-8") as f:
        config = yaml.safe_load(f)

    if args.mode == "to_notion":
        cmd_to_notion(config, args.input, args.output)
    elif args.mode == "to_access":
        cmd_to_access(config, args.input, args.output)
    elif args.mode == "info":
        cmd_info(config)


if __name__ == "__main__":
    main()
