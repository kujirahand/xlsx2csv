# xlsx2csv

ExcelファイルをシートごとにCSVファイルへ変換するCLIツールです。  
xlsx / xls / ods などの形式に対応しています。

## 特徴

- シートごとに個別のCSVファイルを生成
- 日付セルは `YYYY-MM-DD` 形式で出力
- カンマ・ダブルクォート・改行を含むセルは RFC 4180 に準拠してエスケープ
- 入力ファイルと同じディレクトリに出力
- Windowsなら、実行ファイルにExcelファイルをドラッグ＆ドロップすることで変換できます。

## ダウンロード

[こちら](https://github.com/kujirahand/xlsx2csv/releases/)から、最新のバイナリをダウンロードできます。

## 設定について

実行ファイルと同じディレクトリまたはカレントディレクトリに `xlsx2csv.toml` を置くことで、デフォルトの動作を変更できます。

```toml
[default]
encoding = "utf-8"       # 出力エンコーディング (デフォルト: shift_jis)
date_format = "%Y/%m/%d" # 日付フォーマット (デフォルト: %Y-%m-%d)
output_format = "tsv"    # 出力形式: csv または tsv (デフォルト: csv)
```

### date_format の形式指定例

| 設定値 | 出力例 |
|---|---|
| `%Y-%m-%d` | 2025-01-31 |
| `%Y/%m/%d` | 2025/01/31 |
| `%d/%m/%Y` | 31/01/2025 |
| `%Y年%m月%d日` | 2025年1月31日 |

形式指定には [chrono の strftime 記法](https://docs.rs/chrono/latest/chrono/format/strftime/index.html) を使用します。

## コマンドラインオプション

コマンドラインオプションは、設定ファイルの内容を上書きします。例えば、`-e utf-8` を指定すると、設定ファイルで `encoding = "shift_jis"` としていても UTF-8 で出力されます。

| コマンドラインオプション | 設定ファイルのキー | デフォルト | 説明 |
|---|---|---|----|
| `-e`, `--encoding` | `encoding` | `shift_jis` | 文字エンコーディング |
| `-d`, `--date_format` | `date_format` | `%Y-%m-%d` | 日付のフォーマット |
| `--tsv` | `output_format` | `csv` | `tsv` を指定するとタブ区切りで出力 |

## 開発者用のビルド方法

Rust のツールチェーンが必要です。以下のコマンドでビルドしてください。

```bash
cargo build --release
```

ビルド後、バイナリは `target/release/xls2csv` に生成されます。

## 使い方

```bash
xlsx2csv [-e <encoding>] [-d <date_format>] [--tsv] <Excelファイル>
```

### オプション

| オプション | 説明 | デフォルト |
|---|---|---|
| `-e`, `--encoding` | 出力エンコーディング | `shift_jis` |
| `-d`, `--date_format` | 日付フォーマット | `%Y-%m-%d` |
| `--tsv` | タブ区切り(TSV)で出力 | (無効) |
| `-h`, `--help` | ヘルプを表示 | - |
| `-u8`, `--utf8` | UTF-8 で出力 (エンコーディング指定のショートカット) | (無効) |

### 例

```bash
# Shift-JIS で出力（デフォルト）
xlsx2csv data.xlsx

# UTF-8 で出力
xlsx2csv -e utf-8 data.xlsx

# EUC-JP で出力
xlsx2csv --encoding euc-jp data.xlsx

# 日付を YYYY/MM/DD 形式で出力
xlsx2csv -d "%Y/%m/%d" data.xlsx

# エンコーディングと日付フォーマットを同時に指定
xlsx2csv -e utf-8 -d "%Y/%m/%d" data.xlsx

# TSV で出力
xlsx2csv --tsv data.xlsx

# TSV かつ UTF-8 で出力
xlsx2csv --tsv -e utf-8 data.xlsx
```

シートごとに `<ファイル名>_<シート名>.csv` という名前でCSVが生成されます。

```
data_Sheet1.csv
data_Sheet2.csv
```

## 仕様

- 対応フォーマット: xlsx, xls, ods など（拡張子から自動判定）
- 出力先: 入力ファイルと同じディレクトリ
- ファイル名: `<入力ファイル名>_<シート名>.csv`
- 出力エンコーディング: デフォルトは Shift-JIS（`-e utf-8` などで変更可能）
- 日付フォーマット: デフォルトは `%Y-%m-%d`（`-d "%Y/%m/%d"` などで変更可能）
- カンマ・ダブルクォート・改行を含むセルは RFC 4180 に準拠してエスケープ

## ライセンス

MIT License
