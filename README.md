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
encoding = "utf-8"      # 出力エンコーディング (デフォルト: shift_jis)
date_format = "%Y/%m/%d" # 日付フォーマット (デフォルト: %Y-%m-%d)
```

### date_format の形式指定例

| 設定値 | 出力例 |
|---|---|
| `%Y-%m-%d` | 2025-01-31 |
| `%Y/%m/%d` | 2025/01/31 |
| `%d/%m/%Y` | 31/01/2025 |
| `%Y年%m月%d日` | 2025年1月31日 |

形式指定には [chrono の strftime 記法](https://docs.rs/chrono/latest/chrono/format/strftime/index.html) を使用します。

## 開発者用のビルド方法

Rust のツールチェーンが必要です。以下のコマンドでビルドしてください。

```bash
cargo build --release
```

ビルド後、バイナリは `target/release/xls2csv` に生成されます。

## 使い方

```bash
xlsx2csv [-e <encoding>] <Excelファイル>
```

### オプション

| オプション | 説明 | デフォルト |
|---|---|---|
| `-e`, `--encoding` | 出力エンコーディング | `shift_jis` |

### 例

```bash
# Shift-JIS で出力（デフォルト）
xlsx2csv data.xlsx

# UTF-8 で出力
xlsx2csv -e utf-8 data.xlsx

# EUC-JP で出力
xlsx2csv --encoding euc-jp data.xlsx
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
- 日付セルは `YYYY-MM-DD` 形式で出力（設定ファイルで変更可能）
- カンマ・ダブルクォート・改行を含むセルは RFC 4180 に準拠してエスケープ

## ライセンス

MIT License
