// Excelファイル(xlsx/xls/ods等)をシートごとにCSVへ変換するCLIツール
use calamine::{open_workbook_auto, Data, Range, Reader};
use encoding_rs::Encoding;
use std::env;
use std::fs;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::Path;

// アプリケーション設定
struct Config {
    encoding: String,
    date_format: String,
    output_format: String, // "csv" または "tsv"
}

impl Config {
    // デフォルト値で初期化する
    fn default() -> Self {
        Config {
            encoding: "shift_jis".to_string(),
            date_format: "%Y-%m-%d".to_string(),
            output_format: "csv".to_string(),
        }
    }

    // xlsx2csv.toml を読み込み、設定を上書きする
    // 探索順: 1) 実行ファイルと同じディレクトリ  2) カレントディレクトリ
    fn load_from_file(&mut self) {
        let mut candidates: Vec<std::path::PathBuf> = Vec::new();
        if let Ok(exe_path) = env::current_exe() {
            if let Some(exe_dir) = exe_path.parent() {
                candidates.push(exe_dir.join("xlsx2csv.toml"));
            }
        }
        if let Ok(cwd) = env::current_dir() {
            candidates.push(cwd.join("xlsx2csv.toml"));
        }

        for path in &candidates {
            if let Ok(content) = fs::read_to_string(path) {
                if let Ok(table) = content.parse::<toml::Table>() {
                    if let Some(enc) = table
                        .get("default")
                        .and_then(|v| v.get("encoding"))
                        .and_then(|v| v.as_str())
                    {
                        self.encoding = enc.to_string();
                    }
                    if let Some(fmt) = table
                        .get("default")
                        .and_then(|v| v.get("date_format"))
                        .and_then(|v| v.as_str())
                    {
                        self.date_format = fmt.to_string();
                    }
                    if let Some(fmt) = table
                        .get("default")
                        .and_then(|v| v.get("output_format"))
                        .and_then(|v| v.as_str())
                    {
                        match fmt {
                            "csv" | "tsv" => self.output_format = fmt.to_string(),
                            other => {
                                eprintln!("Warning: unknown output_format '{}', using csv", other);
                            }
                        }
                    }
                    // 今後の設定項目はここに追加する
                }
                break; // 最初に見つかったファイルのみ読む
            }
        }
    }

    // コマンドライン引数で設定を上書きする
    // 戻り値: 入力ファイルパス
    fn apply_args(&mut self, args: &[String]) -> Option<String> {
        let mut input_path_opt: Option<String> = None;
        let mut i = 1;
        while i < args.len() {
            match args[i].as_str() {
                "--encoding" | "-e" => {
                    i += 1;
                    if i < args.len() {
                        self.encoding = args[i].clone();
                    } else {
                        eprintln!("Error: --encoding requires a value");
                        std::process::exit(1);
                    }
                }
                "--date_format" | "-d" => {
                    i += 1;
                    if i < args.len() {
                        self.date_format = args[i].clone();
                    } else {
                        eprintln!("Error: --date_format requires a value");
                        std::process::exit(1);
                    }
                }
                "--tsv" => {
                    self.output_format = "tsv".to_string();
                }
                "--utf8" | "-u8" => {
                    self.encoding = "utf-8".to_string();
                }
                "--help" | "-h" => {
                    print_help();
                    std::process::exit(0);
                }
                arg if !arg.starts_with('-') => {
                    input_path_opt = Some(arg.to_string());
                }
                unknown => {
                    eprintln!("Unknown option: {}", unknown);
                    std::process::exit(1);
                }
            }
            i += 1;
        }
        input_path_opt
    }
}

// ヘルプメッセージを表示する
fn print_help() {
    println!("Usage: xlsx2csv [-e <encoding>] [-d <date_format>] [--tsv] [--utf8] <excel_file>");
    println!();
    println!("オプション:");
    println!("  -e, --encoding    出力エンコーディング (デフォルト: shift_jis)");
    println!("  -d, --date_format 日付フォーマット (デフォルト: %Y-%m-%d)");
    println!("  --tsv             タブ区切り(TSV)で出力する");
    println!("  -u8, --utf8       UTF-8 で出力する (--encoding utf-8 のショートカット)");
    println!("  -h, --help        このヘルプを表示する");
    println!();
    println!("例:");
    println!("  xlsx2csv data.xlsx");
    println!("  xlsx2csv -e utf-8 data.xlsx");
    println!("  xlsx2csv --utf8 --tsv data.xlsx");
    println!("  xlsx2csv -d \"%Y/%m/%d\" data.xlsx");
}

// メイン処理
fn main() {
    let args: Vec<String> = env::args().collect();

    // 設定をロードし、引数で上書きする
    let mut config = Config::default();
    config.load_from_file();
    let input_path = match config.apply_args(&args) {
        Some(p) => p,
        None => {
            print_help();
            std::process::exit(1);
        }
    };

    // エンコーディングを解決する
    let encoding = Encoding::for_label(config.encoding.as_bytes()).unwrap_or_else(|| {
        eprintln!("Unknown encoding: {}", config.encoding);
        std::process::exit(1);
    });

    // 入力ファイルのパスから「拡張子を除いたファイル名」と「親ディレクトリ」を取り出す
    let path = Path::new(&input_path);
    let stem = path
        .file_stem()
        .expect("Invalid input file name")
        .to_string_lossy()
        .to_string();
    let parent = path.parent().unwrap_or_else(|| Path::new(""));

    convert_workbook_to_csv(&input_path, &stem, parent, encoding, &config.date_format, &config.output_format);
}

// Excelファイルを開き、全シートをCSVへ書き出す
fn convert_workbook_to_csv(input_path: &str, stem: &str, parent: &Path, encoding: &'static Encoding, date_format: &str, output_format: &str) {
    // Excelファイルを開く (拡張子から自動でフォーマットを判定)
    let mut workbook = open_workbook_auto(input_path).expect("Cannot open Excel file");
    // 全シート名を取得する
    let sheet_names = workbook.sheet_names().to_owned();

    // シートごとに CSV ファイルを生成する
    for sheet_name in &sheet_names {
        // シートの内容を取得する
        let range = match workbook.worksheet_range(sheet_name) {
            Ok(r) => r,
            Err(e) => {
                // 取得失敗したシートはスキップ
                eprintln!("Skip sheet '{}': {}", sheet_name, e);
                continue;
            }
        };
        // 1シート分をCSVに書き出す
        write_sheet_to_csv(&range, stem, sheet_name, parent, encoding, date_format, output_format);
    }
}

// 1シート分のセル範囲をCSVファイルに書き出す
fn write_sheet_to_csv(range: &Range<Data>, stem: &str, sheet_name: &str, parent: &Path, encoding: &'static Encoding, date_format: &str, output_format: &str) {
    let tsv = output_format == "tsv";
    // 出力ファイルのパスを「入力ファイル名_シート名.csv/tsv」で組み立てる
    let ext = if tsv { "tsv" } else { "csv" };
    let output_filename = format!("{}_{}.{}", stem, sheet_name, ext);
    let output_path = parent.join(&output_filename);
    // 出力ファイルを作成し、バッファ付きで書き込む
    let file = File::create(&output_path).expect("Cannot create output file");
    let mut writer = BufWriter::new(file);

    // 区切り文字を選択する
    let delimiter = if tsv { "\t" } else { "," };

    // 1行ずつセルを文字列化&エスケープし、区切り文字で書き出す
    for row in range.rows() {
        let fields: Vec<String> = row
            .iter()
            .map(|cell| {
                let s = cell_to_string(cell, date_format);
                if tsv { s } else { escape_csv_field(&s) }
            })
            .collect();
        // 指定エンコーディングに変換してバイト列として書き出す
        let line = fields.join(delimiter) + "\n";
        let (encoded, _, _) = encoding.encode(&line);
        writer.write_all(&encoded).expect("Write failed");
    }

    // 出力したファイルパスを表示する
    println!("出力: {} ({})", output_path.display(), encoding.name());
}

// CSVフィールドの値をエスケープする
// カンマ・ダブルクォート・改行を含む場合は "..." で囲み、内部の " は "" に置換する
fn escape_csv_field(s: &str) -> String {
    if s.contains(',') || s.contains('"') || s.contains('\n') || s.contains('\r') {
        let escaped = s.replace('"', "\"\"");
        format!("\"{}\"", escaped)
    } else {
        s.to_string()
    }
}

// Excelの1セル(Data型)を文字列に変換する
fn cell_to_string(cell: &Data, date_format: &str) -> String {
    match cell {
        // 空セル
        Data::Empty => String::new(),
        // 文字列セル
        Data::String(s) => s.clone(),
        // 数値(浮動小数)セル
        Data::Float(f) => f.to_string(),
        // 整数セル
        Data::Int(i) => i.to_string(),
        // 真偽値セル
        Data::Bool(b) => b.to_string(),
        // 日時セル: Excel内部のシリアル値を NaiveDateTime に変換して出力
        Data::DateTime(dt) => match dt.as_datetime() {
            Some(ndt) => ndt.format(date_format).to_string(),
            // 変換失敗時はシリアル値(数値)をそのまま出力する
            None => dt.as_f64().to_string(),
        },
        // ISO形式の日時文字列セル: パースして date_format に整形する
        Data::DateTimeIso(s) => {
            if let Ok(ndt) = chrono::NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S") {
                ndt.format(date_format).to_string()
            } else if let Ok(nd) = chrono::NaiveDate::parse_from_str(s, "%Y-%m-%d") {
                nd.format(date_format).to_string()
            } else {
                // 想定外のフォーマットならそのまま出力
                s.clone()
            }
        }
        // ISO形式の期間文字列セル(例: PT1H30M)はそのまま出力
        Data::DurationIso(s) => s.clone(),
        // エラーセル(#REF! など)はデバッグ表現で出力
        Data::Error(e) => format!("{:?}", e),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::Data;

    // --- escape_csv_field ---

    #[test]
    fn test_escape_csv_field_plain() {
        // 特殊文字なし → そのまま返す
        assert_eq!(escape_csv_field("hello"), "hello");
    }

    #[test]
    fn test_escape_csv_field_with_comma() {
        // カンマを含む → ダブルクォートで囲む
        assert_eq!(escape_csv_field("a,b"), "\"a,b\"");
    }

    #[test]
    fn test_escape_csv_field_with_double_quote() {
        // ダブルクォートを含む → "" にエスケープして囲む
        assert_eq!(escape_csv_field("say \"hi\""), "\"say \"\"hi\"\"\"");
    }

    #[test]
    fn test_escape_csv_field_with_newline() {
        // 改行を含む → ダブルクォートで囲む
        assert_eq!(escape_csv_field("a\nb"), "\"a\nb\"");
    }

    #[test]
    fn test_escape_csv_field_with_cr() {
        // キャリッジリターンを含む → ダブルクォートで囲む
        assert_eq!(escape_csv_field("a\rb"), "\"a\rb\"");
    }

    #[test]
    fn test_escape_csv_field_empty() {
        // 空文字列 → そのまま空文字列
        assert_eq!(escape_csv_field(""), "");
    }

    // --- cell_to_string ---

    #[test]
    fn test_cell_to_string_empty() {
        assert_eq!(cell_to_string(&Data::Empty, "%Y-%m-%d"), "");
    }

    #[test]
    fn test_cell_to_string_string() {
        assert_eq!(cell_to_string(&Data::String("hello".to_string()), "%Y-%m-%d"), "hello");
    }

    #[test]
    fn test_cell_to_string_int() {
        assert_eq!(cell_to_string(&Data::Int(42), "%Y-%m-%d"), "42");
    }

    #[test]
    fn test_cell_to_string_float() {
        assert_eq!(cell_to_string(&Data::Float(3.14), "%Y-%m-%d"), "3.14");
    }

    #[test]
    fn test_cell_to_string_bool_true() {
        assert_eq!(cell_to_string(&Data::Bool(true), "%Y-%m-%d"), "true");
    }

    #[test]
    fn test_cell_to_string_bool_false() {
        assert_eq!(cell_to_string(&Data::Bool(false), "%Y-%m-%d"), "false");
    }

    #[test]
    fn test_cell_to_string_datetime_iso_date() {
        // DateTimeIso: 日付のみ形式
        let result = cell_to_string(&Data::DateTimeIso("2025-01-31".to_string()), "%Y-%m-%d");
        assert_eq!(result, "2025-01-31");
    }

    #[test]
    fn test_cell_to_string_datetime_iso_date_slash_format() {
        // DateTimeIso: スラッシュ形式に変換
        let result = cell_to_string(&Data::DateTimeIso("2025-01-31".to_string()), "%Y/%m/%d");
        assert_eq!(result, "2025/01/31");
    }

    #[test]
    fn test_cell_to_string_datetime_iso_datetime() {
        // DateTimeIso: 日時形式 → 日付部分のみ抽出
        let result = cell_to_string(&Data::DateTimeIso("2025-01-31T12:34:56".to_string()), "%Y-%m-%d");
        assert_eq!(result, "2025-01-31");
    }

    #[test]
    fn test_cell_to_string_datetime_iso_unknown() {
        // DateTimeIso: 解析できない文字列 → そのまま返す
        let result = cell_to_string(&Data::DateTimeIso("not-a-date".to_string()), "%Y-%m-%d");
        assert_eq!(result, "not-a-date");
    }

    #[test]
    fn test_cell_to_string_duration_iso() {
        // DurationIso: そのまま返す
        assert_eq!(cell_to_string(&Data::DurationIso("PT1H30M".to_string()), "%Y-%m-%d"), "PT1H30M");
    }

    // --- Config::default ---

    #[test]
    fn test_config_default_encoding() {
        let c = Config::default();
        assert_eq!(c.encoding, "shift_jis");
    }

    #[test]
    fn test_config_default_date_format() {
        let c = Config::default();
        assert_eq!(c.date_format, "%Y-%m-%d");
    }

    #[test]
    fn test_config_default_output_format() {
        let c = Config::default();
        assert_eq!(c.output_format, "csv");
    }

    // --- Config::apply_args ---

    fn make_args(args: &[&str]) -> Vec<String> {
        std::iter::once("xlsx2csv")
            .chain(args.iter().copied())
            .map(String::from)
            .collect()
    }

    #[test]
    fn test_apply_args_input_file_only() {
        let mut c = Config::default();
        let result = c.apply_args(&make_args(&["data.xlsx"]));
        assert_eq!(result, Some("data.xlsx".to_string()));
    }

    #[test]
    fn test_apply_args_no_args_returns_none() {
        let mut c = Config::default();
        let result = c.apply_args(&make_args(&[]));
        assert_eq!(result, None);
    }

    #[test]
    fn test_apply_args_encoding_long() {
        let mut c = Config::default();
        c.apply_args(&make_args(&["--encoding", "utf-8", "data.xlsx"]));
        assert_eq!(c.encoding, "utf-8");
    }

    #[test]
    fn test_apply_args_encoding_short() {
        let mut c = Config::default();
        c.apply_args(&make_args(&["-e", "euc-jp", "data.xlsx"]));
        assert_eq!(c.encoding, "euc-jp");
    }

    #[test]
    fn test_apply_args_utf8_flag() {
        let mut c = Config::default();
        c.apply_args(&make_args(&["--utf8", "data.xlsx"]));
        assert_eq!(c.encoding, "utf-8");
    }

    #[test]
    fn test_apply_args_utf8_short_flag() {
        let mut c = Config::default();
        c.apply_args(&make_args(&["-u8", "data.xlsx"]));
        assert_eq!(c.encoding, "utf-8");
    }

    #[test]
    fn test_apply_args_date_format_long() {
        let mut c = Config::default();
        c.apply_args(&make_args(&["--date_format", "%Y/%m/%d", "data.xlsx"]));
        assert_eq!(c.date_format, "%Y/%m/%d");
    }

    #[test]
    fn test_apply_args_date_format_short() {
        let mut c = Config::default();
        c.apply_args(&make_args(&["-d", "%d/%m/%Y", "data.xlsx"]));
        assert_eq!(c.date_format, "%d/%m/%Y");
    }

    #[test]
    fn test_apply_args_tsv() {
        let mut c = Config::default();
        c.apply_args(&make_args(&["--tsv", "data.xlsx"]));
        assert_eq!(c.output_format, "tsv");
    }

    #[test]
    fn test_apply_args_combination() {
        // 複数オプションの組み合わせ
        let mut c = Config::default();
        let result = c.apply_args(&make_args(&["--utf8", "--tsv", "-d", "%Y/%m/%d", "data.xlsx"]));
        assert_eq!(result, Some("data.xlsx".to_string()));
        assert_eq!(c.encoding, "utf-8");
        assert_eq!(c.output_format, "tsv");
        assert_eq!(c.date_format, "%Y/%m/%d");
    }

    // --- Config::load_from_file (TOML パース部分の単体検証) ---

    #[test]
    fn test_toml_encoding_parse() {
        let toml = "[default]\nencoding = \"utf-8\"\n";
        let table: toml::Table = toml.parse().unwrap();
        let enc = table.get("default").and_then(|v| v.get("encoding")).and_then(|v| v.as_str());
        assert_eq!(enc, Some("utf-8"));
    }

    #[test]
    fn test_toml_date_format_parse() {
        let toml = "[default]\ndate_format = \"%Y/%m/%d\"\n";
        let table: toml::Table = toml.parse().unwrap();
        let fmt = table.get("default").and_then(|v| v.get("date_format")).and_then(|v| v.as_str());
        assert_eq!(fmt, Some("%Y/%m/%d"));
    }

    #[test]
    fn test_toml_output_format_tsv() {
        let toml = "[default]\noutput_format = \"tsv\"\n";
        let table: toml::Table = toml.parse().unwrap();
        let fmt = table.get("default").and_then(|v| v.get("output_format")).and_then(|v| v.as_str());
        assert_eq!(fmt, Some("tsv"));
    }

    #[test]
    fn test_toml_missing_key_returns_none() {
        // キーが存在しない場合は None
        let toml = "[default]\n";
        let table: toml::Table = toml.parse().unwrap();
        let enc = table.get("default").and_then(|v| v.get("encoding")).and_then(|v| v.as_str());
        assert_eq!(enc, None);
    }
}
