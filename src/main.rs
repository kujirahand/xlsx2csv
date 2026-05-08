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
}

impl Config {
    // デフォルト値で初期化する
    fn default() -> Self {
        Config {
            encoding: "shift_jis".to_string(),
            date_format: "%Y-%m-%d".to_string(),
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

// メイン処理
fn main() {
    let args: Vec<String> = env::args().collect();

    // 設定をロードし、引数で上書きする
    let mut config = Config::default();
    config.load_from_file();
    let input_path = match config.apply_args(&args) {
        Some(p) => p,
        None => {
            eprintln!("Usage: {} [-e <encoding>] <excel_file>", args[0]);
            eprintln!("  -e, --encoding  出力エンコーディング (デフォルト: shift_jis)");
            eprintln!("  例: utf-8, shift_jis, euc-jp");
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

    convert_workbook_to_csv(&input_path, &stem, parent, encoding, &config.date_format);
}

// Excelファイルを開き、全シートをCSVへ書き出す
fn convert_workbook_to_csv(input_path: &str, stem: &str, parent: &Path, encoding: &'static Encoding, date_format: &str) {
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
        write_sheet_to_csv(&range, stem, sheet_name, parent, encoding, date_format);
    }
}

// 1シート分のセル範囲をCSVファイルに書き出す
fn write_sheet_to_csv(range: &Range<Data>, stem: &str, sheet_name: &str, parent: &Path, encoding: &'static Encoding, date_format: &str) {
    // 出力ファイルのパスを「入力ファイル名_シート名.csv」で組み立てる
    let output_filename = format!("{}_{}.csv", stem, sheet_name);
    let output_path = parent.join(&output_filename);
    // 出力ファイルを作成し、バッファ付きで書き込む
    let file = File::create(&output_path).expect("Cannot create output file");
    let mut writer = BufWriter::new(file);

    // 1行ずつセルを文字列化&エスケープし、カンマ区切りで書き出す
    for row in range.rows() {
        let fields: Vec<String> = row
            .iter()
            .map(|cell| escape_csv_field(&cell_to_string(cell, date_format)))
            .collect();
        // 指定エンコーディングに変換してバイト列として書き出す
        let line = fields.join(",") + "\n";
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
