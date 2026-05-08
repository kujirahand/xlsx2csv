// Excelファイル(xlsx/xls/ods等)をシートごとにCSVへ変換するCLIツール
use calamine::{open_workbook_auto, Data, Range, Reader};
use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::Path;

// メイン処理 --- (*1)
fn main() {
    // コマンドライン引数を取得する --- (*2)
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <excel_file>", args[0]);
        std::process::exit(1);
    }
    // 入力ファイルのパスを解析 --- (*3)
    // 入力ファイルのパスから「拡張子を除いたファイル名」と「親ディレクトリ」を取り出す
    // 出力ファイル名 "入力ファイル名_シート名.csv" を組み立てるために使う
    let input_path = &args[1];
    let path = Path::new(input_path);
    let stem = path
        .file_stem()
        .expect("Invalid input file name")
        .to_string_lossy()
        .to_string();
    let parent = path.parent().unwrap_or_else(|| Path::new(""));

    // 入力Excelをシートごとに分割してCSVへ出力する --- (*4)
    convert_workbook_to_csv(input_path, &stem, parent);
}

// Excelファイルを開き、全シートをCSVへ書き出す --- (*5)
fn convert_workbook_to_csv(input_path: &str, stem: &str, parent: &Path) {
    // Excelファイルを開く (拡張子から自動でフォーマットを判定) --- (*6)
    let mut workbook = open_workbook_auto(input_path).expect("Cannot open Excel file");
    // 全シート名を取得する --- (*7)
    let sheet_names = workbook.sheet_names().to_owned();

    // シートごとに CSV ファイルを生成する --- (*8)
    for sheet_name in &sheet_names {
        // シートの内容を取得する --- (*9)
        let range = match workbook.worksheet_range(sheet_name) {
            Ok(r) => r,
            Err(e) => {
                // 取得失敗したシートはスキップ
                eprintln!("Skip sheet '{}': {}", sheet_name, e);
                continue;
            }
        };
        // 1シート分をCSVに書き出す --- (*10)
        write_sheet_to_csv(&range, stem, sheet_name, parent);
    }
}

// 1シート分のセル範囲をCSVファイルに書き出す --- (*11)
fn write_sheet_to_csv(range: &Range<Data>, stem: &str, sheet_name: &str, parent: &Path) {
    // 出力ファイルのパスを「入力ファイル名_シート名.csv」で組み立てる --- (*12)
    let output_filename = format!("{}_{}.csv", stem, sheet_name);
    let output_path = parent.join(&output_filename);
    // 出力ファイルを作成し、バッファ付きで書き込む --- (*13)
    let file = File::create(&output_path).expect("Cannot create output file");
    let mut writer = BufWriter::new(file);

    // 1行ずつセルを文字列化&エスケープし、カンマ区切りで書き出す --- (*14)
    for row in range.rows() {
        let fields: Vec<String> = row
            .iter()
            .map(|cell| escape_csv_field(&cell_to_string(cell)))
            .collect();
        writeln!(writer, "{}", fields.join(",")).expect("Write failed");
    }

    // 出力したファイルパスを表示する
    println!("出力: {}", output_path.display());
}

// CSVフィールドの値をエスケープする --- (*15)
// カンマ・ダブルクォート・改行を含む場合は "..." で囲み、内部の " は "" に置換する
fn escape_csv_field(s: &str) -> String {
    if s.contains(',') || s.contains('"') || s.contains('\n') || s.contains('\r') {
        let escaped = s.replace('"', "\"\"");
        format!("\"{}\"", escaped)
    } else {
        s.to_string()
    }
}

// Excelの1セル(Data型)を文字列に変換する --- (*16)
fn cell_to_string(cell: &Data) -> String {
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
        // 日時セル: Excel内部のシリアル値を NaiveDateTime に変換し YYYY-MM-DD で出力
        Data::DateTime(dt) => match dt.as_datetime() {
            Some(ndt) => ndt.format("%Y-%m-%d").to_string(),
            // 変換失敗時はシリアル値(数値)をそのまま出力する
            None => dt.as_f64().to_string(),
        },
        // ISO形式の日時文字列セル: パースして YYYY-MM-DD に整形する
        Data::DateTimeIso(s) => {
            if let Ok(ndt) = chrono::NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S") {
                ndt.format("%Y-%m-%d").to_string()
            } else if let Ok(nd) = chrono::NaiveDate::parse_from_str(s, "%Y-%m-%d") {
                nd.format("%Y-%m-%d").to_string()
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
