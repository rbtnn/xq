use calamine::{open_workbook_auto, Data, Reader};
use serde_json::{json, Map, Value};
use std::env;
use std::fs;
use std::path::Path;
use std::process::exit;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("usage: xq-rs <sheets|cells|names|vba|info> <file> [options]");
        exit(2);
    }
    let cmd = args[1].as_str();
    let file = match args.get(2) {
        Some(f) => f.clone(),
        None => {
            eprintln!("missing file");
            exit(2);
        }
    };

    let result = match cmd {
        "sheets" => cmd_sheets(&file, &args[3..]),
        "cells" => cmd_cells(&file, &args[3..]),
        "names" => cmd_names(&file, &args[3..]),
        "vba" => cmd_vba(&file, &args[3..]),
        "info" => cmd_info(&file, &args[3..]),
        _ => {
            eprintln!("unknown command: {}", cmd);
            exit(2);
        }
    };

    if let Err(e) = result {
        eprintln!("xq-rs: {}", e);
        exit(1);
    }
}

type BoxErr = Box<dyn std::error::Error>;

fn cmd_sheets(file: &str, _args: &[String]) -> Result<(), BoxErr> {
    let mut wb = open_workbook_auto(file)?;
    let mut out = Vec::new();
    for name in wb.sheet_names() {
        let dim = match wb.worksheet_range(&name) {
            Ok(range) => {
                if let (Some(s), Some(e)) = (range.start(), range.end()) {
                    Some(format!(
                        "{}:{}",
                        cell_ref(s.0, s.1),
                        cell_ref(e.0, e.1)
                    ))
                } else {
                    None
                }
            }
            Err(_) => None,
        };
        out.push(json!({ "name": name, "dimension": dim }));
    }
    println!("{}", serde_json::to_string(&out)?);
    Ok(())
}

fn cmd_cells(file: &str, args: &[String]) -> Result<(), BoxErr> {
    let mut wb = open_workbook_auto(file)?;

    let mut sheet: Option<String> = None;
    let mut range_arg: Option<(usize, usize, usize, usize)> = None;
    let mut as_json = false;
    let mut with_formulas = false;
    let mut all = false;
    let mut with_text = false;
    let mut limit: Option<usize> = None;

    let mut i = 0;
    while i < args.len() {
        match args[i].as_str() {
            "--sheet" => {
                i += 1;
                sheet = args.get(i).cloned();
            }
            "--range" => {
                i += 1;
                if let Some(s) = args.get(i) {
                    range_arg = parse_range_spec(&mut wb, s);
                }
            }
            "--json" => as_json = true,
            "--formulas" => with_formulas = true,
            "--all" => all = true,
            "--text" => with_text = true,
            "--limit" => {
                i += 1;
                match args.get(i).and_then(|s| s.parse::<usize>().ok()) {
                    Some(n) => limit = Some(n),
                    None => {
                        eprintln!("--limit needs a non-negative integer");
                        exit(2);
                    }
                }
            }
            _ => {}
        }
        i += 1;
    }

    let sheet_names: Vec<String> = match sheet {
        Some(s) => vec![s],
        None => wb.sheet_names(),
    };

    let mut rows: Vec<Value> = Vec::new();

    'outer: for sheet_name in sheet_names {
        if let Some(n) = limit {
            if n == 0 {
                break;
            }
        }

        let range = match wb.worksheet_range(&sheet_name) {
            Ok(r) => r,
            Err(_) => continue,
        };

        let formula_range = if with_formulas {
            wb.worksheet_formula(&sheet_name).ok()
        } else {
            None
        };

        let (r1, c1, r2, c2) = match range_arg {
            Some(r) => r,
            None => match (range.start(), range.end()) {
                (Some(s), Some(e)) => (s.0 as usize + 1, s.1 as usize + 1, e.0 as usize + 1, e.1 as usize + 1),
                _ => continue,
            },
        };

        let start = match range.start() {
            Some(s) => s,
            None => continue,
        };

        for (rel_row, rel_col, value) in range.cells() {
            if let Some(n) = limit {
                if rows.len() >= n {
                    break 'outer;
                }
            }

            let row = start.0 + rel_row as u32;
            let col = start.1 + rel_col as u32;
            let rr = row as usize + 1;
            let cc = col as usize + 1;
            if rr < r1 || rr > r2 || cc < c1 || cc > c2 {
                continue;
            }

            if matches!(value, Data::Empty) && !all {
                continue;
            }

            let (v, t) = value_json(value);
            let f = formula_range
                .as_ref()
                .and_then(|fr| fr.get_value((row, col)))
                .filter(|s| !s.is_empty())
                .cloned();

            let mut obj = Map::new();
            obj.insert("sheet".to_string(), json!(sheet_name));
            obj.insert("ref".to_string(), json!(cell_ref(row, col)));
            obj.insert("r".to_string(), json!(rr));
            obj.insert("c".to_string(), json!(cc));
            obj.insert("v".to_string(), v);
            obj.insert("t".to_string(), json!(t));
            if let Some(f) = f {
                obj.insert("f".to_string(), json!(f));
            }
            if with_text {
                obj.insert("text".to_string(), json!(value_text(value)));
            }

            rows.push(Value::Object(obj));

            if let Some(n) = limit {
                if rows.len() >= n {
                    break 'outer;
                }
            }
        }
    }

    if as_json {
        println!("{}", serde_json::to_string(&rows)?);
    } else {
        for row in rows {
            println!("{}", serde_json::to_string(&row)?);
        }
    }

    Ok(())
}

fn cmd_names(file: &str, args: &[String]) -> Result<(), BoxErr> {
    let wb = open_workbook_auto(file)?;
    let mut as_json = false;
    for a in args {
        if a == "--json" {
            as_json = true;
        }
    }

    let mut out = Vec::new();
    for (name, formula) in wb.defined_names() {
        let builtin = name.starts_with("_xlnm.");
        out.push(json!({
            "name": name,
            "scope": "workbook",
            "sheet": Value::Null,
            "sheetId": Value::Null,
            "hidden": false,
            "builtin": builtin,
            "ref": formula,
            "comment": Value::Null
        }));
    }

    if as_json {
        println!("{}", serde_json::to_string(&out)?);
    } else {
        for row in out {
            println!("{}", serde_json::to_string(&row)?);
        }
    }
    Ok(())
}

fn cmd_vba(file: &str, args: &[String]) -> Result<(), BoxErr> {
    let mut wb = open_workbook_auto(file)?;
    let mut as_json = false;
    for a in args {
        if a == "--json" {
            as_json = true;
        }
    }

    let mut modules = Vec::new();
    let code_page = 0;
    if let Some(Ok(vba)) = wb.vba_project() {
        for name in vba.get_module_names() {
            if let Ok(code) = vba.get_module(name) {
                modules.push(json!({ "name": name, "code": code }));
            }
        }
    }

    let out = json!({ "codePage": code_page, "modules": modules });
    if as_json {
        println!("{}", serde_json::to_string(&out)?);
    } else {
        println!("{}", serde_json::to_string_pretty(&out)?);
    }
    Ok(())
}

fn cmd_info(file: &str, _args: &[String]) -> Result<(), BoxErr> {
    let mut wb = open_workbook_auto(file)?;
    let mut sheets = Vec::new();
    for name in wb.sheet_names() {
        let dim = match wb.worksheet_range(&name) {
            Ok(range) => {
                if let (Some(s), Some(e)) = (range.start(), range.end()) {
                    Some(format!(
                        "{}:{}",
                        cell_ref(s.0, s.1),
                        cell_ref(e.0, e.1)
                    ))
                } else {
                    None
                }
            }
            Err(_) => None,
        };
        sheets.push(json!({ "name": name, "dimension": dim }));
    }

    let (has_vba, vba_modules) = match wb.vba_project() {
        Some(Ok(vba)) => (true, vba.get_module_names().len()),
        _ => (false, 0),
    };

    let size = fs::metadata(Path::new(file)).map(|m| m.len()).unwrap_or(0);
    let out = json!({
        "file": file,
        "size": size,
        "sheets": sheets,
        "hasVba": has_vba,
        "vbaModules": vba_modules,
        "vbaCodePage": 0,
        "definedNames": wb.defined_names().len()
    });
    println!("{}", serde_json::to_string(&out)?);
    Ok(())
}

fn value_json(v: &Data) -> (Value, &'static str) {
    match v {
        Data::Int(i) => (json!(i), "n"),
        Data::Float(f) => (json!(f), "n"),
        Data::String(s) => (json!(s), "s"),
        Data::Bool(b) => (json!(if *b { "true" } else { "false" }), "b"),
        Data::DateTime(dt) => (json!(dt.as_f64()), "n"),
        Data::DateTimeIso(s) => (json!(s), "s"),
        Data::DurationIso(s) => (json!(s), "s"),
        Data::Error(e) => (json!(format!("{:?}", e)), "e"),
        Data::Empty => (Value::Null, "z"),
    }
}

fn value_text(v: &Data) -> Option<String> {
    match v {
        Data::Int(i) => Some(i.to_string()),
        Data::Float(f) => Some(f.to_string()),
        Data::String(s) => Some(s.clone()),
        Data::Bool(b) => Some(if *b { "true" } else { "false" }.to_string()),
        Data::DateTime(dt) => Some(dt.as_f64().to_string()),
        Data::DateTimeIso(s) => Some(s.clone()),
        Data::DurationIso(s) => Some(s.clone()),
        Data::Error(e) => Some(format!("{:?}", e)),
        Data::Empty => None,
    }
}

fn cell_ref(row: u32, col: u32) -> String {
    format!("{}{}", col_name(col as usize + 1), row + 1)
}

fn col_name(mut n: usize) -> String {
    let mut s = String::new();
    while n > 0 {
        let r = (n - 1) % 26;
        s.insert(0, (b'A' + r as u8) as char);
        n = (n - 1) / 26;
    }
    s
}

fn parse_range_spec(
    wb: &mut calamine::Sheets<std::io::BufReader<std::fs::File>>,
    spec: &str,
) -> Option<(usize, usize, usize, usize)> {
    let mut text = spec.trim().to_string();

    if let Some((_, formula)) = wb
        .defined_names()
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(&text))
    {
        text = formula.clone();
    }

    if text.starts_with('=') {
        text = text[1..].to_string();
    }
    if let Some(pos) = text.rfind('!') {
        text = text[pos + 1..].to_string();
    }
    text = text.replace('$', "");

    parse_a1_range(&text)
}

fn parse_a1_range(s: &str) -> Option<(usize, usize, usize, usize)> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.is_empty() || parts.len() > 2 {
        return None;
    }
    let (r1, c1) = parse_a1(parts[0])?;
    let (r2, c2) = if parts.len() == 2 {
        parse_a1(parts[1])?
    } else {
        (r1, c1)
    };
    Some((r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
}

fn parse_a1(s: &str) -> Option<(usize, usize)> {
    let s = s.trim();
    let mut col = 0usize;
    let mut i = 0;
    while i < s.len() && s.as_bytes()[i].is_ascii_alphabetic() {
        let c = s.as_bytes()[i].to_ascii_uppercase();
        col = col * 26 + (c - b'A' + 1) as usize;
        i += 1;
    }
    if col == 0 || i == s.len() {
        return None;
    }
    let row: usize = s[i..].parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((row, col))
}
