use crate::App;
use crate::cell::CellValue;
use std::collections::HashMap;
use std::fs;
use std::io::{Read, Write};
use std::path::Path;
use serde::{Deserialize, Serialize};
use unicode_width::UnicodeWidthStr;

/// JSON file format for vicalc
#[derive(Serialize, Deserialize)]
struct VicalcFile {
    version: String,
    name: String,
    #[serde(default, skip_serializing_if = "HashMap::is_empty")]
    col_widths: HashMap<String, usize>,
    cells: HashMap<String, CellData>,
}

#[derive(Serialize, Deserialize)]
struct CellData {
    value: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    formula: Option<String>,
}

pub fn execute_command(app: &mut App, cmd: &str) {
    let cmd = cmd.trim();
    
    // Handle forward search :/pattern
    if cmd.starts_with('/') {
        let search_term = &cmd[1..];
        if !search_term.is_empty() {
            app.last_search = search_term.to_string();
            app.search_forward = true;
        }
        search_forward(app);
        return;
    }
    
    // Handle backward search :?pattern
    if cmd.starts_with('?') {
        let search_term = &cmd[1..];
        if !search_term.is_empty() {
            app.last_search = search_term.to_string();
            app.search_forward = false;
        }
        search_backward(app);
        return;
    }

    let parts: Vec<&str> = cmd.splitn(2, ' ').collect();
    let command = parts[0].to_lowercase();
    let args = if parts.len() > 1 { parts[1] } else { "" };

    match command.as_str() {
        "q" | "quit" => {
            app.running = false;
        }
        "q!" => {
            app.running = false;
        }
        "w" | "write" => {
            let filename = if args.is_empty() {
                app.current_file.clone().unwrap_or_else(|| "spreadsheet".to_string())
            } else {
                args.to_string()
            };
            match save_file(app, &filename) {
                Ok(actual_filename) => {
                    app.current_file = Some(actual_filename.clone());
                    app.status_message = format!("Saved to {}", actual_filename);
                }
                Err(e) => {
                    app.status_message = format!("Error saving: {}", e);
                }
            }
        }
        "wq" => {
            let filename = if args.is_empty() {
                app.current_file.clone().unwrap_or_else(|| "spreadsheet".to_string())
            } else {
                args.to_string()
            };
            match save_file(app, &filename) {
                Ok(actual_filename) => {
                    app.current_file = Some(actual_filename.clone());
                    app.status_message = format!("Saved to {}", actual_filename);
                    app.running = false;
                }
                Err(e) => {
                    app.status_message = format!("Error saving: {}", e);
                }
            }
        }
        "e" | "edit" | "open" => {
            if args.is_empty() {
                app.status_message = "Usage: :e <filename>".to_string();
            } else if let Err(e) = load_file(app, args) {
                app.status_message = format!("Error loading: {}", e);
            } else {
                app.current_file = Some(args.to_string());
                app.status_message = format!("Loaded {}", args);
            }
        }
        "export" => {
            if args.is_empty() {
                app.status_message = "Usage: :export <filename.csv>".to_string();
            } else if let Err(e) = export_csv(app, args) {
                app.status_message = format!("Error exporting: {}", e);
            } else {
                app.status_message = format!("Exported to {}", args);
            }
        }
        "import" => {
            if args.is_empty() {
                app.status_message = "Usage: :import <filename.csv>".to_string();
            } else if let Err(e) = import_csv(app, args) {
                app.status_message = format!("Error importing: {}", e);
            } else {
                app.status_message = format!("Imported {}", args);
            }
        }
        "goto" | "go" | "g" => {
            if let Some((col, row, _, _)) = crate::formula::parse_cell_ref(args) {
                app.cursor_col = col;
                app.cursor_row = row;
                app.adjust_view();
                app.status_message = format!("Moved to {}", crate::formula::cell_name(col, row));
            } else {
                app.status_message = "Invalid cell reference".to_string();
            }
        }
        "set" => {
            handle_set(app, args);
        }
        "delrow" | "dr" => {
            let row = if args.is_empty() {
                app.cursor_row
            } else {
                args.parse::<usize>().unwrap_or(app.cursor_row + 1).saturating_sub(1)
            };
            app.save_undo();
            app.sheet.adjust_formulas_for_row_delete(row);
            app.sheet.delete_row(row);
            app.status_message = format!("Deleted row {}", row + 1);
        }
        "delcol" | "dc" => {
            let col = if args.is_empty() {
                app.cursor_col
            } else {
                crate::formula::parse_cell_ref(&format!("{}1", args))
                    .map(|(c, _, _, _)| c)
                    .unwrap_or(app.cursor_col)
            };
            app.save_undo();
            app.sheet.adjust_formulas_for_col_delete(col);
            app.sheet.delete_col(col);
            app.status_message = format!("Deleted column {}", crate::formula::col_to_name(col));
        }
        "insrow" | "ir" => {
            let row = if args.is_empty() {
                app.cursor_row
            } else {
                args.parse::<usize>().unwrap_or(app.cursor_row + 1).saturating_sub(1)
            };
            app.save_undo();
            app.sheet.adjust_formulas_for_row_insert(row);
            app.sheet.insert_row(row);
            app.status_message = format!("Inserted row at {}", row + 1);
        }
        "inscol" | "ic" => {
            let col = if args.is_empty() {
                app.cursor_col
            } else {
                crate::formula::parse_cell_ref(&format!("{}1", args))
                    .map(|(c, _, _, _)| c)
                    .unwrap_or(app.cursor_col)
            };
            app.save_undo();
            app.sheet.adjust_formulas_for_col_insert(col);
            app.sheet.insert_col(col);
            app.status_message = format!("Inserted column at {}", crate::formula::col_to_name(col));
        }
        "clear" => {
            app.save_undo();
            app.sheet = crate::sheet::Sheet::new();
            app.cursor_col = 0;
            app.cursor_row = 0;
            app.view_col = 0;
            app.view_row = 0;
            app.current_file = None;
            app.status_message = "Sheet cleared".to_string();
        }
        "autowidth" | "aw" => {
            autowidth(app, args);
        }
        "help" | "h" => {
            app.status_message = "Commands: :w :q :wq :e :export :import :goto :set :autowidth :help".to_string();
        }
        "" => {}
        _ => {
            app.status_message = format!("Unknown command: {}", command);
        }
    }
}

fn handle_set(app: &mut App, args: &str) {
    let parts: Vec<&str> = args.splitn(2, '=').collect();
    if parts.len() != 2 {
        app.status_message = "Usage: :set option=value".to_string();
        return;
    }

    let option = parts[0].trim().to_lowercase();
    let _value = parts[1].trim();

    match option.as_str() {
        "name" | "sheet" => {
            app.sheet.name = _value.to_string();
            app.status_message = format!("Sheet name set to '{}'", _value);
        }
        _ => {
            app.status_message = format!("Unknown option: {}", option);
        }
    }
}

/// Auto-adjust column widths to fit content
fn autowidth(app: &mut App, args: &str) {
    const MIN_WIDTH: usize = 4;
    const MAX_WIDTH: usize = 50;
    
    let max_row = app.sheet.max_row().unwrap_or(0);
    
    if args.is_empty() {
        // Adjust all columns with data
        let max_col = app.sheet.max_col().unwrap_or(0);
        let mut adjusted = 0;
        
        for col in 0..=max_col {
            let width = calc_column_width(app, col, max_row, MIN_WIDTH, MAX_WIDTH);
            app.sheet.set_col_width(col, width);
            adjusted += 1;
        }
        
        app.status_message = format!("Auto-adjusted {} columns", adjusted);
    } else {
        // Parse column range (e.g., "A", "A:C", "B:D")
        let args_upper = args.to_uppercase();
        
        if let Some(colon_pos) = args_upper.find(':') {
            // Range: A:C
            let start_str = &args_upper[..colon_pos];
            let end_str = &args_upper[colon_pos + 1..];
            
            if let (Some(start_col), Some(end_col)) = (parse_col_name(start_str), parse_col_name(end_str)) {
                for col in start_col..=end_col {
                    let width = calc_column_width(app, col, max_row, MIN_WIDTH, MAX_WIDTH);
                    app.sheet.set_col_width(col, width);
                }
                app.status_message = format!("Auto-adjusted columns {}:{}", start_str, end_str);
            } else {
                app.status_message = "Invalid column range".to_string();
            }
        } else {
            // Single column: A
            if let Some(col) = parse_col_name(&args_upper) {
                let width = calc_column_width(app, col, max_row, MIN_WIDTH, MAX_WIDTH);
                app.sheet.set_col_width(col, width);
                app.status_message = format!("Column {} width set to {}", args_upper, width);
            } else {
                app.status_message = "Invalid column name".to_string();
            }
        }
    }
}

/// Calculate optimal width for a column
fn calc_column_width(app: &App, col: usize, max_row: usize, min_width: usize, max_width: usize) -> usize {
    let mut width = min_width;
    
    // Check header (column name)
    let col_name = crate::formula::col_to_name(col);
    width = width.max(UnicodeWidthStr::width(col_name.as_str()) + 2);
    
    // Check all cells in the column
    for row in 0..=max_row {
        let value = app.sheet.evaluate(col, row);
        let cell_width = UnicodeWidthStr::width(value.as_str()) + 2;  // +2 for padding
        width = width.max(cell_width);
    }
    
    width.min(max_width)
}

/// Parse column name (A, B, AA, etc.) to index
fn parse_col_name(name: &str) -> Option<usize> {
    let name = name.trim();
    if name.is_empty() {
        return None;
    }
    
    let mut col = 0usize;
    for c in name.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    
    Some(col.saturating_sub(1))
}

/// Search forward from current position
pub fn search_forward(app: &mut App) {
    if app.last_search.is_empty() {
        app.status_message = "No search pattern".to_string();
        return;
    }

    let term = app.last_search.clone();
    let term_upper = term.to_uppercase();
    let start_col = app.cursor_col;
    let start_row = app.cursor_row;

    // Search from current position forward
    for row in start_row..10000 {
        let col_start = if row == start_row { start_col + 1 } else { 0 };
        for col in col_start..256 {
            let value = app.sheet.evaluate(col, row);
            if value.to_uppercase().contains(&term_upper) {
                app.cursor_col = col;
                app.cursor_row = row;
                app.adjust_view();
                app.status_message = format!("/{} -> {}", term, crate::formula::cell_name(col, row));
                return;
            }
        }
    }

    // Wrap around
    for row in 0..=start_row {
        let col_end = if row == start_row { start_col } else { 256 };
        for col in 0..col_end {
            let value = app.sheet.evaluate(col, row);
            if value.to_uppercase().contains(&term_upper) {
                app.cursor_col = col;
                app.cursor_row = row;
                app.adjust_view();
                app.status_message = format!("/{} -> {} (wrapped)", term, crate::formula::cell_name(col, row));
                return;
            }
        }
    }

    app.status_message = format!("Pattern not found: {}", term);
}

/// Search backward from current position
pub fn search_backward(app: &mut App) {
    if app.last_search.is_empty() {
        app.status_message = "No search pattern".to_string();
        return;
    }

    let term = app.last_search.clone();
    let term_upper = term.to_uppercase();
    let start_col = app.cursor_col;
    let start_row = app.cursor_row;

    // Search backward from current position
    for row in (0..=start_row).rev() {
        let col_end = if row == start_row { start_col } else { 256 };
        for col in (0..col_end).rev() {
            let value = app.sheet.evaluate(col, row);
            if value.to_uppercase().contains(&term_upper) {
                app.cursor_col = col;
                app.cursor_row = row;
                app.adjust_view();
                app.status_message = format!("?{} -> {}", term, crate::formula::cell_name(col, row));
                return;
            }
        }
    }

    // Wrap around from end
    for row in (start_row..10000).rev() {
        let col_start = if row == start_row { start_col + 1 } else { 0 };
        for col in (col_start..256).rev() {
            let value = app.sheet.evaluate(col, row);
            if value.to_uppercase().contains(&term_upper) {
                app.cursor_col = col;
                app.cursor_row = row;
                app.adjust_view();
                app.status_message = format!("?{} -> {} (wrapped)", term, crate::formula::cell_name(col, row));
                return;
            }
        }
    }

    app.status_message = format!("Pattern not found: {}", term);
}

/// Search next (n key) - same direction as last search
pub fn search_next(app: &mut App) {
    if app.search_forward {
        search_forward(app);
    } else {
        search_backward(app);
    }
}

/// Search previous (N key) - opposite direction
pub fn search_prev(app: &mut App) {
    if app.search_forward {
        search_backward(app);
    } else {
        search_forward(app);
    }
}

/// Save file. Returns the actual filename used.
fn save_file(app: &App, filename: &str) -> Result<String, String> {
    let path = Path::new(filename);
    let ext = path.extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    // Auto-add .json extension if no extension provided
    let filename = if ext.is_empty() {
        format!("{}.json", filename)
    } else {
        filename.to_string()
    };

    let ext = Path::new(&filename).extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    match ext.as_str() {
        "csv" => {
            export_csv(app, &filename).map_err(|e| e.to_string())?;
            Ok(filename)
        }
        _ => {
            // Default to JSON
            save_json(app, &filename).map_err(|e| e.to_string())?;
            Ok(filename)
        }
    }
}

/// Load file
fn load_file(app: &mut App, filename: &str) -> Result<(), String> {
    let path = Path::new(filename);
    let ext = path.extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    match ext.as_str() {
        "csv" => {
            import_csv(app, filename).map_err(|e| e.to_string())
        }
        _ => {
            // Default to JSON
            load_json(app, filename).map_err(|e| e.to_string())
        }
    }
}

fn save_json(app: &App, filename: &str) -> std::io::Result<()> {
    use crate::sheet::DEFAULT_COL_WIDTH;
    
    // Build col_widths map (only non-default widths)
    let mut col_widths = HashMap::new();
    for col in 0..=255 {
        let width = app.sheet.get_col_width(col);
        if width != DEFAULT_COL_WIDTH {
            let col_name = crate::formula::col_to_name(col);
            col_widths.insert(col_name, width);
        }
    }
    
    // Build cells map
    let mut cells = HashMap::new();
    for ((col, row), cell) in app.sheet.cells().iter() {
        let cell_name = crate::formula::cell_name(*col, *row);
        
        // Get the evaluated value for storage
        let evaluated = app.sheet.evaluate(*col, *row);
        
        let cell_data = match &cell.value {
            CellValue::Formula(_) => {
                CellData {
                    value: evaluated,
                    formula: Some(cell.raw_input.clone()),
                }
            }
            _ => {
                CellData {
                    value: cell.raw_input.clone(),
                    formula: None,
                }
            }
        };
        
        cells.insert(cell_name, cell_data);
    }
    
    let file_data = VicalcFile {
        version: "1.0".to_string(),
        name: app.sheet.name.clone(),
        col_widths,
        cells,
    };
    
    let json = serde_json::to_string_pretty(&file_data)
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::Other, e))?;
    let mut file = fs::File::create(filename)?;
    file.write_all(json.as_bytes())?;
    Ok(())
}

fn load_json(app: &mut App, filename: &str) -> std::io::Result<()> {
    let mut file = fs::File::open(filename)?;
    let mut contents = String::new();
    file.read_to_string(&mut contents)?;
    
    let file_data: VicalcFile = serde_json::from_str(&contents)
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::Other, e))?;
    
    app.save_undo();
    
    // Create new sheet
    let mut sheet = crate::sheet::Sheet::new();
    sheet.name = file_data.name;
    
    // Restore col_widths
    for (col_name, width) in file_data.col_widths {
        if let Some((col, _, _, _)) = crate::formula::parse_cell_ref(&format!("{}1", col_name)) {
            sheet.set_col_width(col, width);
        }
    }
    
    // Restore cells
    for (cell_name, cell_data) in file_data.cells {
        if let Some((col, row, _, _)) = crate::formula::parse_cell_ref(&cell_name) {
            // If formula exists, use formula; otherwise use value
            let input = cell_data.formula.unwrap_or(cell_data.value);
            sheet.set_cell(col, row, input);
        }
    }
    
    app.sheet = sheet;
    app.cursor_col = 0;
    app.cursor_row = 0;
    app.view_col = 0;
    app.view_row = 0;
    Ok(())
}

fn export_csv(app: &App, filename: &str) -> std::io::Result<()> {
    let max_col = app.sheet.max_col().unwrap_or(0);
    let max_row = app.sheet.max_row().unwrap_or(0);

    let mut csv = String::new();
    for row in 0..=max_row {
        let mut row_values = Vec::new();
        for col in 0..=max_col {
            let value = app.sheet.evaluate(col, row);
            // Escape quotes and wrap in quotes if needed
            if value.contains(',') || value.contains('"') || value.contains('\n') {
                row_values.push(format!("\"{}\"", value.replace('"', "\"\"")));
            } else {
                row_values.push(value);
            }
        }
        csv.push_str(&row_values.join(","));
        csv.push('\n');
    }

    let mut file = fs::File::create(filename)?;
    file.write_all(csv.as_bytes())?;
    Ok(())
}

fn import_csv(app: &mut App, filename: &str) -> std::io::Result<()> {
    let mut file = fs::File::open(filename)?;
    let mut contents = String::new();
    file.read_to_string(&mut contents)?;

    app.save_undo();
    app.sheet = crate::sheet::Sheet::new();

    for (row, line) in contents.lines().enumerate() {
        let mut col = 0;
        let mut current = String::new();
        let mut in_quotes = false;
        let mut chars = line.chars().peekable();

        while let Some(c) = chars.next() {
            if c == '"' {
                if in_quotes && chars.peek() == Some(&'"') {
                    // Escaped quote
                    current.push('"');
                    chars.next();
                } else {
                    in_quotes = !in_quotes;
                }
            } else if c == ',' && !in_quotes {
                if !current.is_empty() {
                    app.sheet.set_cell(col, row, current.clone());
                }
                current.clear();
                col += 1;
            } else {
                current.push(c);
            }
        }
        
        if !current.is_empty() {
            app.sheet.set_cell(col, row, current);
        }
    }

    app.cursor_col = 0;
    app.cursor_row = 0;
    app.view_col = 0;
    app.view_row = 0;
    Ok(())
}
