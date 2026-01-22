mod cell;
mod engine;
mod formula;
mod sheet;
mod ui;
mod commands;

use crossterm::{
    cursor::{Hide, Show},
    event::{self, Event, KeyCode, KeyEvent, KeyModifiers, MouseEvent, MouseEventKind, MouseButton, EnableMouseCapture, DisableMouseCapture},
    execute,
    terminal::{self, EnterAlternateScreen, LeaveAlternateScreen},
};
use std::io::{stdout, Result};

use sheet::Sheet;
use ui::UI;

/// Operation modes
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Mode {
    Normal,         // Navigation and commands
    EditSingle,     // r - single cell edit, return to Normal after
    EditContinuous, // R - continuous edit, stay in edit mode
    EditPreserve,   // F2 - edit preserving content
    Command,        // : commands
    Visual,         // Range selection
}

/// Edit axis (row-oriented or column-oriented)
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum EditAxis {
    Row,    // Row mode (horizontal editing) - default
    Column, // Column mode (vertical editing)
}

pub struct App {
    pub sheet: Sheet,
    pub mode: Mode,
    pub axis: EditAxis,
    pub input_buffer: String,
    pub command_buffer: String,
    pub status_message: String,
    pub cursor_col: usize,
    pub cursor_row: usize,
    pub view_col: usize,
    pub view_row: usize,
    pub clipboard: Option<ClipboardContent>,
    pub undo_stack: Vec<Sheet>,
    pub redo_stack: Vec<Sheet>,
    pub running: bool,
    pub pending_operator: Option<char>,
    pub count_buffer: String,
    pub slash_pending: bool,
    pub current_file: Option<String>,
    // Visual mode selection
    pub visual_start_col: usize,
    pub visual_start_row: usize,
    // Original cell content before editing (for cancel)
    pub edit_original: String,
    // Search
    pub last_search: String,
    pub search_forward: bool,
    // Register pending ("* for system clipboard)
    pub register_pending: bool,
    // Last paste info for repeat paste (pp)
    pub last_paste_cols: usize,
    pub last_paste_rows: usize,
}

#[derive(Clone)]
pub struct ClipboardContent {
    pub cells: Vec<Vec<(String, crate::cell::CellValue)>>,  // [row][col] = (raw_input, value)
    pub start_col: usize,
    pub start_row: usize,
    pub width: usize,
    pub height: usize,
}

impl App {
    pub fn new() -> Self {
        let mut app = App {
            sheet: Sheet::new(),
            mode: Mode::Normal,
            axis: EditAxis::Row,
            input_buffer: String::new(),
            command_buffer: String::new(),
            status_message: String::new(),
            cursor_col: 0,
            cursor_row: 0,
            view_col: 0,
            view_row: 0,
            clipboard: None,
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            running: true,
            pending_operator: None,
            count_buffer: String::new(),
            slash_pending: false,
            current_file: None,
            visual_start_col: 0,
            visual_start_row: 0,
            edit_original: String::new(),
            last_search: String::new(),
            search_forward: true,
            register_pending: false,
            last_paste_cols: 0,
            last_paste_rows: 0,
        };
        app.update_status();
        app
    }

    fn update_status(&mut self) {
        let mode_str = match self.mode {
            Mode::Normal => "NORMAL",
            Mode::EditSingle => "EDIT",
            Mode::EditContinuous => "EDIT+",
            Mode::EditPreserve => "EDIT",
            Mode::Command => "COMMAND",
            Mode::Visual => "VISUAL",
        };
        let axis_str = match self.axis {
            EditAxis::Row => "Row",
            EditAxis::Column => "Col",
        };
        let cell_name = crate::formula::cell_name(self.cursor_col, self.cursor_row);
        let file_str = self.current_file.as_deref().unwrap_or("[New]");
        self.status_message = format!("{} | {} | {} | {}", mode_str, cell_name, axis_str, file_str);
    }

    pub fn save_undo(&mut self) {
        self.undo_stack.push(self.sheet.clone());
        self.redo_stack.clear();
        if self.undo_stack.len() > 100 {
            self.undo_stack.remove(0);
        }
    }

    pub fn undo(&mut self) {
        if let Some(prev) = self.undo_stack.pop() {
            self.redo_stack.push(self.sheet.clone());
            self.sheet = prev;
            self.status_message = "Undo".to_string();
        } else {
            self.status_message = "Nothing to undo".to_string();
        }
    }

    pub fn redo(&mut self) {
        if let Some(next) = self.redo_stack.pop() {
            self.undo_stack.push(self.sheet.clone());
            self.sheet = next;
            self.status_message = "Redo".to_string();
        } else {
            self.status_message = "Nothing to redo".to_string();
        }
    }

    pub fn get_count(&mut self) -> usize {
        let count = self.count_buffer.parse::<usize>().unwrap_or(1);
        self.count_buffer.clear();
        count
    }

    pub fn move_cursor(&mut self, dx: isize, dy: isize) {
        let count = self.get_count() as isize;
        let new_col = (self.cursor_col as isize + dx * count).max(0).min(255) as usize;
        let new_row = (self.cursor_row as isize + dy * count).max(0).min(9999) as usize;
        self.cursor_col = new_col;
        self.cursor_row = new_row;
        self.adjust_view();
    }

    pub fn move_cursor_to(&mut self, col: usize, row: usize) {
        self.cursor_col = col.min(255);
        self.cursor_row = row.min(9999);
        self.adjust_view();
    }

    pub fn adjust_view(&mut self) {
        const ROW_LABEL_WIDTH: usize = 5;
        
        let (term_width, term_height) = terminal::size().unwrap_or((80, 24));
        let available_width = (term_width as usize).saturating_sub(ROW_LABEL_WIDTH);
        let visible_rows = (term_height as usize).saturating_sub(5);

        // Adjust view_col to ensure cursor is visible
        if self.cursor_col < self.view_col {
            self.view_col = self.cursor_col;
        } else {
            // Check if cursor column is visible
            let mut x = 0;
            let mut col = self.view_col;
            let mut cursor_visible = false;
            
            while x < available_width && col <= 255 {
                let col_width = self.sheet.get_col_width(col);
                if col == self.cursor_col {
                    if x + col_width <= available_width {
                        cursor_visible = true;
                    }
                    break;
                }
                x += col_width;
                col += 1;
            }
            
            if !cursor_visible {
                // Scroll right to show cursor
                self.view_col = self.cursor_col;
            }
        }

        // Adjust view_row
        if self.cursor_row < self.view_row {
            self.view_row = self.cursor_row;
        } else if self.cursor_row >= self.view_row + visible_rows {
            self.view_row = self.cursor_row.saturating_sub(visible_rows - 1);
        }
    }

    pub fn screen_to_cell(&self, screen_col: u16, screen_row: u16) -> Option<(usize, usize)> {
        const ROW_LABEL_WIDTH: usize = 5;
        const HEADER_ROWS: usize = 2;  // status bar + column headers

        let screen_col = screen_col as usize;
        let screen_row = screen_row as usize;

        // Get terminal size for bounds checking
        let (term_width, term_height) = terminal::size().unwrap_or((80, 24));
        let grid_height = (term_height as usize).saturating_sub(4);  // subtract header (2) + footer (2)

        // Check if click is in the grid area
        if screen_col < ROW_LABEL_WIDTH || screen_row < HEADER_ROWS {
            return None;
        }

        // Check if click is below grid or in footer area
        if screen_row >= HEADER_ROWS + grid_height {
            return None;
        }

        // Calculate which column was clicked based on variable widths
        let mut x = ROW_LABEL_WIDTH;
        let mut col = self.view_col;
        while x < term_width as usize && col <= 255 {
            let col_width = self.sheet.get_col_width(col);
            if screen_col < x + col_width {
                // Click is in this column
                let row = self.view_row + (screen_row - HEADER_ROWS);
                return Some((col, row));
            }
            x += col_width;
            col += 1;
        }

        None
    }

    // Axis-dependent movement
    pub fn goto_axis_start(&mut self) {
        match self.axis {
            EditAxis::Row => self.cursor_col = 0,
            EditAxis::Column => self.cursor_row = 0,
        }
        self.adjust_view();
    }

    pub fn goto_first_non_empty(&mut self) {
        match self.axis {
            EditAxis::Row => {
                if let Some(col) = self.sheet.first_non_empty_col_in_row(self.cursor_row) {
                    self.cursor_col = col;
                }
            }
            EditAxis::Column => {
                if let Some(row) = self.sheet.first_non_empty_row_in_col(self.cursor_col) {
                    self.cursor_row = row;
                }
            }
        }
        self.adjust_view();
    }

    pub fn goto_axis_end(&mut self) {
        match self.axis {
            EditAxis::Row => {
                self.cursor_col = self.sheet.max_col_in_row(self.cursor_row).unwrap_or(0);
            }
            EditAxis::Column => {
                self.cursor_row = self.sheet.max_row_in_col(self.cursor_col).unwrap_or(0);
            }
        }
        self.adjust_view();
    }

    // Structure operations
    pub fn insert_at_cursor(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                self.sheet.shift_cells_right(self.cursor_col, self.cursor_row);
            }
            EditAxis::Column => {
                self.sheet.shift_cells_down(self.cursor_col, self.cursor_row);
            }
        }
        self.mode = Mode::EditSingle;
        self.input_buffer.clear();
        self.edit_original.clear();
        self.update_status();
    }

    pub fn insert_at_start(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                self.cursor_col = 0;
                self.sheet.shift_cells_right(0, self.cursor_row);
            }
            EditAxis::Column => {
                self.cursor_row = 0;
                self.sheet.shift_cells_down(self.cursor_col, 0);
            }
        }
        self.adjust_view();
        self.mode = Mode::EditSingle;
        self.input_buffer.clear();
        self.edit_original.clear();
        self.update_status();
    }

    pub fn append_after_cursor(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                self.cursor_col += 1;
                self.sheet.shift_cells_right(self.cursor_col, self.cursor_row);
            }
            EditAxis::Column => {
                self.cursor_row += 1;
                self.sheet.shift_cells_down(self.cursor_col, self.cursor_row);
            }
        }
        self.adjust_view();
        self.mode = Mode::EditSingle;
        self.input_buffer.clear();
        self.edit_original.clear();
        self.update_status();
    }

    pub fn goto_axis_end_next(&mut self) {
        match self.axis {
            EditAxis::Row => {
                let end = self.sheet.max_col_in_row(self.cursor_row).unwrap_or(0);
                self.cursor_col = (end + 1).min(255);
            }
            EditAxis::Column => {
                let end = self.sheet.max_row_in_col(self.cursor_col).unwrap_or(0);
                self.cursor_row = (end + 1).min(9999);
            }
        }
        self.adjust_view();
    }

    pub fn delete_structure(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                self.sheet.adjust_formulas_for_row_delete(self.cursor_row);
                self.sheet.delete_row(self.cursor_row);
                self.status_message = "Row deleted".to_string();
            }
            EditAxis::Column => {
                self.sheet.adjust_formulas_for_col_delete(self.cursor_col);
                self.sheet.delete_col(self.cursor_col);
                self.status_message = "Column deleted".to_string();
            }
        }
    }

    pub fn insert_structure_after(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                self.sheet.adjust_formulas_for_row_insert(self.cursor_row + 1);
                self.sheet.insert_row(self.cursor_row + 1);
                self.cursor_row += 1;
                self.status_message = "Row inserted below".to_string();
            }
            EditAxis::Column => {
                self.sheet.adjust_formulas_for_col_insert(self.cursor_col + 1);
                self.sheet.insert_col(self.cursor_col + 1);
                self.cursor_col += 1;
                self.status_message = "Column inserted right".to_string();
            }
        }
        self.adjust_view();
    }

    pub fn insert_structure_before(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                self.sheet.adjust_formulas_for_row_insert(self.cursor_row);
                self.sheet.insert_row(self.cursor_row);
                self.status_message = "Row inserted above".to_string();
            }
            EditAxis::Column => {
                self.sheet.adjust_formulas_for_col_insert(self.cursor_col);
                self.sheet.insert_col(self.cursor_col);
                self.status_message = "Column inserted left".to_string();
            }
        }
        self.adjust_view();
    }

    pub fn commit_input_and_move(&mut self) {
        if !self.input_buffer.is_empty() {
            self.save_undo();
            self.sheet.set_cell(self.cursor_col, self.cursor_row, self.input_buffer.clone());
        }
        self.input_buffer.clear();

        match self.axis {
            EditAxis::Row => {
                self.cursor_col = (self.cursor_col + 1).min(255);
            }
            EditAxis::Column => {
                self.cursor_row = (self.cursor_row + 1).min(9999);
            }
        }
        self.adjust_view();
    }

    // Content clear operations (clear cell contents, not structure)
    
    /// Clear current cell (x command)
    pub fn clear_current_cell(&mut self) {
        self.save_undo();
        self.sheet.clear_cell(self.cursor_col, self.cursor_row);
        self.status_message = "Cell cleared".to_string();
    }

    /// Clear cells from current to end of axis (d$ command)
    pub fn clear_to_axis_end(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                // Clear from current column to end of row
                let max_col = self.sheet.max_col_in_row(self.cursor_row).unwrap_or(self.cursor_col);
                let mut count = 0;
                for col in self.cursor_col..=max_col {
                    self.sheet.clear_cell(col, self.cursor_row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
            EditAxis::Column => {
                // Clear from current row to end of column
                let max_row = self.sheet.max_row_in_col(self.cursor_col).unwrap_or(self.cursor_row);
                let mut count = 0;
                for row in self.cursor_row..=max_row {
                    self.sheet.clear_cell(self.cursor_col, row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
        }
    }

    /// Clear cells from axis start to current (d0 command)
    pub fn clear_from_axis_start(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                // Clear from column A to current column
                let mut count = 0;
                for col in 0..=self.cursor_col {
                    self.sheet.clear_cell(col, self.cursor_row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
            EditAxis::Column => {
                // Clear from row 1 to current row
                let mut count = 0;
                for row in 0..=self.cursor_row {
                    self.sheet.clear_cell(self.cursor_col, row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
        }
    }

    /// Clear cells from first non-empty to current (d^ command)
    pub fn clear_from_first_non_empty(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                let start_col = self.sheet.first_non_empty_col_in_row(self.cursor_row).unwrap_or(0);
                let mut count = 0;
                for col in start_col..=self.cursor_col {
                    self.sheet.clear_cell(col, self.cursor_row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
            EditAxis::Column => {
                let start_row = self.sheet.first_non_empty_row_in_col(self.cursor_col).unwrap_or(0);
                let mut count = 0;
                for row in start_row..=self.cursor_row {
                    self.sheet.clear_cell(self.cursor_col, row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
        }
    }

    /// Clear cells from current to sheet end along axis (dG command)
    pub fn clear_to_sheet_end(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                let max_col = self.sheet.max_col().unwrap_or(self.cursor_col);
                let mut count = 0;
                for col in self.cursor_col..=max_col {
                    self.sheet.clear_cell(col, self.cursor_row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
            EditAxis::Column => {
                let max_row = self.sheet.max_row().unwrap_or(self.cursor_row);
                let mut count = 0;
                for row in self.cursor_row..=max_row {
                    self.sheet.clear_cell(self.cursor_col, row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
        }
    }

    /// Clear cells from sheet start to current along axis (dgg command)
    pub fn clear_from_sheet_start(&mut self) {
        self.save_undo();
        match self.axis {
            EditAxis::Row => {
                let mut count = 0;
                for col in 0..=self.cursor_col {
                    self.sheet.clear_cell(col, self.cursor_row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
            EditAxis::Column => {
                let mut count = 0;
                for row in 0..=self.cursor_row {
                    self.sheet.clear_cell(self.cursor_col, row);
                    count += 1;
                }
                self.status_message = format!("{} cells cleared", count);
            }
        }
    }

    /// Get selection bounds (min_col, min_row, max_col, max_row)
    pub fn get_selection_bounds(&self) -> (usize, usize, usize, usize) {
        let min_col = self.visual_start_col.min(self.cursor_col);
        let max_col = self.visual_start_col.max(self.cursor_col);
        let min_row = self.visual_start_row.min(self.cursor_row);
        let max_row = self.visual_start_row.max(self.cursor_row);
        (min_col, min_row, max_col, max_row)
    }

    /// Clear selected range (Visual mode)
    pub fn clear_selection(&mut self) {
        let (min_col, min_row, max_col, max_row) = self.get_selection_bounds();
        self.save_undo();
        let mut count = 0;
        for col in min_col..=max_col {
            for row in min_row..=max_row {
                self.sheet.clear_cell(col, row);
                count += 1;
            }
        }
        self.status_message = format!("{} cells cleared", count);
        self.mode = Mode::Normal;
    }

    /// Copy current cell or selection to internal clipboard
    pub fn yank(&mut self) {
        let (min_col, min_row, max_col, max_row) = if self.mode == Mode::Visual {
            self.get_selection_bounds()
        } else {
            (self.cursor_col, self.cursor_row, self.cursor_col, self.cursor_row)
        };

        let width = max_col - min_col + 1;
        let height = max_row - min_row + 1;

        let mut cells = Vec::new();
        for row in min_row..=max_row {
            let mut row_data = Vec::new();
            for col in min_col..=max_col {
                let cell = self.sheet.get_cell(col, row);
                row_data.push((cell.raw_input.clone(), cell.value.clone()));
            }
            cells.push(row_data);
        }

        self.clipboard = Some(ClipboardContent {
            cells,
            start_col: min_col,
            start_row: min_row,
            width,
            height,
        });

        // Move cursor to selection start (top-left)
        self.cursor_col = min_col;
        self.cursor_row = min_row;
        self.adjust_view();

        self.status_message = format!("Copied {}x{} cells", width, height);
        self.mode = Mode::Normal;
    }

    /// Copy to system clipboard as TSV
    pub fn yank_to_system(&mut self) {
        let (min_col, min_row, max_col, max_row) = if self.mode == Mode::Visual {
            self.get_selection_bounds()
        } else {
            (self.cursor_col, self.cursor_row, self.cursor_col, self.cursor_row)
        };

        let width = max_col - min_col + 1;
        let height = max_row - min_row + 1;

        let mut tsv = String::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                if col > min_col {
                    tsv.push('\t');
                }
                let value = self.sheet.evaluate(col, row);
                tsv.push_str(&value);
            }
            tsv.push('\n');
        }

        // Copy to system clipboard
        if let Ok(mut clipboard) = arboard::Clipboard::new() {
            if clipboard.set_text(&tsv).is_ok() {
                self.status_message = format!("Copied {}x{} cells to clipboard", width, height);
            } else {
                self.status_message = "Failed to copy to clipboard".to_string();
            }
        } else {
            self.status_message = "Clipboard not available".to_string();
        }

        // Move cursor to selection start (top-left)
        self.cursor_col = min_col;
        self.cursor_row = min_row;
        self.adjust_view();

        self.mode = Mode::Normal;
    }

    /// Paste from internal clipboard
    pub fn paste(&mut self, count: usize) {
        if self.clipboard.is_none() {
            self.status_message = "Nothing to paste".to_string();
            return;
        }

        let clip = self.clipboard.clone().unwrap();
        self.save_undo();

        let mut paste_col = self.cursor_col;
        let mut paste_row = self.cursor_row;

        for _ in 0..count {
            // Paste cells with formula adjustment
            for (r_offset, row_data) in clip.cells.iter().enumerate() {
                for (c_offset, (raw_input, _value)) in row_data.iter().enumerate() {
                    let dst_col = paste_col + c_offset;
                    let dst_row = paste_row + r_offset;

                    // Adjust formula if it starts with =
                    let adjusted = if raw_input.starts_with('=') {
                        let col_delta = (dst_col as isize) - (clip.start_col as isize) - (c_offset as isize);
                        let row_delta = (dst_row as isize) - (clip.start_row as isize) - (r_offset as isize);
                        formula::adjust_formula(raw_input, col_delta, row_delta)
                    } else {
                        raw_input.clone()
                    };

                    self.sheet.set_cell(dst_col, dst_row, adjusted);
                }
            }

            // Move for next paste based on axis
            match self.axis {
                EditAxis::Row => paste_col += clip.width,
                EditAxis::Column => paste_row += clip.height,
            }
        }

        self.last_paste_cols = clip.width;
        self.last_paste_rows = clip.height;

        let total = clip.width * clip.height * count;
        self.status_message = format!("Pasted {} cells", total);
    }

    /// Paste from system clipboard
    pub fn paste_from_system(&mut self) {
        let text = if let Ok(mut clipboard) = arboard::Clipboard::new() {
            clipboard.get_text().unwrap_or_default()
        } else {
            self.status_message = "Clipboard not available".to_string();
            return;
        };

        if text.is_empty() {
            self.status_message = "Clipboard is empty".to_string();
            return;
        }

        self.save_undo();

        // Parse TSV/CSV
        let lines: Vec<&str> = text.lines().collect();
        let mut height = 0;
        let mut width = 0;

        for (r_offset, line) in lines.iter().enumerate() {
            // Detect delimiter (tab first, then comma)
            let cells: Vec<&str> = if line.contains('\t') {
                line.split('\t').collect()
            } else {
                line.split(',').collect()
            };

            for (c_offset, cell_value) in cells.iter().enumerate() {
                let dst_col = self.cursor_col + c_offset;
                let dst_row = self.cursor_row + r_offset;
                self.sheet.set_cell(dst_col, dst_row, cell_value.to_string());
                width = width.max(c_offset + 1);
            }
            height = r_offset + 1;
        }

        self.last_paste_cols = width;
        self.last_paste_rows = height;

        self.status_message = format!("Pasted {}x{} cells from clipboard", width, height);
    }
}

fn handle_key(app: &mut App, key: KeyEvent) {
    // Handle register pending ("*)
    if app.register_pending {
        match key.code {
            KeyCode::Char('*') => {
                // "* - system clipboard register selected
                app.status_message = "\"* ...".to_string();
                // Keep register_pending true, but now we expect y or p
                return;
            }
            KeyCode::Char('y') => {
                // Check if it's "*y (status shows "*)
                if app.status_message.starts_with("\"*") {
                    app.yank_to_system();
                } else {
                    // Just " followed by y - treat as normal yank
                    app.yank();
                }
                app.register_pending = false;
                return;
            }
            KeyCode::Char('p') => {
                // Check if it's "*p
                if app.status_message.starts_with("\"*") {
                    app.paste_from_system();
                } else {
                    // Just " followed by p - treat as normal paste
                    let count = app.get_count().max(1);
                    app.paste(count);
                }
                app.register_pending = false;
                return;
            }
            KeyCode::Esc => {
                app.register_pending = false;
                app.update_status();
                return;
            }
            _ => {
                app.register_pending = false;
                app.update_status();
                return;
            }
        }
    }

    // Handle slash commands (/c /r)
    if app.slash_pending {
        app.slash_pending = false;
        match key.code {
            KeyCode::Char('c') => {
                app.axis = EditAxis::Column;
                app.update_status();
                return;
            }
            KeyCode::Char('r') => {
                app.axis = EditAxis::Row;
                app.update_status();
                return;
            }
            _ => {
                app.update_status();
                return;
            }
        }
    }

    match app.mode {
        Mode::Normal => handle_normal_mode(app, key),
        Mode::EditSingle | Mode::EditContinuous | Mode::EditPreserve => handle_edit_mode(app, key),
        Mode::Command => handle_command_mode(app, key),
        Mode::Visual => handle_visual_mode(app, key),
    }
}

fn handle_mouse(app: &mut App, mouse: MouseEvent) {
    match mouse.kind {
        MouseEventKind::Down(MouseButton::Left) => {
            let result = app.screen_to_cell(mouse.column, mouse.row);
            if let Some((col, row)) = result {
                match app.mode {
                    Mode::Normal => {
                        app.move_cursor_to(col, row);
                    }
                    Mode::EditSingle | Mode::EditContinuous | Mode::EditPreserve => {
                        // Cancel edit on mouse click (restore original)
                        app.input_buffer.clear();
                        app.mode = Mode::Normal;
                        app.move_cursor_to(col, row);
                        app.update_status();
                    }
                    Mode::Command => {
                        // Exit command mode on mouse click
                        app.mode = Mode::Normal;
                        app.command_buffer.clear();
                        app.move_cursor_to(col, row);
                        app.update_status();
                    }
                    Mode::Visual => {
                        // Extend selection on mouse click
                        app.move_cursor_to(col, row);
                        let (min_col, min_row, max_col, max_row) = app.get_selection_bounds();
                        app.status_message = format!("VISUAL {}x{}", max_col - min_col + 1, max_row - min_row + 1);
                    }
                }
            }
        }
        MouseEventKind::ScrollUp => {
            // Scroll up 3 rows
            let scroll = 3;
            app.view_row = app.view_row.saturating_sub(scroll);
            app.cursor_row = app.cursor_row.saturating_sub(scroll);
        }
        MouseEventKind::ScrollDown => {
            // Scroll down 3 rows
            let scroll = 3;
            app.view_row = (app.view_row + scroll).min(9999);
            app.cursor_row = (app.cursor_row + scroll).min(9999);
        }
        _ => {}
    }
}

fn handle_normal_mode(app: &mut App, key: KeyEvent) {
    // Handle Ctrl combinations first
    if key.modifiers.contains(KeyModifiers::CONTROL) {
        // Get terminal height for page calculations
        let (_, term_height) = terminal::size().unwrap_or((80, 24));
        let page_size = (term_height as usize).saturating_sub(4);  // Grid height
        let half_page = page_size / 2;
        
        match key.code {
            KeyCode::Char('q') => app.running = false,
            KeyCode::Char('r') => app.redo(),
            KeyCode::Char('s') => {
                commands::execute_command(app, "w");
            }
            // Full page down - scroll view and cursor together
            KeyCode::Char('f') => {
                let count = app.get_count().max(1);
                let scroll = page_size * count;
                app.view_row = (app.view_row + scroll).min(9999);
                app.cursor_row = (app.cursor_row + scroll).min(9999);
            }
            // Full page up - scroll view and cursor together
            KeyCode::Char('b') => {
                let count = app.get_count().max(1);
                let scroll = page_size * count;
                app.view_row = app.view_row.saturating_sub(scroll);
                app.cursor_row = app.cursor_row.saturating_sub(scroll);
            }
            // Half page down - scroll view and cursor together
            KeyCode::Char('d') => {
                let count = app.get_count().max(1);
                let scroll = half_page * count;
                app.view_row = (app.view_row + scroll).min(9999);
                app.cursor_row = (app.cursor_row + scroll).min(9999);
            }
            // Half page up - scroll view and cursor together
            KeyCode::Char('u') => {
                let count = app.get_count().max(1);
                let scroll = half_page * count;
                app.view_row = app.view_row.saturating_sub(scroll);
                app.cursor_row = app.cursor_row.saturating_sub(scroll);
            }
            _ => {}
        }
        return;
    }

    match key.code {
        // Count prefix
        KeyCode::Char(c @ '1'..='9') => {
            app.count_buffer.push(c);
        }
        KeyCode::Char('0') if !app.count_buffer.is_empty() => {
            app.count_buffer.push('0');
        }

        // Movement (always the same)
        KeyCode::Char('h') | KeyCode::Left => app.move_cursor(-1, 0),
        KeyCode::Char('j') | KeyCode::Down => app.move_cursor(0, 1),
        KeyCode::Char('k') | KeyCode::Up => app.move_cursor(0, -1),
        KeyCode::Char('l') | KeyCode::Right => app.move_cursor(1, 0),

        // Global movement
        KeyCode::Char('g') => {
            if app.pending_operator == Some('d') {
                // dg - waiting for second g
                app.pending_operator = Some('D');  // Use 'D' to indicate "dg" state
                app.status_message = "dg...".to_string();
            } else if app.pending_operator == Some('D') {
                // dgg - clear from sheet start to current
                app.clear_from_sheet_start();
                app.pending_operator = None;
            } else if app.pending_operator == Some('g') {
                // gg - goto A1
                app.cursor_col = 0;
                app.cursor_row = 0;
                app.adjust_view();
                app.pending_operator = None;
            } else {
                app.pending_operator = Some('g');
            }
        }
        KeyCode::Char('G') => {
            if app.pending_operator == Some('d') {
                // dG - clear from current to sheet end
                app.clear_to_sheet_end();
                app.pending_operator = None;
            } else {
                app.cursor_row = app.sheet.max_row().unwrap_or(0);
                app.cursor_col = app.sheet.max_col().unwrap_or(0);
                app.adjust_view();
            }
        }

        // Slash commands (/c /r)
        KeyCode::Char('/') => {
            app.slash_pending = true;
            app.status_message = "/ ...".to_string();
        }

        // Cell content operations
        KeyCode::Char('x') => {
            // Clear current cell
            app.clear_current_cell();
        }
        // = - Formula input (start with =)
        KeyCode::Char('=') => {
            let cell = app.sheet.get_cell(app.cursor_col, app.cursor_row);
            app.edit_original = cell.raw_input.clone();
            app.mode = Mode::EditSingle;
            app.input_buffer = "=".to_string();
            app.update_status();
        }
        // r - Single cell edit (return to Normal after Enter/arrows)
        KeyCode::Char('r') => {
            let cell = app.sheet.get_cell(app.cursor_col, app.cursor_row);
            app.edit_original = cell.raw_input.clone();
            app.mode = Mode::EditSingle;
            app.input_buffer.clear();
            app.update_status();
        }
        // R - Continuous edit (stay in edit mode after Enter/arrows)
        KeyCode::Char('R') => {
            let cell = app.sheet.get_cell(app.cursor_col, app.cursor_row);
            app.edit_original = cell.raw_input.clone();
            app.mode = Mode::EditContinuous;
            app.input_buffer.clear();
            app.update_status();
        }
        // F2 - Edit cell content (preserve existing content)
        KeyCode::F(2) => {
            let cell = app.sheet.get_cell(app.cursor_col, app.cursor_row);
            app.edit_original = cell.raw_input.clone();
            app.mode = Mode::EditPreserve;
            app.input_buffer = cell.raw_input.clone();
            app.update_status();
        }

        // Structure operations (axis-dependent)
        KeyCode::Char('i') => app.insert_at_cursor(),
        KeyCode::Char('I') => app.insert_at_start(),
        KeyCode::Char('a') => app.append_after_cursor(),
        KeyCode::Char('A') => app.goto_axis_end_next(),

        KeyCode::Char('d') => {
            if app.pending_operator == Some('d') {
                // dd - delete structure (row or column)
                app.delete_structure();
                app.pending_operator = None;
            } else {
                // Start d operator, wait for motion or second d
                app.pending_operator = Some('d');
                app.status_message = "d...".to_string();
            }
        }

        KeyCode::Char('$') => {
            if app.pending_operator == Some('d') {
                // d$ - clear to axis end
                app.clear_to_axis_end();
                app.pending_operator = None;
            } else {
                app.goto_axis_end();
            }
        }

        KeyCode::Char('0') => {
            if app.pending_operator == Some('d') {
                // d0 - clear from axis start
                app.clear_from_axis_start();
                app.pending_operator = None;
            } else if app.count_buffer.is_empty() {
                app.goto_axis_start();
            } else {
                app.count_buffer.push('0');
            }
        }

        KeyCode::Char('^') => {
            if app.pending_operator == Some('d') {
                // d^ - clear from first non-empty
                app.clear_from_first_non_empty();
                app.pending_operator = None;
            } else {
                app.goto_first_non_empty();
            }
        }

        KeyCode::Char('o') => app.insert_structure_after(),
        KeyCode::Char('O') => app.insert_structure_before(),

        // Column width adjustment
        KeyCode::Char('<') => {
            let count = app.get_count() as isize;
            app.sheet.adjust_col_width(app.cursor_col, -count);
            let width = app.sheet.get_col_width(app.cursor_col);
            app.status_message = format!("Column width: {}", width);
        }
        KeyCode::Char('>') => {
            let count = app.get_count() as isize;
            app.sheet.adjust_col_width(app.cursor_col, count);
            let width = app.sheet.get_col_width(app.cursor_col);
            app.status_message = format!("Column width: {}", width);
        }

        // Undo
        KeyCode::Char('u') => app.undo(),

        // Yank (copy)
        KeyCode::Char('y') => {
            app.yank();
        }

        // Paste
        KeyCode::Char('p') => {
            let count = app.get_count().max(1);
            app.paste(count);
        }

        // Register prefix (for "*)
        KeyCode::Char('"') => {
            app.register_pending = true;
            app.status_message = "\" ...".to_string();
        }

        // Visual mode
        KeyCode::Char('v') => {
            app.mode = Mode::Visual;
            app.visual_start_col = app.cursor_col;
            app.visual_start_row = app.cursor_row;
            app.status_message = "-- VISUAL --".to_string();
        }

        // Visual Line/Column mode (V)
        KeyCode::Char('V') => {
            app.mode = Mode::Visual;
            match app.axis {
                EditAxis::Row => {
                    // Select entire row
                    app.visual_start_col = 0;
                    app.visual_start_row = app.cursor_row;
                    app.cursor_col = app.sheet.max_col().unwrap_or(255);
                    app.status_message = "-- VISUAL LINE --".to_string();
                }
                EditAxis::Column => {
                    // Select entire column
                    app.visual_start_col = app.cursor_col;
                    app.visual_start_row = 0;
                    app.cursor_row = app.sheet.max_row().unwrap_or(9999);
                    app.status_message = "-- VISUAL COLUMN --".to_string();
                }
            }
        }

        // Search next/prev
        KeyCode::Char('n') => {
            commands::search_next(app);
        }
        KeyCode::Char('N') => {
            commands::search_prev(app);
        }

        // Command mode
        KeyCode::Char(':') => {
            app.mode = Mode::Command;
            app.command_buffer.clear();
            app.status_message = ":".to_string();
        }

        KeyCode::Esc => {
            app.pending_operator = None;
            app.count_buffer.clear();
            app.slash_pending = false;
            app.register_pending = false;
            app.update_status();
        }

        _ => {}
    }
}

fn handle_command_mode(app: &mut App, key: KeyEvent) {
    match key.code {
        KeyCode::Esc => {
            app.mode = Mode::Normal;
            app.command_buffer.clear();
            app.update_status();
        }
        KeyCode::Enter => {
            let cmd = app.command_buffer.clone();
            app.mode = Mode::Normal;
            commands::execute_command(app, &cmd);
            app.command_buffer.clear();
        }
        KeyCode::Backspace => {
            app.command_buffer.pop();
            if app.command_buffer.is_empty() {
                app.mode = Mode::Normal;
                app.update_status();
            } else {
                app.status_message = format!(":{}", app.command_buffer);
            }
        }
        KeyCode::Tab => {
            // Tab completion for file names
            complete_filename(app);
        }
        KeyCode::Char(c) => {
            app.command_buffer.push(c);
            app.status_message = format!(":{}", app.command_buffer);
        }
        _ => {}
    }
}

fn handle_visual_mode(app: &mut App, key: KeyEvent) {
    match key.code {
        KeyCode::Esc => {
            app.mode = Mode::Normal;
            app.update_status();
        }

        // Movement - extend selection
        KeyCode::Char('h') | KeyCode::Left => {
            app.cursor_col = app.cursor_col.saturating_sub(1);
            app.adjust_view();
            update_visual_status(app);
        }
        KeyCode::Char('j') | KeyCode::Down => {
            app.cursor_row = (app.cursor_row + 1).min(9999);
            app.adjust_view();
            update_visual_status(app);
        }
        KeyCode::Char('k') | KeyCode::Up => {
            app.cursor_row = app.cursor_row.saturating_sub(1);
            app.adjust_view();
            update_visual_status(app);
        }
        KeyCode::Char('l') | KeyCode::Right => {
            app.cursor_col = (app.cursor_col + 1).min(255);
            app.adjust_view();
            update_visual_status(app);
        }

        // Jump to start/end
        KeyCode::Char('0') => {
            app.cursor_col = 0;
            app.adjust_view();
            update_visual_status(app);
        }
        KeyCode::Char('$') => {
            if let Some(max_col) = app.sheet.max_col_in_row(app.cursor_row) {
                app.cursor_col = max_col;
            }
            app.adjust_view();
            update_visual_status(app);
        }
        KeyCode::Char('g') => {
            app.cursor_col = 0;
            app.cursor_row = 0;
            app.adjust_view();
            update_visual_status(app);
        }
        KeyCode::Char('G') => {
            app.cursor_row = app.sheet.max_row().unwrap_or(0);
            app.cursor_col = app.sheet.max_col().unwrap_or(0);
            app.adjust_view();
            update_visual_status(app);
        }

        // Yank (copy) selection
        KeyCode::Char('y') => {
            app.yank();
        }

        // Actions on selection
        KeyCode::Char('d') | KeyCode::Char('x') => {
            app.clear_selection();
        }

        // Register prefix for system clipboard
        KeyCode::Char('"') => {
            app.register_pending = true;
            app.status_message = "\" ...".to_string();
        }

        _ => {}
    }
}

fn update_visual_status(app: &mut App) {
    let (min_col, min_row, max_col, max_row) = app.get_selection_bounds();
    let cols = max_col - min_col + 1;
    let rows = max_row - min_row + 1;
    app.status_message = format!("-- VISUAL -- {}x{}", cols, rows);
}

/// Complete filename in command buffer
fn complete_filename(app: &mut App) {
    let cmd = &app.command_buffer;
    
    // Check if command is :e or :w with partial filename
    let (prefix, partial) = if cmd.starts_with("e ") {
        ("e ", &cmd[2..])
    } else if cmd.starts_with("w ") {
        ("w ", &cmd[2..])
    } else if cmd == "e" {
        ("e ", "")
    } else if cmd == "w" {
        ("w ", "")
    } else {
        return;
    };
    
    // Get directory and file prefix
    let (dir, file_prefix) = if partial.contains('/') || partial.contains('\\') {
        let path = std::path::Path::new(partial);
        if let Some(parent) = path.parent() {
            (parent.to_string_lossy().to_string(), 
             path.file_name().map(|f| f.to_string_lossy().to_string()).unwrap_or_default())
        } else {
            (".".to_string(), partial.to_string())
        }
    } else {
        (".".to_string(), partial.to_string())
    };
    
    // List matching files
    let matches: Vec<String> = match std::fs::read_dir(&dir) {
        Ok(entries) => {
            entries
                .filter_map(|e| e.ok())
                .filter_map(|e| {
                    let name = e.file_name().to_string_lossy().to_string();
                    let lower_name = name.to_lowercase();
                    let lower_prefix = file_prefix.to_lowercase();
                    
                    // Match files starting with prefix
                    if lower_name.starts_with(&lower_prefix) {
                        // Filter for spreadsheet-related extensions or directories
                        let is_dir = e.file_type().map(|t| t.is_dir()).unwrap_or(false);
                        let ext = std::path::Path::new(&name)
                            .extension()
                            .map(|e| e.to_string_lossy().to_lowercase())
                            .unwrap_or_default();
                        
                        if is_dir || ext == "json" || ext == "csv" || file_prefix.is_empty() {
                            if dir == "." {
                                Some(name)
                            } else {
                                Some(format!("{}/{}", dir, name))
                            }
                        } else {
                            None
                        }
                    } else {
                        None
                    }
                })
                .collect()
        }
        Err(_) => Vec::new(),
    };
    
    if matches.len() == 1 {
        // Single match - complete it
        app.command_buffer = format!("{}{}", prefix, matches[0]);
        app.status_message = format!(":{}", app.command_buffer);
    } else if matches.len() > 1 {
        // Multiple matches - find common prefix and show options
        let common = find_common_prefix(&matches);
        if common.len() > partial.len() {
            app.command_buffer = format!("{}{}", prefix, common);
        }
        // Limit displayed matches to avoid overflow
        let display_matches: Vec<&str> = matches.iter()
            .take(5)
            .map(|s| s.as_str())
            .collect();
        let suffix = if matches.len() > 5 {
            format!(" +{}", matches.len() - 5)
        } else {
            String::new()
        };
        app.status_message = format!(":{} [{}{}]", app.command_buffer, display_matches.join(" "), suffix);
    } else if !partial.is_empty() {
        // No matches - keep as is (will create new file)
        app.status_message = format!(":{} (new file)", app.command_buffer);
    }
}

/// Find common prefix among strings
fn find_common_prefix(strings: &[String]) -> String {
    if strings.is_empty() {
        return String::new();
    }
    
    let first = &strings[0];
    let mut prefix_len = first.len();
    
    for s in strings.iter().skip(1) {
        prefix_len = first.chars()
            .zip(s.chars())
            .take_while(|(a, b)| a.to_lowercase().eq(b.to_lowercase()))
            .count()
            .min(prefix_len);
    }
    
    first[..prefix_len].to_string()
}

fn handle_edit_mode(app: &mut App, key: KeyEvent) {
    let current_mode = app.mode;
    
    match key.code {
        KeyCode::Esc => {
            // Cancel - restore original content
            app.input_buffer.clear();
            app.mode = Mode::Normal;
            app.update_status();
        }
        KeyCode::Enter => {
            // Commit if there's input, then move/exit based on mode
            if !app.input_buffer.is_empty() {
                app.save_undo();
                app.sheet.set_cell(app.cursor_col, app.cursor_row, app.input_buffer.clone());
            }
            app.input_buffer.clear();
            
            match current_mode {
                Mode::EditSingle | Mode::EditPreserve => {
                    // Return to Normal mode
                    app.mode = Mode::Normal;
                    app.update_status();
                }
                Mode::EditContinuous => {
                    // Move to next cell and continue editing
                    match app.axis {
                        EditAxis::Row => app.cursor_col = (app.cursor_col + 1).min(255),
                        EditAxis::Column => app.cursor_row = (app.cursor_row + 1).min(9999),
                    }
                    app.adjust_view();
                    // Store new cell's original content
                    let cell = app.sheet.get_cell(app.cursor_col, app.cursor_row);
                    app.edit_original = cell.raw_input.clone();
                }
                _ => {}
            }
        }
        KeyCode::Backspace => {
            app.input_buffer.pop();
        }
        KeyCode::Delete => {
            if !app.input_buffer.is_empty() {
                app.input_buffer.remove(0);
            }
        }
        KeyCode::Char(c) => {
            app.input_buffer.push(c);
        }
        KeyCode::Up | KeyCode::Down | KeyCode::Left | KeyCode::Right | KeyCode::Tab | KeyCode::BackTab => {
            // Calculate new position
            let (new_col, new_row) = match key.code {
                KeyCode::Up => (app.cursor_col, app.cursor_row.saturating_sub(1)),
                KeyCode::Down => (app.cursor_col, (app.cursor_row + 1).min(9999)),
                KeyCode::Left => (app.cursor_col.saturating_sub(1), app.cursor_row),
                KeyCode::Right | KeyCode::Tab => ((app.cursor_col + 1).min(255), app.cursor_row),
                KeyCode::BackTab => (app.cursor_col.saturating_sub(1), app.cursor_row),
                _ => (app.cursor_col, app.cursor_row),
            };
            
            // Commit if there's input
            if !app.input_buffer.is_empty() {
                app.save_undo();
                app.sheet.set_cell(app.cursor_col, app.cursor_row, app.input_buffer.clone());
            }
            app.input_buffer.clear();
            
            match current_mode {
                Mode::EditSingle | Mode::EditPreserve => {
                    // Return to Normal mode and move
                    app.mode = Mode::Normal;
                    app.cursor_col = new_col;
                    app.cursor_row = new_row;
                    app.adjust_view();
                    app.update_status();
                }
                Mode::EditContinuous => {
                    // Move and continue editing
                    app.cursor_col = new_col;
                    app.cursor_row = new_row;
                    app.adjust_view();
                    // Store new cell's original content
                    let cell = app.sheet.get_cell(app.cursor_col, app.cursor_row);
                    app.edit_original = cell.raw_input.clone();
                }
                _ => {}
            }
        }
        _ => {}
    }
}

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().collect();
    
    let mut stdout = stdout();
    terminal::enable_raw_mode()?;
    execute!(stdout, EnterAlternateScreen, Hide, EnableMouseCapture)?;

    let mut app = App::new();

    // Open file from command line argument
    if args.len() > 1 {
        let filename = &args[1];
        let cmd = format!("e {}", filename);
        commands::execute_command(&mut app, &cmd);
    }

    UI::draw(&app)?;

    while app.running {
        if event::poll(std::time::Duration::from_millis(100))? {
            match event::read()? {
                Event::Key(key) => {
                    if key.kind == event::KeyEventKind::Press {
                        handle_key(&mut app, key);
                        UI::draw(&app)?;
                    }
                }
                Event::Mouse(mouse) => {
                    handle_mouse(&mut app, mouse);
                    UI::draw(&app)?;
                }
                Event::Resize(_, _) => {
                    UI::draw(&app)?;
                }
                _ => {}
            }
        }
    }

    execute!(stdout, Show, DisableMouseCapture, LeaveAlternateScreen)?;
    terminal::disable_raw_mode()?;
    Ok(())
}
