use crossterm::{
    cursor::{Hide, MoveTo, Show},
    queue,
    style::{Color, ResetColor, SetBackgroundColor, SetForegroundColor},
    terminal,
};
use std::io::{stdout, Result, Write};
use unicode_width::UnicodeWidthStr;

use crate::{App, Mode, EditAxis};
use crate::cell::CellValue;
use crate::formula;

const ROW_LABEL_WIDTH: usize = 5;

// Colors
const GREEN: Color = Color::Rgb { r: 0, g: 170, b: 0 };
const ORANGE: Color = Color::Rgb { r: 255, g: 136, b: 0 };
const FRAME_COLOR: Color = Color::Rgb { r: 180, g: 180, b: 180 };

// Box drawing characters
const BOX_VERTICAL: char = '│';

/// Truncate string to fit within max_width (display width) - keeps left side
fn truncate_to_width(s: &str, max_width: usize) -> String {
    let mut result = String::new();
    let mut width = 0;
    for c in s.chars() {
        let w = unicode_width::UnicodeWidthChar::width(c).unwrap_or(1);
        if width + w > max_width {
            break;
        }
        result.push(c);
        width += w;
    }
    result
}

/// Truncate string to fit within max_width - keeps right side (for editing)
fn truncate_from_end(s: &str, max_width: usize) -> String {
    let total_width = UnicodeWidthStr::width(s);
    if total_width <= max_width {
        return s.to_string();
    }
    
    // Need to skip (total_width - max_width) worth of characters from start
    let skip_width = total_width - max_width;
    let mut skipped = 0;
    let mut result = String::new();
    
    for c in s.chars() {
        let w = unicode_width::UnicodeWidthChar::width(c).unwrap_or(1);
        if skipped < skip_width {
            skipped += w;
        } else {
            result.push(c);
        }
    }
    result
}

/// Pad string to target display width
fn pad_to_width(s: &str, target_width: usize, align_right: bool) -> String {
    let current = UnicodeWidthStr::width(s);
    if current >= target_width {
        return truncate_to_width(s, target_width);
    }
    let padding = target_width - current;
    if align_right {
        format!("{}{}", " ".repeat(padding), s)
    } else {
        format!("{}{}", s, " ".repeat(padding))
    }
}

/// Get display width of a string
fn display_width(s: &str) -> usize {
    UnicodeWidthStr::width(s)
}

pub struct UI;

impl UI {
    fn cursor_color(mode: Mode) -> Color {
        match mode {
            Mode::Normal => GREEN,
            Mode::EditSingle | Mode::EditContinuous | Mode::EditPreserve => ORANGE,
            Mode::Command => GREEN,
            Mode::Visual => Color::Rgb { r: 100, g: 100, b: 255 },
        }
    }

    const SELECTION_BG: Color = Color::Rgb { r: 60, g: 60, b: 120 };

    /// Calculate how many columns fit in the terminal and their positions
    fn calc_visible_cols(app: &App, term_width: usize) -> Vec<(usize, usize)> {
        // Returns Vec of (col_index, col_width)
        let mut cols = Vec::new();
        let mut used_width = ROW_LABEL_WIDTH;
        let mut col = app.view_col;
        
        while used_width < term_width && col <= 255 {
            let col_width = app.sheet.get_col_width(col);
            if used_width + col_width > term_width {
                break;
            }
            cols.push((col, col_width));
            used_width += col_width;
            col += 1;
        }
        
        cols
    }

    pub fn draw(app: &App) -> Result<()> {
        let mut stdout = stdout();
        let (term_width, term_height) = terminal::size()?;
        let grid_height = (term_height as usize).saturating_sub(4);
        let visible_cols = Self::calc_visible_cols(app, term_width as usize);

        let cursor_color = Self::cursor_color(app.mode);

        queue!(stdout, Hide)?;
        queue!(stdout, MoveTo(0, 0))?;

        Self::draw_status_bar(&mut stdout, app, term_width)?;
        Self::draw_column_headers(&mut stdout, app, &visible_cols, term_width)?;
        Self::draw_grid(&mut stdout, app, grid_height, &visible_cols, term_width, cursor_color)?;
        Self::draw_formula_bar(&mut stdout, app, term_height, term_width)?;

        queue!(stdout, Show)?;
        
        stdout.flush()?;
        Ok(())
    }

    fn draw_status_bar(stdout: &mut std::io::Stdout, app: &App, term_width: u16) -> Result<()> {
        queue!(
            stdout,
            MoveTo(0, 0),
            SetBackgroundColor(GREEN),
            SetForegroundColor(Color::Black),
        )?;

        let cell_name = formula::cell_name(app.cursor_col, app.cursor_row);
        let cell = app.sheet.get_cell(app.cursor_col, app.cursor_row);
        
        let value_display = match &cell.value {
            CellValue::Formula(_) => {
                let evaluated = app.sheet.evaluate(app.cursor_col, app.cursor_row);
                format!("{} → {}", cell.raw_input, evaluated)
            }
            _ => app.sheet.evaluate(app.cursor_col, app.cursor_row),
        };

        let mode_str = match app.mode {
            Mode::Normal => "NORMAL",
            Mode::EditSingle => "EDIT",
            Mode::EditContinuous => "EDIT+",
            Mode::EditPreserve => "EDIT",
            Mode::Command => "COMMAND",
            Mode::Visual => "VISUAL",
        };

        let axis_str = match app.axis {
            EditAxis::Row => "Row",
            EditAxis::Column => "Col",
        };

        let file_str = app.current_file.as_deref().unwrap_or("[New]");

        let left = format!(" {} | {} ", cell_name, value_display);
        let right = format!(" {} | {} | {} ", axis_str, mode_str, file_str);
        
        // Use display width for proper padding calculation
        let left_width = display_width(&left);
        let right_width = display_width(&right);
        let padding = (term_width as usize).saturating_sub(left_width + right_width);

        write!(stdout, "{}{:width$}{}", left, "", right, width = padding)?;
        queue!(stdout, ResetColor)?;
        Ok(())
    }

    fn draw_column_headers(stdout: &mut std::io::Stdout, _app: &App, visible_cols: &[(usize, usize)], term_width: u16) -> Result<()> {
        queue!(
            stdout,
            MoveTo(0, 1),
            SetBackgroundColor(GREEN),
            SetForegroundColor(Color::Black),
        )?;

        write!(stdout, "{:width$}", "", width = ROW_LABEL_WIDTH)?;

        let mut used = ROW_LABEL_WIDTH;
        for &(col, col_width) in visible_cols {
            let col_name = formula::col_to_name(col);
            write!(stdout, "{:^width$}", col_name, width = col_width)?;
            used += col_width;
        }

        let remaining = (term_width as usize).saturating_sub(used);
        write!(stdout, "{:width$}", "", width = remaining)?;

        queue!(stdout, ResetColor)?;
        Ok(())
    }

    fn draw_grid(stdout: &mut std::io::Stdout, app: &App, grid_height: usize, visible_cols: &[(usize, usize)], term_width: u16, cursor_color: Color) -> Result<()> {
        for row in 0..grid_height {
            let actual_row = app.view_row + row;
            
            queue!(stdout, MoveTo(0, (row + 2) as u16))?;

            // Row label
            queue!(
                stdout,
                SetBackgroundColor(GREEN),
                SetForegroundColor(Color::Black),
            )?;
            write!(stdout, "{:>width$}", actual_row + 1, width = ROW_LABEL_WIDTH)?;
            queue!(stdout, ResetColor)?;

            let mut used = ROW_LABEL_WIDTH;

            // Get selection bounds for Visual mode
            let (sel_min_col, sel_min_row, sel_max_col, sel_max_row) = if app.mode == Mode::Visual {
                app.get_selection_bounds()
            } else {
                (usize::MAX, usize::MAX, 0, 0)  // No selection
            };

            // Cells
            for &(actual_col, col_width) in visible_cols {
                let is_cursor = actual_col == app.cursor_col && actual_row == app.cursor_row;
                let is_current_col = actual_col == app.cursor_col;
                let is_selected = app.mode == Mode::Visual 
                    && actual_col >= sel_min_col && actual_col <= sel_max_col
                    && actual_row >= sel_min_row && actual_row <= sel_max_row;

                // Flag for edit mode cursor
                let is_editing = is_cursor && matches!(app.mode, Mode::EditSingle | Mode::EditContinuous | Mode::EditPreserve);

                // Get cell value and type
                let cell = app.sheet.get_cell(actual_col, actual_row);
                let is_number = matches!(cell.value, CellValue::Number(_) | CellValue::Formula(_));

                // Column mode: draw with frame
                if app.axis == EditAxis::Column && is_current_col && !is_cursor && !is_selected {
                    // Inner width = col_width - 2 (for borders)
                    let inner_width = col_width.saturating_sub(2);
                    
                    // Get content
                    let value = app.sheet.evaluate(actual_col, actual_row);
                    let content = if display_width(&value) > inner_width {
                        if is_number {
                            "#".repeat(inner_width)
                        } else {
                            let truncated = truncate_to_width(&value, inner_width.saturating_sub(1));
                            format!("{}…", truncated)
                        }
                    } else {
                        value
                    };
                    
                    // Format with proper width
                    let formatted = if is_number {
                        pad_to_width(&content, inner_width, true)
                    } else {
                        pad_to_width(&content, inner_width, false)
                    };
                    
                    // Left border
                    queue!(stdout, SetBackgroundColor(Color::Black), SetForegroundColor(FRAME_COLOR))?;
                    write!(stdout, "{}", BOX_VERTICAL)?;
                    
                    // Content
                    queue!(stdout, SetForegroundColor(GREEN))?;
                    write!(stdout, "{}", formatted)?;
                    
                    // Right border
                    queue!(stdout, SetForegroundColor(FRAME_COLOR))?;
                    write!(stdout, "{}", BOX_VERTICAL)?;
                } else {
                    // Content width = col_width - 1 (right padding)
                    let content_width = col_width.saturating_sub(1);
                    
                    // Get content
                    let content = if is_editing {
                        let input = &app.input_buffer;
                        let input_width = display_width(input);
                        // Reserve 1 char for cursor indicator
                        let available_width = content_width.saturating_sub(1);
                        if input_width > available_width {
                            // Show end of input (slide to cursor position)
                            let truncated = truncate_from_end(input, available_width);
                            format!("{}▏", truncated)
                        } else {
                            format!("{}▏", input)
                        }
                    } else {
                        let value = app.sheet.evaluate(actual_col, actual_row);
                        if display_width(&value) > content_width {
                            if is_number {
                                "#".repeat(content_width)
                            } else {
                                let truncated = truncate_to_width(&value, content_width.saturating_sub(1));
                                format!("{}…", truncated)
                            }
                        } else {
                            value
                        }
                    };
                    
                    // Set colors based on cell type
                    let (bg, fg) = if is_cursor {
                        (cursor_color, Color::Black)
                    } else if is_selected {
                        (Self::SELECTION_BG, Color::White)
                    } else {
                        (Color::Black, GREEN)
                    };
                    
                    queue!(stdout, SetBackgroundColor(bg), SetForegroundColor(fg))?;
                    
                    // Format and write
                    let formatted = if is_number && !is_editing {
                        pad_to_width(&content, content_width, true)
                    } else {
                        pad_to_width(&content, content_width, false)
                    };
                    write!(stdout, "{} ", formatted)?;  // +1 for right padding
                }

                queue!(stdout, ResetColor)?;
                used += col_width;
            }
            
            // Clear rest of line
            let remaining = (term_width as usize).saturating_sub(used);
            if remaining > 0 {
                queue!(stdout, SetBackgroundColor(Color::Black))?;
                write!(stdout, "{:width$}", "", width = remaining)?;
                queue!(stdout, ResetColor)?;
            }
        }

        Ok(())
    }

    fn draw_formula_bar(stdout: &mut std::io::Stdout, app: &App, term_height: u16, term_width: u16) -> Result<()> {
        queue!(
            stdout,
            MoveTo(0, term_height - 2),
            SetBackgroundColor(GREEN),
            SetForegroundColor(Color::Black),
        )?;

        let content = match app.mode {
            Mode::EditSingle | Mode::EditContinuous | Mode::EditPreserve => {
                format!(" fx: {}_ ", app.input_buffer)
            }
            Mode::Command => {
                format!(" :{}_ ", app.command_buffer)
            }
            Mode::Visual => {
                let (min_col, min_row, max_col, max_row) = app.get_selection_bounds();
                let start = crate::formula::cell_name(min_col, min_row);
                let end = crate::formula::cell_name(max_col, max_row);
                format!(" Selection: {}:{} ", start, end)
            }
            Mode::Normal => {
                let cell = app.sheet.get_cell(app.cursor_col, app.cursor_row);
                format!(" fx: {} ", cell.raw_input)
            }
        };

        let content_width = display_width(&content);
        let display = if content_width > term_width as usize {
            let truncated = truncate_to_width(&content, term_width as usize - 3);
            format!("{}...", truncated)
        } else {
            pad_to_width(&content, term_width as usize, false)
        };
        
        write!(stdout, "{}", display)?;
        queue!(stdout, ResetColor)?;

        // Status line - generate real-time instead of using cached status_message
        queue!(
            stdout,
            MoveTo(0, term_height - 1),
            SetBackgroundColor(Color::Black),
            SetForegroundColor(GREEN),
        )?;
        
        let mode_str = match app.mode {
            Mode::Normal => "NORMAL",
            Mode::EditSingle => "EDIT",
            Mode::EditContinuous => "EDIT+",
            Mode::EditPreserve => "EDIT",
            Mode::Command => "COMMAND",
            Mode::Visual => "VISUAL",
        };
        let axis_str = match app.axis {
            crate::EditAxis::Row => "Row",
            crate::EditAxis::Column => "Col",
        };
        let cell_name = crate::formula::cell_name(app.cursor_col, app.cursor_row);
        let file_str = app.current_file.as_deref().unwrap_or("[New]");
        
        let status = format!("{} | {} | {} | {}", mode_str, cell_name, axis_str, file_str);
        let status_display = pad_to_width(&status, term_width as usize, false);
        write!(stdout, "{}", status_display)?;
        queue!(stdout, ResetColor)?;

        Ok(())
    }
}
