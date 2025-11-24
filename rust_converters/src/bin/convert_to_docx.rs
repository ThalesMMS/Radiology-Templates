use anyhow::Result;
use docx_rust::document::{Paragraph, Run};
use docx_rust::formatting::{CharacterProperty, Fonts, Justification, JustificationVal, ParagraphProperty};
use docx_rust::Docx;
use std::fs;
use std::path::{Path, PathBuf};

const FONT_NAME: &str = "Arial";
const FONT_SIZE_PT: i32 = 10;
const SOURCE_DIR: &str = "Templates_markdown";
const TARGET_DIR: &str = "Templates_docx";

#[derive(Clone, Copy, Debug)]
enum Alignment {
    Justify,
    Center,
}

fn normalize_heading(line: &str) -> (String, bool) {
    let stripped = line.trim_start();
    if stripped.starts_with('#') {
        let text = stripped.trim_start_matches('#').trim().to_string();
        (text, true)
    } else {
        (line.to_string(), false)
    }
}

fn append_run<'a>(
    para: Paragraph<'a>,
    text: &str,
    bold: bool,
    italic: bool,
    force_italic: bool,
    font_size_pt: i32,
) -> Paragraph<'a> {
    if text.is_empty() {
        return para;
    }

    let mut prop = CharacterProperty::default();
    let fonts = Fonts::default().ascii(FONT_NAME.to_string());
    // In DOCX, font size is in half-points.
    let size_half_points = (font_size_pt * 2) as isize;

    prop = prop.fonts(fonts).size(size_half_points);
    if bold {
        prop = prop.bold(true);
    }
    if italic || force_italic {
        prop = prop.italics(true);
    }

    let run = Run::default()
        .property(prop)
        .push_text(text.to_string());

    para.push(run)
}

fn add_markdown_paragraph<'a>(
    docx: &mut Docx<'a>,
    raw_line: &str,
    alignment: Alignment,
    force_italic: bool,
    font_size_pt: i32,
) {
    let (text, heading) = normalize_heading(raw_line);

    let justification_val = match alignment {
        Alignment::Center => JustificationVal::Center,
        Alignment::Justify => JustificationVal::Both,
    };

    let para_prop = ParagraphProperty::default().justification(Justification::from(justification_val));
    let mut para = Paragraph::default().property(para_prop);

    if text.is_empty() {
        // Empty paragraph
        para = para.push_text(String::new());
        docx.document.push(para);
        return;
    }

    let chars: Vec<char> = text.chars().collect();
    let mut buffer = String::new();
    let mut bold = heading;
    let mut italic = false;
    let mut i = 0;

    while i < chars.len() {
        let c = chars[i];

        // ** ou __ → toggle bold
        if (c == '*' || c == '_') && i + 1 < chars.len() && chars[i + 1] == c {
            para = append_run(para, &buffer, bold, italic, force_italic, font_size_pt);
            buffer.clear();
            bold = !bold;
            i += 2;
            continue;
        }

        // * ou _ simples → toggle italic
        if c == '*' || c == '_' {
            para = append_run(para, &buffer, bold, italic, force_italic, font_size_pt);
            buffer.clear();
            italic = !italic;
            i += 1;
            continue;
        }

        buffer.push(c);
        i += 1;
    }

    para = append_run(para, &buffer, bold, italic, force_italic, font_size_pt);
    docx.document.push(para);
}

fn convert_file(md_path: &Path, output_path: &Path) -> Result<()> {
    let mut docx: Docx = Docx::default();

    let content = fs::read_to_string(md_path)?;
    let lines: Vec<&str> = content.lines().collect();

    if lines.is_empty() {
        let para_prop =
            ParagraphProperty::default().justification(Justification::from(JustificationVal::Both));
        let para = Paragraph::default()
            .property(para_prop)
            .push_text(String::new());
        docx.document.push(para);
    } else {
        // first_written: index of the first non-empty line
        let first_written = lines
            .iter()
            .enumerate()
            .find(|(_, line)| !line.trim().is_empty())
            .map(|(i, _)| i);

        // last_written: index of the last non-empty line
        let last_written = lines
            .iter()
            .enumerate()
            .rev()
            .find(|(_, line)| !line.trim().is_empty())
            .map(|(i, _)| i);

        for (idx, line) in lines.iter().enumerate() {
            let mut alignment = Alignment::Justify;
            let mut force_italic = false;
            let mut font_size_pt = FONT_SIZE_PT;

            if Some(idx) == first_written {
                alignment = Alignment::Center;
            }
            if Some(idx) == last_written {
                alignment = Alignment::Center;
                force_italic = true;
                font_size_pt = 8;
            }

            add_markdown_paragraph(
                &mut docx,
                line,
                alignment,
                force_italic,
                font_size_pt,
            );
        }
    }

    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)?;
    }
    docx.write_file(output_path.to_string_lossy().as_ref())?;
    Ok(())
}

fn main() -> Result<()> {
    let source_dir = PathBuf::from(SOURCE_DIR);
    if !source_dir.exists() {
        anyhow::bail!("Source folder not found: {}", source_dir.display());
    }

    let target_dir = PathBuf::from(TARGET_DIR);
    fs::create_dir_all(&target_dir)?;

    let mut md_files: Vec<PathBuf> = fs::read_dir(&source_dir)?
        .filter_map(|entry| entry.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("md"))
        .collect();
    md_files.sort();

    for md_file in md_files {
        let output_file = target_dir.join(
            md_file
                .file_stem()
                .expect("md file without stem")
                .to_string_lossy()
                .to_string()
                + ".docx",
        );
        convert_file(&md_file, &output_file)?;
    }

    Ok(())
}
