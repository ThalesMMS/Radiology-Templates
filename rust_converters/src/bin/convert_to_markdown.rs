use anyhow::Result;
use docx_rust::document::{BodyContent, Paragraph, ParagraphContent};
use docx_rust::formatting::{Bold, Italics, Underline, UnderlineStyle};
use docx_rust::DocxFile;
use regex::Regex;
use std::fs;
use std::path::{Path, PathBuf};

fn convert_docx_to_markdown(docx_path: &Path) -> Result<String> {
    let file = DocxFile::from_file(docx_path)?;
    let docx = file.parse()?;

    let mut markdown_lines: Vec<String> = Vec::new();

    // Walk the document body in order.
    // Only handle paragraphs; tables/SDT/etc. are ignored.
    let body = &docx.document.body;
    for item in &body.content {
        match item {
            BodyContent::Paragraph(p) => {
                markdown_lines.push(paragraph_to_markdown(p));
            }
            // Se quiser mapear tabelas em Markdown, implementar aqui
            BodyContent::Table(_t) => {
                // Ignored for simplicity — the Python version outputs table Markdown.
            }
            _ => {}
        }
    }

    Ok(markdown_lines.join("\n"))
}

fn bold_is_on(flag: &Option<Bold>) -> bool {
    flag.as_ref()
        .map(|b| b.value.unwrap_or(true))
        .unwrap_or(false)
}

fn italics_is_on(flag: &Option<Italics>) -> bool {
    flag.as_ref()
        .map(|i| i.value.unwrap_or(true))
        .unwrap_or(false)
}

fn underline_is_on(flag: &Option<Underline>) -> bool {
    flag.as_ref()
        .map(|u| match u.val.as_ref() {
            Some(UnderlineStyle::None) => false,
            Some(_) => true,
            None => true, // element is present, default style is "single"
        })
        .unwrap_or(false)
}

fn paragraph_to_markdown(p: &Paragraph) -> String {
    let plain = p.text();
    if plain.trim().is_empty() {
        return String::new();
    }

    // Process runs to preserve bold/italic/underline.
    let mut text_parts: Vec<String> = Vec::new();
    for pc in &p.content {
        if let ParagraphContent::Run(run) = pc {
            let mut text = run.text();
            if text.is_empty() {
                continue;
            }

            if let Some(prop) = &run.property {
                if bold_is_on(&prop.bold) {
                    text = format!("**{}**", text);
                }
                if italics_is_on(&prop.italics) {
                    text = format!("*{}*", text);
                }
                if underline_is_on(&prop.underline) {
                    text = format!("__{}__", text);
                }
            }

            text_parts.push(text);
        }
    }

    if text_parts.is_empty() {
        // Fall back to the aggregate paragraph text
        plain
    } else {
        text_parts.join("")
    }
}

fn convert_rtf_to_markdown(rtf_path: &Path) -> Result<String> {
    let bytes = fs::read(rtf_path)?;
    // Python tries multiple encodings; here we take a simpler step.
    let mut rtf_text = String::from_utf8_lossy(&bytes).to_string();

    // Remove simple RTF groups { ... } (no deep nesting)
    let re_group = Regex::new(r"\{[^{}]*\}")?;
    while rtf_text.contains('{') && rtf_text.contains('}') {
        let new = re_group.replace_all(&rtf_text, "");
        let new_owned = new.into_owned();
        if new_owned == rtf_text {
            break;
        }
        rtf_text = new_owned;
    }

    // Remove simple RTF commands \wordN?
    let re_cmd = Regex::new(r"\\[a-zA-Z]+\d*\s*")?;
    rtf_text = re_cmd.replace_all(&rtf_text, " ").into_owned();

    // Remove brace escapes
    let re_brace_cmd = Regex::new(r"\\[{}]")?;
    rtf_text = re_brace_cmd.replace_all(&rtf_text, "").into_owned();

    // Remove hex escapes \\'hh
    let re_hex = Regex::new(r"\\'[0-9a-fA-F]{2}")?;
    rtf_text = re_hex.replace_all(&rtf_text, "").into_owned();

    // Remove loose numbers from commands
    let re_nums = Regex::new(r"\s+\d+\s+")?;
    rtf_text = re_nums.replace_all(&rtf_text, " ").into_owned();

    let mut cleaned_lines: Vec<String> = Vec::new();

    for raw_line in rtf_text.lines() {
        let mut line = raw_line.trim().to_string();

        // Remove control characters
        line = line
            .chars()
            .filter(|c| !(*c as u32 <= 0x1F || (0x7F..=0x9F).contains(&(*c as u32))))
            .collect::<String>()
            .trim()
            .to_string();

        if line.is_empty() {
            cleaned_lines.push(String::new());
            continue;
        }

        let lower = line.to_lowercase();

        // Common fonts / artifacts
        let is_font_name = matches!(
            lower.as_str(),
            "times new roman"
                | "arial"
                | "calibri"
                | "helvetica"
                | "trebuchet ms"
                | "cambria"
                | "times"
        );

        let re_only_words = Regex::new(r"^[a-z\s]+\}?$").unwrap();
        let re_only_nums = Regex::new(r"^[\d\s\-]+$").unwrap();

        if is_font_name
            || re_only_words.is_match(&lower)
            || re_only_nums.is_match(&line)
            || line.chars().filter(|c| *c == '}').count()
                > line.chars().filter(|c| *c == ' ').count()
            || (line.len() < 3 && !line.chars().all(|c| c.is_alphanumeric()))
        {
            continue;
        }

        // Additional cleanup
        let re_spaces = Regex::new(r"\s+").unwrap();
        line = re_spaces.replace_all(&line, " ").into_owned();

        let re_stray_words_start = Regex::new(r"^[a-zA-Z]+\s+[a-zA-Z]+\s+").unwrap();
        line = re_stray_words_start.replace_all(&line, "").into_owned();

        let re_brace_end = Regex::new(r"\s*\}\s*$").unwrap();
        line = re_brace_end.replace_all(&line, "").into_owned();

        let re_brace_start = Regex::new(r"^\s*\{\s*").unwrap();
        line = re_brace_start.replace_all(&line, "").into_owned();

        let re_nums_start = Regex::new(r"^\s*[\d\-]+\s+").unwrap();
        line = re_nums_start.replace_all(&line, "").into_owned();

        let re_nums_end = Regex::new(r"\s+[\d\-]+\s*$").unwrap();
        line = re_nums_end.replace_all(&line, "").into_owned();

        if !line.trim().is_empty() {
            cleaned_lines.push(line.trim().to_string());
        }
    }

    // Markdown formatting heuristics
    let mut markdown_lines: Vec<String> = Vec::new();

    for line in cleaned_lines {
        if line.is_empty() {
            markdown_lines.push(String::new());
            continue;
        }

        let upper = line.to_uppercase();
        let lower = line.to_lowercase();

        // Main heading (all caps, no trailing period, contains keywords)
        if upper == line
            && line.len() > 15
            && line.len() < 120
            && !line.ends_with('.')
            && (upper.contains("TOMOGRAFIA")
                || upper.contains("ANGIO")
                || upper.contains("COMPUTADORIZADA"))
        {
            markdown_lines.push(format!("## {}", line));
            continue;
        }

        // Important sections
        let keywords = [
            "indicação clínica",
            "técnica do exame",
            "aspectos observados",
            "impressão",
        ];

        if keywords.iter().any(|k| lower.contains(k)) {
            if upper == line && line.len() > 10 {
                markdown_lines.push(format!("## {}", line));
                continue;
            } else {
                let start_keywords = ["indicação", "técnica", "aspectos", "impressão"];
                let mut handled = false;
                for k in &start_keywords {
                    if lower.starts_with(k) {
                        markdown_lines.push(format!("**{}**", line));
                        handled = true;
                        break;
                    }
                }
                if !handled {
                    markdown_lines.push(line);
                }
                continue;
            }
        }

        // Footnote-like notes → italic
        if lower.contains("probabilidade")
            || lower.contains("médico")
            || lower.contains("diagnóstica")
        {
            markdown_lines.push(format!("*{}*", line));
        } else {
            markdown_lines.push(line);
        }
    }

    // Remove duplicate empty lines
    let mut result: Vec<String> = Vec::new();
    let mut prev_empty = false;
    for line in markdown_lines {
        if line.trim().is_empty() {
            if !prev_empty {
                result.push(String::new());
                prev_empty = true;
            }
        } else {
            result.push(line);
            prev_empty = false;
        }
    }

    Ok(result.join("\n"))
}

fn main() -> Result<()> {
    let reports_dir = PathBuf::from("Templates_docx");
    if !reports_dir.exists() {
        eprintln!("Error: Folder {} not found!", reports_dir.display());
        return Ok(());
    }

    let markdown_dir = reports_dir
        .parent()
        .unwrap_or_else(|| Path::new("."))
        .join("Templates_markdown");
    fs::create_dir_all(&markdown_dir)?;

    // Process .docx
    let docx_files: Vec<PathBuf> = fs::read_dir(&reports_dir)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("docx"))
        .collect();

    println!("Found {} .docx files", docx_files.len());
    for docx_file in &docx_files {
        println!(
            "Converting {}...",
            docx_file.file_name().unwrap().to_string_lossy()
        );
        let markdown_content = convert_docx_to_markdown(docx_file)?;
        let output_file = markdown_dir.join(
            docx_file
                .file_stem()
                .unwrap()
                .to_string_lossy()
                .to_string()
                + ".md",
        );
        fs::write(&output_file, markdown_content)?;
        println!("  ✓ Saved to {}", output_file.file_name().unwrap().to_string_lossy());
    }

    // Process .rtf
    let rtf_files: Vec<PathBuf> = fs::read_dir(&reports_dir)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("rtf"))
        .collect();

    println!("\nFound {} .rtf files", rtf_files.len());
    for rtf_file in &rtf_files {
        println!(
            "Converting {}...",
            rtf_file.file_name().unwrap().to_string_lossy()
        );
        let markdown_content = convert_rtf_to_markdown(rtf_file)?;
        let output_file = markdown_dir.join(
            rtf_file
                .file_stem()
                .unwrap()
                .to_string_lossy()
                .to_string()
                + ".md",
        );
        fs::write(&output_file, markdown_content)?;
        println!("  ✓ Saved to {}", output_file.file_name().unwrap().to_string_lossy());
    }

    println!("\n✓ Conversion finished! Files saved to {}", markdown_dir.display());
    Ok(())
}
