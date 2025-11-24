use anyhow::Result;
use std::env;
use std::fs;
use std::path::{Path, PathBuf};

const SECTION_PREFIXES: &[&str] = &[
    "técnica do exame:",
    "aspectos observados:",
    "impressão:",
    "informe clínico:",
    "indicação clínica:",
    "indicação:",
];

fn find_first_last_nonempty(lines: &[String]) -> Option<(usize, usize)> {
    let nonempty_indices: Vec<usize> = lines
        .iter()
        .enumerate()
        .filter_map(|(i, line)| if line.trim().is_empty() { None } else { Some(i) })
        .collect();

    if nonempty_indices.is_empty() {
        None
    } else {
        Some((nonempty_indices[0], *nonempty_indices.last().unwrap()))
    }
}

fn should_bold_section(line: &str) -> bool {
    let lowered = line.to_lowercase();
    let trimmed = lowered.trim();

    if SECTION_PREFIXES
        .iter()
        .any(|prefix| trimmed.starts_with(prefix))
    {
        return true;
    }

    if trimmed.ends_with(':') && trimmed.len() <= 120 {
        return true;
    }

    false
}

fn format_lines_as_markdown(lines: &[String]) -> Vec<String> {
    let first_last = find_first_last_nonempty(lines);
    let (first_idx, last_idx) = first_last.unwrap_or((usize::MAX, usize::MAX));

    let mut output: Vec<String> = Vec::new();

    for (idx, line) in lines.iter().enumerate() {
        let stripped = line.trim();
        if stripped.is_empty() {
            output.push(String::new());
            continue;
        }

        let is_first = idx == first_idx;
        let is_last = idx == last_idx;

        let mut text = stripped.to_string();
        if is_last {
            text = format!("*{}*", text);
        } else if is_first || should_bold_section(stripped) {
            text = format!("**{}**", text);
        }

        output.push(text);
    }

    output
}

fn convert_txt_file(txt_path: &Path, output_dir: &Path) -> Result<()> {
    let content = fs::read_to_string(txt_path)?;
    let lines: Vec<String> = content.lines().map(|s| s.to_string()).collect();
    let formatted = format_lines_as_markdown(&lines);

    fs::create_dir_all(output_dir)?;
    let md_path = output_dir.join(
        txt_path
            .file_stem()
            .unwrap()
            .to_string_lossy()
            .to_string()
            + ".md",
    );
    fs::write(md_path.clone(), formatted.join("\n"))?;
    println!(
        "✓ {} -> {}",
        txt_path.file_name().unwrap().to_string_lossy(),
        md_path.file_name().unwrap().to_string_lossy()
    );
    Ok(())
}

fn convert_folder(txt_dir: &Path, output_dir: &Path) -> Result<()> {
    let mut txt_files: Vec<PathBuf> = fs::read_dir(txt_dir)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("txt"))
        .collect();

    if txt_files.is_empty() {
        println!("No .txt files found in {}", txt_dir.display());
        return Ok(());
    }

    txt_files.sort();

    for txt_file in txt_files {
        convert_txt_file(&txt_file, output_dir)?;
    }

    Ok(())
}

fn main() -> Result<()> {
    let mut args = env::args().skip(1).peekable();

    let mut txt_dir_arg: Option<PathBuf> = None;
    let mut output_dir_arg: Option<PathBuf> = None;

    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--txt-dir" => {
                if let Some(p) = args.next() {
                    txt_dir_arg = Some(PathBuf::from(p));
                } else {
                    anyhow::bail!("--txt-dir requires a path");
                }
            }
            "--output-dir" => {
                if let Some(p) = args.next() {
                    output_dir_arg = Some(PathBuf::from(p));
                } else {
                    anyhow::bail!("--output-dir requires a path");
                }
            }
            other => {
                eprintln!("Unknown argument ignored: {}", other);
            }
        }
    }

    let txt_dir = txt_dir_arg.unwrap_or_else(|| PathBuf::from("Templates_txt"));
    let output_dir = output_dir_arg.unwrap_or_else(|| PathBuf::from("Templates_markdown"));

    if !txt_dir.exists() {
        anyhow::bail!("Source folder not found: {}", txt_dir.display());
    }

    convert_folder(&txt_dir, &output_dir)?;
    println!("\n✓ Markdown generated in {}", output_dir.display());
    Ok(())
}
