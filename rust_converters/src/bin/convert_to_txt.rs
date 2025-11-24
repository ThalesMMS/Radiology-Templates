use anyhow::Result;
use docx_rust::document::{BodyContent, Paragraph, ParagraphContent};
use docx_rust::DocxFile;
use std::env;
use std::fs;
use std::path::{Path, PathBuf};
use tempfile::TempDir;

fn clean_markdown_text(text: &str) -> String {
    text.replace('*', "").replace('#', "")
}

fn convert_md_file(md_path: &Path, output_dir: &Path) -> Result<()> {
    fs::create_dir_all(output_dir)?;
    let txt_path = output_dir.join(
        md_path
            .file_stem()
            .unwrap()
            .to_string_lossy()
            .to_string()
            + ".txt",
    );
    let content = fs::read_to_string(md_path)?;
    let cleaned = clean_markdown_text(&content);
    fs::write(txt_path, cleaned)?;
    Ok(())
}

fn convert_markdown_folder(md_dir: &Path, output_dir: &Path) -> Result<()> {
    let mut md_files: Vec<PathBuf> = fs::read_dir(md_dir)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("md"))
        .collect();

    if md_files.is_empty() {
        println!("No .md files found in {}", md_dir.display());
        return Ok(());
    }

    md_files.sort();
    for md_file in md_files {
        convert_md_file(&md_file, output_dir)?;
        println!(
            "✓ {} -> {}.txt",
            md_file.file_name().unwrap().to_string_lossy(),
            md_file
                .file_stem()
                .unwrap()
                .to_string_lossy()
        );
    }

    Ok(())
}

fn paragraph_to_markdown(p: &Paragraph) -> String {
    let plain = p.text();
    if plain.trim().is_empty() {
        return String::new();
    }

    let mut text_parts: Vec<String> = Vec::new();
    for pc in &p.content {
        if let ParagraphContent::Run(run) = pc {
            let mut text = run.text();
            if text.is_empty() {
                continue;
            }
            if let Some(prop) = &run.property {
                if prop.bold.is_some() {
                    text = format!("**{}**", text);
                }
                if prop.italics.is_some() {
                    text = format!("*{}*", text);
                }
                if prop.underline.is_some() {
                    text = format!("__{}__", text);
                }
            }
            text_parts.push(text);
        }
    }

    if text_parts.is_empty() {
        plain
    } else {
        text_parts.join("")
    }
}

fn convert_docx_to_markdown(docx_path: &Path) -> Result<String> {
    let file = DocxFile::from_file(docx_path)?;
    let docx = file.parse()?;

    let mut markdown_lines: Vec<String> = Vec::new();
    let body = &docx.document.body;

    for item in &body.content {
        if let BodyContent::Paragraph(p) = item {
            markdown_lines.push(paragraph_to_markdown(p));
        }
    }

    Ok(markdown_lines.join("\n"))
}

fn convert_from_docx(output_dir: &Path) -> Result<()> {
    let docx_dir = PathBuf::from("Templates_docx");
    let mut docx_files: Vec<PathBuf> = fs::read_dir(&docx_dir)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("docx"))
        .collect();

    if docx_files.is_empty() {
        println!("No .docx files found in {}", docx_dir.display());
        return Ok(());
    }

    docx_files.sort();

    let tmp_dir: TempDir = TempDir::new()?;
    let tmp_md_dir = tmp_dir.path();

    for docx_file in &docx_files {
        let markdown_content = convert_docx_to_markdown(docx_file)?;
        let md_output = tmp_md_dir.join(
            docx_file
                .file_stem()
                .unwrap()
                .to_string_lossy()
                .to_string()
                + ".md",
        );
        fs::write(&md_output, markdown_content)?;
        println!(
            "Generated temporary: {}",
            md_output.file_name().unwrap().to_string_lossy()
        );
    }

    convert_markdown_folder(tmp_md_dir, output_dir)?;
    // TempDir cleans up automatically when it goes out of scope
    Ok(())
}

fn main() -> Result<()> {
    let args: Vec<String> = env::args().skip(1).collect();
    let from_docx = args.iter().any(|a| a == "--from-docx");

    let md_dir = PathBuf::from("Templates_markdown");
    let txt_dir = PathBuf::from("Templates_txt");

    if from_docx {
        convert_from_docx(&txt_dir)?;
    } else {
        if !md_dir.exists() {
            anyhow::bail!("Source folder not found: {}", md_dir.display());
        }
        convert_markdown_folder(&md_dir, &txt_dir)?;
    }

    println!("\n✓ Files generated in {}", txt_dir.display());
    Ok(())
}
