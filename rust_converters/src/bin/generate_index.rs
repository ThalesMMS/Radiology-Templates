use anyhow::Result;
use serde::Serialize;
use serde_json::ser::{PrettyFormatter, Serializer};
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

const TARGETS: [(&str, &str); 3] = [
    ("Templates_docx", "docx"),
    ("Templates_markdown", "md"),
    ("Templates_txt", "txt"),
];

fn collect_files(root: &Path) -> Result<BTreeMap<String, Vec<String>>> {
    let mut index = BTreeMap::new();

    for (folder, ext) in TARGETS {
        let dir = root.join(folder);
        if !dir.exists() {
            eprintln!("Skipping missing folder: {}", dir.display());
            index.insert(folder.to_string(), Vec::new());
            continue;
        }

        let mut files: Vec<String> = fs::read_dir(&dir)?
            .filter_map(|e| e.ok())
            .map(|e| e.path())
            .filter(|p| {
                p.is_file()
                    && p.extension()
                        .and_then(|s| s.to_str())
                        .map(|s| s.eq_ignore_ascii_case(ext))
                        .unwrap_or(false)
            })
            .map(|p| {
                p.strip_prefix(root)
                    .unwrap_or(&p)
                    .to_string_lossy()
                    .to_string()
            })
            .collect();

        files.sort();
        index.insert(folder.to_string(), files);
    }

    Ok(index)
}

fn write_json_pretty<T: Serialize>(value: &T, path: &Path) -> Result<()> {
    let mut buffer = Vec::new();
    let formatter = PrettyFormatter::with_indent(b"  ");
    let mut serializer = Serializer::with_formatter(&mut buffer, formatter);
    value.serialize(&mut serializer)?;
    fs::write(path, buffer)?;
    Ok(())
}

fn main() -> Result<()> {
    let root = PathBuf::from(".");
    let index = collect_files(&root)?;
    let output = root.join("reports_index.json");
    write_json_pretty(&index, &output)?;
    println!("\nIndex written to {}", output.display());
    Ok(())
}
