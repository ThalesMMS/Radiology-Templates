use anyhow::{anyhow, Result};
use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::{Path, PathBuf};

const TARGETS: [(&str, &str); 3] = [
    ("Templates_docx", "docx"),
    ("Templates_markdown", "md"),
    ("Templates_txt", "txt"),
];

type IndexMap = HashMap<String, Vec<String>>;

fn load_index(root: &Path) -> Result<IndexMap> {
    let path = root.join("reports_index.json");
    if !path.exists() {
        return Err(anyhow!(
            "Index file not found at {}. Run generate_index first.",
            path.display()
        ));
    }
    let contents = fs::read_to_string(&path)?;
    let parsed: IndexMap = serde_json::from_str(&contents)?;
    Ok(parsed)
}

fn should_keep(path: &Path, expected: &HashSet<String>, root: &Path) -> bool {
    let rel = path
        .strip_prefix(root)
        .unwrap_or(path)
        .to_string_lossy()
        .to_string();
    expected.contains(&rel)
}

fn move_unindexed(root: &Path, index: &IndexMap) -> Result<usize> {
    let backup_dir = root.join("backup");
    fs::create_dir_all(&backup_dir)?;
    let mut moved = 0usize;

    for (folder, ext) in TARGETS {
        let dir = root.join(folder);
        if !dir.exists() {
            eprintln!("Skipping missing folder: {}", dir.display());
            continue;
        }

        let expected: HashSet<String> = index
            .get(folder)
            .cloned()
            .unwrap_or_default()
            .into_iter()
            .collect();

        for entry in fs::read_dir(&dir)? {
            let path = entry?.path();
            if !path.is_file() {
                continue;
            }
            let has_ext = path
                .extension()
                .and_then(|s| s.to_str())
                .map(|s| s.eq_ignore_ascii_case(ext))
                .unwrap_or(false);
            if !has_ext {
                continue;
            }
            if should_keep(&path, &expected, root) {
                continue;
            }

            let rel = path.strip_prefix(root).unwrap_or(&path);
            let dest = backup_dir.join(rel);
            if dest.exists() {
                eprintln!(
                    "Skip {}: destination already exists",
                    rel.to_string_lossy()
                );
                continue;
            }
            if let Some(parent) = dest.parent() {
                fs::create_dir_all(parent)?;
            }

            fs::rename(&path, &dest)?;
            moved += 1;
            println!(
                "Moved {} -> {}",
                rel.to_string_lossy(),
                dest.strip_prefix(root)
                    .unwrap_or(&dest)
                    .to_string_lossy()
            );
        }
    }

    Ok(moved)
}

fn main() -> Result<()> {
    let root = PathBuf::from(".");
    let index = load_index(&root)?;
    let moved = move_unindexed(&root, &index)?;
    println!("\nDone. Files moved: {}", moved);
    Ok(())
}
