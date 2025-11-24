#!/usr/bin/env python3
"""
Script to convert .docx and .rtf files to Markdown while preserving formatting.
"""

import os
import sys
from pathlib import Path
import re

try:
    from docx import Document
except ImportError:
    print("Installing python-docx...")
    os.system(f"{sys.executable} -m pip install python-docx")
    from docx import Document

try:
    from striprtf.striprtf import rtf_to_text
except ImportError:
    print("Installing striprtf...")
    os.system(f"{sys.executable} -m pip install striprtf")
    from striprtf.striprtf import rtf_to_text


def convert_docx_to_markdown(docx_path):
    """Convert a .docx file to Markdown preserving formatting."""
    try:
        doc = Document(docx_path)
        markdown_lines = []
        
        for paragraph in doc.paragraphs:
            if not paragraph.text.strip():
                markdown_lines.append("")
                continue
            
            # Process runs to preserve formatting
            text_parts = []
            for run in paragraph.runs:
                text = run.text
                if not text:
                    continue

                # Apply formatting using pure Markdown (no HTML)
                if run.bold:
                    text = f"**{text}**"
                if run.italic:
                    text = f"*{text}*"
                if run.underline:
                    text = f"__{text}__"

                text_parts.append(text)
            
            # If there are no formatted runs, fall back to the paragraph text
            if not text_parts:
                para_text = paragraph.text
            else:
                para_text = "".join(text_parts)
            
            # Inspect paragraph style
            style = paragraph.style.name.lower()
            
            # Headings
            if 'heading' in style or 'title' in style:
                level = 1
                if 'heading 1' in style or 'title' in style:
                    level = 1
                elif 'heading 2' in style:
                    level = 2
                elif 'heading 3' in style:
                    level = 3
                elif 'heading 4' in style:
                    level = 4
                elif 'heading 5' in style:
                    level = 5
                elif 'heading 6' in style:
                    level = 6
                
                markdown_lines.append(f"{'#' * level} {para_text}")
            else:
                markdown_lines.append(para_text)
        
        # Process tables
        for table in doc.tables:
            markdown_lines.append("")
            # Header
            header_row = table.rows[0]
            header_cells = [cell.text.strip() for cell in header_row.cells]
            markdown_lines.append("| " + " | ".join(header_cells) + " |")
            markdown_lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |")
            
            # Data rows
            for row in table.rows[1:]:
                cells = [cell.text.strip() for cell in row.cells]
                markdown_lines.append("| " + " | ".join(cells) + " |")
            markdown_lines.append("")
        
        return "\n".join(markdown_lines)
    
    except Exception as e:
        return f"# Error converting {docx_path}\n\nError: {str(e)}"


def convert_rtf_to_markdown(rtf_path):
    """Convert an .rtf file to Markdown preserving basic formatting."""
    try:
        with open(rtf_path, 'rb') as f:
            rtf_content = f.read()
        
        # Try different encodings
        encodings = ['latin-1', 'cp1252', 'iso-8859-1', 'utf-8']
        rtf_text = None
        
        for encoding in encodings:
            try:
                candidate = rtf_content.decode(encoding, errors='ignore')
                parsed = rtf_to_text(candidate)
                if parsed.strip():
                    rtf_text = candidate
                    break
            except Exception:
                continue
        
        if rtf_text is None:
            rtf_text = rtf_content.decode('latin-1', errors='ignore')
        
        # Use striprtf to extract clean text
        try:
            plain_text = rtf_to_text(rtf_text)
        except Exception as e:
            # Fallback: basic manual extraction removing RTF commands
            plain_text = rtf_text
            # Remove empty RTF groups first
            while '{' in plain_text and '}' in plain_text:
                plain_text = re.sub(r'\{[^{}]*\}', '', plain_text)
            # Remove RTF commands
            plain_text = re.sub(r'\\[a-z]+\d*\s*', ' ', plain_text)
            plain_text = re.sub(r'\\[{}]', '', plain_text)
            # Remove RTF special characters
            plain_text = re.sub(r'\\\'[0-9a-f]{2}', '', plain_text)
            # Remove stray numbers that belong to RTF commands
            plain_text = re.sub(r'\s+\d+\s+', ' ', plain_text)
        
        # Clean extracted text
        # Remove lines that are only font names or commands
        lines = plain_text.split('\n')
        cleaned_lines = []
        
        for line in lines:
            line = line.strip()
            
            # Strip control characters
            line = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', line)
            line = line.strip()
            
            # Drop lines that are only artifacts
            if not line:
                cleaned_lines.append("")
                continue
            
            # Remove lines that are only font names or RTF commands
            if (line.lower() in ['times new roman', 'arial', 'calibri', 'helvetica', 
                                 'trebuchet ms', 'cambria', 'times'] or
                re.match(r'^[a-z\s]+\}?$', line.lower()) or
                re.match(r'^[\d\s\-]+$', line) or
                line.count('}') > line.count(' ') or
                (len(line) < 3 and not line.isalnum())):
                continue
            
            # Remove RTF conversion artifacts
            line = re.sub(r'\s+', ' ', line)  # Multiple spaces
            line = re.sub(r'^[a-z]+\s+[a-z]+\s+', '', line)  # Remove stray words at start
            line = re.sub(r'\s*\}\s*$', '', line)  # Remove braces at end
            line = re.sub(r'^\s*\{\s*', '', line)  # Remove braces at start
            
            # Remove stray numbers at start/end
            line = re.sub(r'^\s*[\d\-]+\s+', '', line)
            line = re.sub(r'\s+[\d\-]+\s*$', '', line)
            
            if line.strip():
                cleaned_lines.append(line.strip())
        
        # Process cleaned lines and apply formatting
        markdown_lines = []
        
        for line in cleaned_lines:
            if not line:
                markdown_lines.append("")
                continue
            
            # Detect primary headings (uppercase, no trailing period, contains keywords)
            if (line.isupper() and 15 < len(line) < 120 and 
                not line.endswith('.') and 
                ('TOMOGRAFIA' in line or 'ANGIO' in line or 'COMPUTADORIZADA' in line)):
                markdown_lines.append(f"## {line}")
            # Detect important sections
            elif any(keyword in line.upper() for keyword in 
                    ['INDICAÇÃO CLÍNICA', 'TÉCNICA DO EXAME', 'ASPECTOS OBSERVADOS', 'IMPRESSÃO']):
                # Uppercase text is a heading
                if line.isupper() and len(line) > 10:
                    markdown_lines.append(f"## {line}")
                else:
                    # If it starts with a keyword, make it bold
                    for keyword in ['INDICAÇÃO', 'TÉCNICA', 'ASPECTOS', 'IMPRESSÃO']:
                        if line.upper().startswith(keyword):
                            markdown_lines.append(f"**{line}**")
                            break
                    else:
                        markdown_lines.append(line)
            # Detect italic text (usually footnotes)
            elif ('probabilidade' in line.lower() or 
                  'médico' in line.lower() or 
                  'diagnóstica' in line.lower()):
                markdown_lines.append(f"*{line}*")
            else:
                markdown_lines.append(line)
        
        # Remove excessive empty lines
        result = []
        prev_empty = False
        for line in markdown_lines:
            if not line.strip():
                if not prev_empty:
                    result.append("")
                prev_empty = True
            else:
                result.append(line)
                prev_empty = False
        
        return "\n".join(result)
    
    except Exception as e:
        import traceback
        return f"# Error converting {rtf_path}\n\nError: {str(e)}\n\nTraceback: {traceback.format_exc()}"


def main():
    """Main function that processes all files inside the Reports folder."""
    reports_dir = Path(__file__).parent / "Templates_docx"
    
    if not reports_dir.exists():
        print(f"Error: Folder {reports_dir} not found!")
        return
    
    # Create folder for Markdown output
    markdown_dir = reports_dir.parent / "Templates_markdown"
    markdown_dir.mkdir(exist_ok=True)
    
    # Process .docx files
    docx_files = list(reports_dir.glob("*.docx"))
    print(f"Found {len(docx_files)} .docx files")
    
    for docx_file in docx_files:
        print(f"Converting {docx_file.name}...")
        markdown_content = convert_docx_to_markdown(docx_file)
        output_file = markdown_dir / f"{docx_file.stem}.md"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        print(f"  ✓ Saved to {output_file.name}")
    
    # Process .rtf files
    rtf_files = list(reports_dir.glob("*.rtf"))
    print(f"\nFound {len(rtf_files)} .rtf files")
    
    for rtf_file in rtf_files:
        print(f"Converting {rtf_file.name}...")
        markdown_content = convert_rtf_to_markdown(rtf_file)
        output_file = markdown_dir / f"{rtf_file.stem}.md"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        print(f"  ✓ Saved to {output_file.name}")
    
    print(f"\n✓ Conversion finished! Files saved to {markdown_dir}")


if __name__ == "__main__":
    main()
