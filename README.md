# Markdown to DOCX Converter

A lightweight, self-contained Markdown to DOCX converter with full LaTeX math support, syntax highlighting, and no system-wide dependencies.

## Features

✨ **Full LaTeX Math Support** - Equations, matrices, special symbols converted to Word's native equation editor  
🎨 **Syntax Highlighting** - Beautiful code blocks with multiple color schemes  
📊 **Tables & Lists** - Converted to native Word formatting  
🚀 **Fast & Lightweight** - Uses Pandoc locally, no system installation required  
📦 **Zero System Dependencies** - Everything contained in your project folder  
🔧 **Flexible** - CLI and Python API support  
🎯 **Batch Processing** - Convert multiple files at once

## Prerequisites

- **Python 3.10+**
- **uv** package manager ([Install uv](https://docs.astral.sh/uv/getting-started/installation/))

## Installation

### 1. Clone or Create Project

```
git clone 990aa/docConv
cd docConv
```

### 2. Install Dependencies

```
# uv will automatically create a virtual environment and install dependencies
uv sync
```

### 3. Setup Pandoc (One-Time)

```
# Download and install Pandoc locally to the project
uv run setup_pandoc.py
```

This installs Pandoc to the `bin/` directory in your project. You only need to run this once.

## Usage

### Quick Start

```
# Convert a single file
uv run main.py document.md

# Specify output filename
uv run main.py document.md output.docx

# Use different syntax highlighting style
uv run main.py document.md --highlight-style=pygments
```

### Convert All Markdown Files in a Folder

#### Option 1: Using Shell Commands

**Windows (PowerShell):**
```
Get-ChildItem *.md | ForEach-Object { uv run main.py $_.Name }
```

**Linux/macOS (Bash):**
```
for file in *.md; do
    uv run main.py "$file"
done
```

#### Option 2: Create a Batch Converter Script

Create `convert_all.py`:

```
"""Convert all Markdown files in current directory to DOCX"""
import glob
from pathlib import Path
from converter import MarkdownToDocxConverter

def convert_all_in_folder(folder='.', output_dir=None):
    """Convert all .md files in folder to DOCX"""
    converter = MarkdownToDocxConverter()
    
    # Find all markdown files
    md_files = glob.glob(f"{folder}/*.md")
    
    if not md_files:
        print(f"No markdown files found in {folder}")
        return
    
    print(f"Found {len(md_files)} markdown file(s)")
    
    # Convert each file
    for md_file in md_files:
        try:
            if output_dir:
                output_path = Path(output_dir) / Path(md_file).with_suffix('.docx').name
            else:
                output_path = None
            
            converter.convert(md_file, output_path)
        except Exception as e:
            print(f"❌ Failed to convert {md_file}: {e}")

if __name__ == '__main__':
    import sys
    folder = sys.argv if len(sys.argv) > 1 else '.'[10]
    output_dir = sys.argv if len(sys.argv) > 2 else None[11]
    convert_all_in_folder(folder, output_dir)
```

Run it:
```
# Convert all .md files in current directory
uv run convert_all.py

# Convert all .md files from specific folder
uv run convert_all.py ./documents

# Convert to specific output directory
uv run convert_all.py ./documents ./output
```

### Programmatic Usage

```
from converter import MarkdownToDocxConverter

# Initialize converter
converter = MarkdownToDocxConverter()

# Convert single file
converter.convert('document.md')

# Convert with custom output name
converter.convert('document.md', 'my-output.docx')

# Convert with custom highlighting
converter.convert('document.md', highlight_style='pygments')

# Batch convert multiple files
files = ['chapter1.md', 'chapter2.md', 'chapter3.md']
converter.batch_convert(files)

# Batch convert to specific directory
converter.batch_convert(files, output_dir='output')
```

## Syntax Highlighting Styles

Available styles for code blocks (use with `--highlight-style`):

- `tango` (default) - Colorful and readable
- `pygments` - Classic Python style
- `kate` - KDE editor style
- `monochrome` - Black and white
- `breezedark` - Dark theme
- `espresso` - Brown/coffee tones
- `zenburn` - Low contrast
- `haddock` - Haskell documentation style
- `textmate` - TextMate editor style

Example:
```
uv run main.py document.md --highlight-style=breezedark
```

## Supported Markdown Features

### Text Formatting
- **Bold**, *italic*, ***bold italic***
- `inline code`
- ~~strikethrough~~
- Headers (H1-H6)

### LaTeX Math

**Inline math:**
```
The equation $$ \frac{1-x}{2} = \frac{y-1}{3} $$ represents a line.
```

**Block equations:**
```
$$
\begin{bmatrix}
2 & -1 & 1 \\
\lambda & 2 & 0 \\
1 & -2 & 3
\end{bmatrix}
$$
```

**Special symbols:**
```
$$\mathbb{R}$$, $$\overrightarrow{a}$$, $$\alpha$$, $$\beta$$, etc.
```

### Code Blocks

```
```python
def hello_world():
    print("Hello, World!")
```

### Tables


| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |


### Lists

```
1. Ordered list item 1
2. Ordered list item 2

- Unordered list item
- Another item
  - Nested item
```

### Images

```

```

### Links

```
[Link text](https://example.com)
```

## Advanced Usage

### Custom Reference Document (Styling Template)

Create a DOCX file with your preferred styles and use it as a template:

```
uv run main.py document.md --reference-doc=template.docx
```

Or programmatically:
```
converter.convert('document.md', reference_doc='my_template.docx')
```

### Custom Pandoc Arguments

```
converter.convert(
    'document.md',
    'output.docx',
    number_sections=True,  # Number headings
    toc=True,              # Add table of contents
)
```

## Troubleshooting

### "Pandoc not found" Error

Run the setup script again:
```
uv run setup_pandoc.py
```

### Math Not Rendering Properly

Ensure you're using the correct LaTeX delimiters:
- Inline: `\( ... \)`
- Block: `\[ ... \]`

### Code Highlighting Not Working

Try a different highlight style:
```
uv run main.py document.md --highlight-style=tango
```

### File Not Found Error

Make sure you're in the correct directory and the file exists:
```
ls -la *.md  # Linux/macOS
dir *.md     # Windows
```

## Examples

### Example 1: Convert Single File

```
uv run main.py notes.md
# Output: notes.docx
```

### Example 2: Convert Multiple Files

```
uv run main.py chapter1.md
uv run main.py chapter2.md
uv run main.py chapter3.md
```

### Example 3: Batch Convert All Files

```
uv run convert_all.py
```

### Example 4: Convert with Custom Styling

```
uv run main.py report.md report.docx --highlight-style=pygments
```

## Sample Markdown File

See `example.md` in the project for a comprehensive example demonstrating:
- LaTeX math equations
- Code blocks with syntax highlighting
- Tables
- Lists
- Images
- Headers and formatting

## Development

### Run Tests

```
# Test with example file
uv run main.py example.md

# Verify conversion
# Open example.docx in Microsoft Word or LibreOffice
```

### Add New Features

Edit `converter.py` to modify the conversion logic or add new features.

## Support

For issues, questions, or suggestions:
1. Check the Troubleshooting section
2. Review `example.md` for syntax examples
3. Open an issue on GitHub

---