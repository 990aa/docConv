# 📝 Document Converter (MD & Code → DOCX)# Markdown to DOCX Converter



Convert **Markdown** files and **source code** (Python, JavaScript, Java, C++, Go, and 50+ languages) to professional **DOCX** documents with:A lightweight, self-contained Markdown to DOCX converter with full LaTeX math support, syntax highlighting, and no system-wide dependencies.



- 🎨 **VS Code-like syntax highlighting** with proper colors## Features

- ✨ Full **LaTeX math** support (inline and block equations)

- 📊 Tables, lists, and rich formatting✨ **Full LaTeX Math Support** - Equations, matrices, special symbols converted to Word's native equation editor  

- 🔧 **Preserves indentation and formatting** perfectly🎨 **Syntax Highlighting** - Beautiful code blocks with multiple color schemes  

- 🌈 Multiple syntax highlighting themes📊 **Tables & Lists** - Converted to native Word formatting  

- 🚀 Powered by **Pandoc** and **pypandoc**🚀 **Fast & Lightweight** - Uses Pandoc locally, no system installation required  

📦 **Zero System Dependencies** - Everything contained in your project folder  

## ✨ Features🔧 **Flexible** - CLI and Python API support  

🎯 **Batch Processing** - Convert multiple files at once

### Supported File Types

## Prerequisites

**Markdown:**

- `.md` - Full markdown with math, tables, code blocks- **Python 3.10+**

- **uv** package manager ([Install uv](https://docs.astral.sh/uv/getting-started/installation/))

**Programming Languages:**

- **Python**: `.py`, `.ipynb` (Jupyter notebooks)## Installation

- **JavaScript/TypeScript**: `.js`, `.jsx`, `.ts`, `.tsx`

- **C/C++**: `.c`, `.cpp`, `.cc`, `.cxx`, `.h`, `.hpp`### 1. Clone or Create Project

- **Java**: `.java`

- **Go**: `.go````

- **Rust**: `.rs`git clone 990aa/docConv

- **Ruby**: `.rb`cd docConv

- **PHP**: `.php````

- **C#**: `.cs`

- **Swift**: `.swift`### 2. Install Dependencies

- **Kotlin**: `.kt`

- **Scala**: `.scala````

- **Shell**: `.sh`, `.bash`, `.zsh`, `.ps1`# uv will automatically create a virtual environment and install dependencies

- **SQL**: `.sql`uv sync

- **R**: `.r`, `.R````

- **Web**: `.html`, `.css`, `.scss`, `.xml`

- **Config**: `.json`, `.yaml`, `.yml`, `.toml`, `.ini`### 3. Setup Pandoc (One-Time)

- And many more!

```

### Syntax Highlighting# Download and install Pandoc locally to the project

uv run setup_pandoc.py

Choose from multiple color schemes for VS Code-like appearance:```

- `tango` (default) - Colorful and clear

- `kate` - Similar to VS Code Dark+This installs Pandoc to the `bin/` directory in your project. You only need to run this once.

- `breezedark` - Dark theme with good contrast

- `pygments` - Classic style## Usage

- `espresso` - Dark coffee theme

- `monochrome` - Black and white### Quick Start



## 🚀 Quick Start```

# Convert a single file

### Prerequisitesuv run main.py document.md



- **Python 3.10+**# Specify output filename

- **uv** package manager ([Install uv](https://docs.astral.sh/uv/getting-started/installation/))uv run main.py document.md output.docx



### Installation# Use different syntax highlighting style

uv run main.py document.md --highlight-style=pygments

1. **Clone the repository:**```

   ```bash

   git clone https://github.com/990aa/docConv### Convert All Markdown Files in a Folder

   cd docConv

   ```#### Option 1: Using Shell Commands



2. **Install dependencies:****Windows (PowerShell):**

   ```bash```

   uv syncGet-ChildItem *.md | ForEach-Object { uv run main.py $_.Name }

   ``````



3. **Install Pandoc locally** (one-time setup):**Linux/macOS (Bash):**

   ```bash```

   uv run setup_pandoc.pyfor file in *.md; do

   ```    uv run main.py "$file"

done

4. **Ready to convert!**```



## 📖 Usage#### Option 2: Create a Batch Converter Script



### Convert Markdown FilesCreate `convert_all.py`:



```bash```

# Basic conversion"""Convert all Markdown files in current directory to DOCX"""

uv run main.py example.mdimport glob

from pathlib import Path

# Custom output namefrom converter import MarkdownToDocxConverter

uv run main.py document.md output.docx

def convert_all_in_folder(folder='.', output_dir=None):

# Custom syntax highlighting style    """Convert all .md files in folder to DOCX"""

uv run main.py document.md --highlight-style=kate    converter = MarkdownToDocxConverter()

```    

    # Find all markdown files

### Convert Source Code Files    md_files = glob.glob(f"{folder}/*.md")

    

**Convert any programming language to DOCX with syntax highlighting:**    if not md_files:

        print(f"No markdown files found in {folder}")

```bash        return

# Python file    

uv run main.py script.py    print(f"Found {len(md_files)} markdown file(s)")

    

# JavaScript file    # Convert each file

uv run main.py app.js    for md_file in md_files:

        try:

# Java file            if output_dir:

uv run main.py Main.java                output_path = Path(output_dir) / Path(md_file).with_suffix('.docx').name

            else:

# C++ file                output_path = None

uv run main.py algorithm.cpp            

            converter.convert(md_file, output_path)

# TypeScript React component        except Exception as e:

uv run main.py Component.tsx            print(f"❌ Failed to convert {md_file}: {e}")



# Go programif __name__ == '__main__':

uv run main.py server.go    import sys

    folder = sys.argv if len(sys.argv) > 1 else '.'[10]

# Jupyter notebook    output_dir = sys.argv if len(sys.argv) > 2 else None[11]

uv run main.py analysis.ipynb    convert_all_in_folder(folder, output_dir)

```

# With custom output and style

uv run main.py script.py output.docx --highlight-style=breezedarkRun it:

``````

# Convert all .md files in current directory

### Batch Convert Multiple Filesuv run convert_all.py



**PowerShell (Windows):**# Convert all .md files from specific folder

```powershelluv run convert_all.py ./documents

# Convert all markdown files

Get-ChildItem *.md | ForEach-Object { uv run main.py $_.Name }# Convert to specific output directory

uv run convert_all.py ./documents ./output

# Convert all Python files```

Get-ChildItem *.py | ForEach-Object { uv run main.py $_.Name }

### Programmatic Usage

# Convert specific files

@('doc.md', 'script.py', 'app.js') | ForEach-Object { uv run main.py $_ }```

```from converter import MarkdownToDocxConverter



**Bash (Linux/macOS):**# Initialize converter

```bashconverter = MarkdownToDocxConverter()

# Convert all markdown files

for file in *.md; do uv run main.py "$file"; done# Convert single file

converter.convert('document.md')

# Convert all Python files

for file in *.py; do uv run main.py "$file"; done# Convert with custom output name

```converter.convert('document.md', 'my-output.docx')



### Python API Usage# Convert with custom highlighting

converter.convert('document.md', highlight_style='pygments')

```python

from converter import MarkdownToDocxConverter# Batch convert multiple files

files = ['chapter1.md', 'chapter2.md', 'chapter3.md']

converter = MarkdownToDocxConverter()converter.batch_convert(files)



# Convert markdown# Batch convert to specific directory

converter.convert('document.md')converter.batch_convert(files, output_dir='output')

```

# Convert Python code with VS Code-like colors

converter.convert('script.py', highlight_style='kate')## Syntax Highlighting Styles



# Convert JavaScript with custom outputAvailable styles for code blocks (use with `--highlight-style`):

converter.convert('app.js', output_file='code_documentation.docx')

- `tango` (default) - Colorful and readable

# Convert multiple files- `pygments` - Classic Python style

files = ['doc.md', 'script.py', 'app.js', 'Main.java']- `kate` - KDE editor style

converter.batch_convert(files, output_dir='output')- `monochrome` - Black and white

- `breezedark` - Dark theme

# Use custom reference document for styling- `espresso` - Brown/coffee tones

converter.convert('document.md', reference_doc='template.docx')- `zenburn` - Low contrast

```- `haddock` - Haskell documentation style

- `textmate` - TextMate editor style

## 🎨 Syntax Highlighting Styles

Example:

For VS Code-like appearance, recommended styles:```

uv run main.py document.md --highlight-style=breezedark

| Style | Description | Best For |```

|-------|-------------|----------|

| `tango` | Bright, colorful (default) | Light backgrounds |## Supported Markdown Features

| `kate` | VS Code Dark+ inspired | Dark theme lovers |

| `breezedark` | KDE dark theme | High contrast |### Text Formatting

| `pygments` | Classic Python highlighting | Documentation |- **Bold**, *italic*, ***bold italic***

| `espresso` | Coffee-dark theme | Dark mode |- `inline code`

- ~~strikethrough~~

Preview all styles:- Headers (H1-H6)

```bash

# Try different styles on demo files### LaTeX Math

uv run main.py demo.py --highlight-style=tango

uv run main.py demo.py --highlight-style=kate**Inline math:**

uv run main.py demo.js --highlight-style=breezedark```

```The equation $$ \frac{1-x}{2} = \frac{y-1}{3} $$ represents a line.

```

## 📁 Examples

**Block equations:**

### Markdown with Math```

$$

The `example.md` file demonstrates:\begin{bmatrix}

- LaTeX math equations (inline and block)2 & -1 & 1 \\

- Code blocks with syntax highlighting\lambda & 2 & 0 \\

- Tables and lists1 & -2 & 3

- Rich text formatting\end{bmatrix}

$$

### Source Code Files```



Try the demo files:**Special symbols:**

```bash```

# Convert Python demo$$\mathbb{R}$$, $$\overrightarrow{a}$$, $$\alpha$$, $$\beta$$, etc.

uv run main.py demo.py```



# Convert JavaScript demo  ### Code Blocks

uv run main.py demo.js

```

# Convert with different styles```python

uv run main.py demo.py --highlight-style=breezedarkdef hello_world():

```    print("Hello, World!")

```

## 🔧 How It Works

### Tables

1. **For Markdown files**: Converts directly with Pandoc, preserving all formatting and math

2. **For source code files**: 

   - Detects the programming language from file extension| Header 1 | Header 2 | Header 3 |

   - Wraps code in a markdown code block with language tag|----------|----------|----------|

   - Applies syntax highlighting using Pandoc's highlighter| Cell 1   | Cell 2   | Cell 3   |

   - Preserves all indentation and formatting| Cell 4   | Cell 5   | Cell 6   |

   - Generates DOCX with colored, formatted code



## 🎯 Key Features for Code Files### Lists



✅ **Perfect indentation preservation** - Spaces and tabs are kept exactly as in source  ```

✅ **Syntax highlighting** - Keywords, strings, comments, functions all colored  1. Ordered list item 1

✅ **Line wrapping** - Long lines handled gracefully  2. Ordered list item 2

✅ **File metadata** - Original filename shown as title  

✅ **Multiple languages** - 50+ programming languages supported  - Unordered list item

✅ **Jupyter notebooks** - Extracts and formats code cells from .ipynb files- Another item

  - Nested item

## 🛠️ Advanced Options```



### Custom Reference Document### Images



Create a `template.docx` with your preferred styles, fonts, and colors:```



```bash```

uv run main.py document.md --reference-doc=template.docx

```### Links



### Batch Processing```

[Link text](https://example.com)

```python```

from converter import MarkdownToDocxConverter

## Advanced Usage

converter = MarkdownToDocxConverter()

### Custom Reference Document (Styling Template)

# Convert multiple files at once

files = [Create a DOCX file with your preferred styles and use it as a template:

    'README.md',

    'script.py',```

    'app.js',uv run main.py document.md --reference-doc=template.docx

    'Main.java',```

    'analysis.ipynb'

]Or programmatically:

```

converter.batch_convert(files, output_dir='converted_docs', highlight_style='kate')converter.convert('document.md', reference_doc='my_template.docx')

``````



### Supported Pandoc Highlight Styles### Custom Pandoc Arguments



Full list: `tango`, `pygments`, `kate`, `monochrome`, `breezedark`, `espresso`, `zenburn`, `haddock`, `textmate````

converter.convert(

## 🐛 Troubleshooting    'document.md',

    'output.docx',

**Issue: "Pandoc not found"**    number_sections=True,  # Number headings

```bash    toc=True,              # Add table of contents

# Run setup again)

uv run setup_pandoc.py```

```

## Troubleshooting

**Issue: Colors not showing in Word**

- Try different highlight styles (`--highlight-style=kate`)### "Pandoc not found" Error

- Ensure you're using Microsoft Word (some viewers may not display colors)

Run the setup script again:

**Issue: Indentation looks wrong**```

- The converter preserves original indentation perfectlyuv run setup_pandoc.py

- Check that your source file uses consistent spacing```

- DOCX will respect all tabs and spaces from the original

### Math Not Rendering Properly

**Issue: Jupyter notebook not converting**

- Ensure the `.ipynb` file is valid JSONEnsure you're using the correct LaTeX delimiters:

- The converter extracts only code cells, not markdown cells- Inline: `\( ... \)`

- Block: `\[ ... \]`

## 📚 Supported Markdown Features (for .md files)

### Code Highlighting Not Working

### Text Formatting

- **Bold**, *italic*, ***bold italic***Try a different highlight style:

- `inline code````

- ~~strikethrough~~uv run main.py document.md --highlight-style=tango

- Headers (H1-H6)```



### LaTeX Math### File Not Found Error



**Inline math:**Make sure you're in the correct directory and the file exists:

```markdown```

The equation \( \frac{1-x}{2} = \frac{y-1}{3} \) represents a line.ls -la *.md  # Linux/macOS

```dir *.md     # Windows

```

**Block equations:**

```markdown## Examples

\[

\begin{bmatrix}### Example 1: Convert Single File

2 & -1 & 1 \\

\lambda & 2 & 0 \\```

1 & -2 & 3uv run main.py notes.md

\end{bmatrix}# Output: notes.docx

\]```

```

### Example 2: Convert Multiple Files

### Code Blocks

```

```markdownuv run main.py chapter1.md

```pythonuv run main.py chapter2.md

def hello_world():uv run main.py chapter3.md

    print("Hello, World!")```

```

```### Example 3: Batch Convert All Files



### Tables```

uv run convert_all.py

```markdown```

| Header 1 | Header 2 |

|----------|----------|### Example 4: Convert with Custom Styling

| Cell 1   | Cell 2   |

``````

uv run main.py report.md report.docx --highlight-style=pygments

### Lists```



```markdown## Sample Markdown File

1. Ordered item

2. Another itemSee `example.md` in the project for a comprehensive example demonstrating:

- LaTeX math equations

- Unordered item- Code blocks with syntax highlighting

- Another item- Tables

```- Lists

- Images

## 💡 Use Cases- Headers and formatting



- **Documentation**: Convert README.md files to professional Word documents## Development

- **Code Portfolios**: Showcase your code with proper syntax highlighting

- **Reports**: Include code snippets in technical reports with proper formatting### Run Tests

- **Education**: Share programming examples in DOCX format for students

- **Archival**: Convert source code to readable, formatted documents```

- **Technical Writing**: Mix explanations (markdown) and code in one document# Test with example file

uv run main.py example.md

## 📄 License

# Verify conversion

MIT License - See LICENSE.md for details# Open example.docx in Microsoft Word or LibreOffice

```

## 🤝 Contributing

### Add New Features

Contributions are welcome! Feel free to:

- Report bugsEdit `converter.py` to modify the conversion logic or add new features.

- Suggest new features

- Submit pull requests## Support

- Add support for more languages

For issues, questions, or suggestions:

---1. Check the Troubleshooting section

2. Review `example.md` for syntax examples

**Made with ❤️ by 990aa**3. Open an issue on GitHub


---