"""
Command-line interface for file to DOCX converter.
Supports Markdown and all programming languages with syntax highlighting.
Usage: uv run main.py input_file [output.docx]
"""
import sys
from pathlib import Path
from converter import MarkdownToDocxConverter

def main():
    """CLI entry point"""
    if len(sys.argv) < 2:
        print("Usage: uv run main.py <input_file> [output.docx] [options]")
        print("\nSupported file types:")
        print("  • Markdown: .md")
        print("  • Python: .py, .ipynb")
        print("  • JavaScript/TypeScript: .js, .jsx, .ts, .tsx")
        print("  • C/C++: .c, .cpp, .cc, .h, .hpp")
        print("  • Java: .java")
        print("  • Go: .go")
        print("  • And many more programming languages!")
        print("\nExamples:")
        print("  uv run main.py document.md")
        print("  uv run main.py script.py")
        print("  uv run main.py app.js output.docx")
        print("  uv run main.py main.java --highlight-style=kate")
        print("\nHighlight styles (for VS Code-like colors):")
        print("  tango (default), kate, breezedark, pygments, espresso")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    if not Path(input_file).exists():
        print(f"❌ Error: File '{input_file}' not found")
        sys.exit(1)
    
    # Parse arguments
    output_file = None
    kwargs = {}
    
    for arg in sys.argv[2:]:
        if arg.startswith('--'):
            key, value = arg[2:].split('=', 1)
            kwargs[key] = value
        else:
            output_file = arg
    
    # Convert
    converter = MarkdownToDocxConverter()
    converter.convert(input_file, output_file, **kwargs)

if __name__ == '__main__':
    main()
