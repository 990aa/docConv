"""
Command-line interface for MD to DOCX converter.
Usage: uv run main.py input.md [output.docx]
"""
import sys
from pathlib import Path
from converter import MarkdownToDocxConverter

def main():
    """CLI entry point"""
    if len(sys.argv) < 2:
        print("Usage: uv run main.py <input.md> [output.docx]")
        print("\nExamples:")
        print("  uv run main.py document.md")
        print("  uv run main.py document.md output.docx")
        print("  uv run main.py document.md --highlight-style=pygments")
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
