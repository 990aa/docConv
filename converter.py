"""
Markdown to DOCX converter using Pandoc.
Supports LaTeX math, code highlighting, tables, and more.
"""
import os
import sys
import pypandoc
from pathlib import Path
from typing import Optional, Dict, Any

class MarkdownToDocxConverter:
    """Convert Markdown files to DOCX with full LaTeX math support"""
    
    def __init__(self, pandoc_path: Optional[str] = None):
        """
        Initialize converter with optional custom Pandoc path.
        
        Args:
            pandoc_path: Path to pandoc executable. If None, uses local bin/pandoc
        """
        if pandoc_path is None:
            # Use local pandoc installation
            bin_dir = Path(__file__).parent / 'bin'
            pandoc_exe = 'pandoc.exe' if sys.platform == 'win32' else 'pandoc'
            pandoc_path = str(bin_dir / pandoc_exe)
        
        if os.path.exists(pandoc_path):
            os.environ['PYPANDOC_PANDOC'] = pandoc_path
            print(f"✓ Using Pandoc at: {pandoc_path}")
        else:
            print(f"⚠️  Pandoc not found at {pandoc_path}")
            print("   Run 'uv run setup_pandoc.py' first to install Pandoc locally")
    
    def convert(
        self, 
        input_file: str, 
        output_file: Optional[str] = None,
        reference_doc: Optional[str] = None,
        highlight_style: str = 'tango',
        **kwargs
    ) -> str:
        """
        Convert Markdown file to DOCX.
        
        Args:
            input_file: Path to input .md file
            output_file: Path to output .docx file (auto-generated if None)
            reference_doc: Path to reference DOCX for custom styling
            highlight_style: Code syntax highlighting style 
                           (tango, pygments, kate, monochrome, breezedark, etc.)
            **kwargs: Additional pandoc arguments
        
        Returns:
            Path to generated DOCX file
        """
        # Auto-generate output filename if not provided
        if output_file is None:
            input_path = Path(input_file)
            output_file = str(input_path.with_suffix('.docx'))
        
        # Build pandoc arguments with proper math support
        extra_args = [
            # Enable LaTeX math recognition with \( \) and \[ \] delimiters
            '--from=markdown+tex_math_single_backslash',
            f'--highlight-style={highlight_style}',
        ]
        
        # Add reference document for custom styling if provided
        if reference_doc and os.path.exists(reference_doc):
            extra_args.append(f'--reference-doc={reference_doc}')
        
        # Add any additional arguments
        for key, value in kwargs.items():
            if isinstance(value, bool):
                if value:
                    extra_args.append(f'--{key}')
            else:
                extra_args.append(f'--{key}={value}')
        
        print(f"📄 Converting {input_file} → {output_file}")
        print(f"   Math support: ENABLED (\\(...\\) and \\[...\\])")
        print(f"   Highlight style: {highlight_style}")
        
        try:
            pypandoc.convert_file(
                input_file,
                'docx',
                outputfile=output_file,
                extra_args=extra_args
            )
            print(f"✅ Conversion successful!")
            return output_file
            
        except Exception as e:
            print(f"❌ Conversion failed: {e}")
            raise
    
    def batch_convert(
        self, 
        input_files: list[str],
        output_dir: Optional[str] = None,
        **kwargs
    ) -> list[str]:
        """
        Convert multiple Markdown files to DOCX.
        
        Args:
            input_files: List of input .md file paths
            output_dir: Directory for output files (same as input if None)
            **kwargs: Additional arguments passed to convert()
        
        Returns:
            List of generated DOCX file paths
        """
        output_files = []
        
        for input_file in input_files:
            input_path = Path(input_file)
            
            if output_dir:
                output_path = Path(output_dir) / input_path.with_suffix('.docx').name
                os.makedirs(output_dir, exist_ok=True)
            else:
                output_path = input_path.with_suffix('.docx')
            
            try:
                output = self.convert(input_file, str(output_path), **kwargs)
                output_files.append(output)
            except Exception as e:
                print(f"⚠️  Skipped {input_file}: {e}")
                
        return output_files


def main():
    """Example usage"""
    converter = MarkdownToDocxConverter()
    
    # Single file conversion
    converter.convert('example.md')
    
    # With custom styling
    # converter.convert('example.md', reference_doc='template.docx')
    
    # Batch conversion
    # md_files = ['doc1.md', 'doc2.md', 'doc3.md']
    # converter.batch_convert(md_files, output_dir='output')


if __name__ == '__main__':
    main()
