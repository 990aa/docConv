"""
Markdown to DOCX converter using Pandoc.
Supports LaTeX math, code highlighting, tables, and more.
Also converts source code files with VS Code-like syntax highlighting.
"""
import os
import sys
import pypandoc
from pathlib import Path
from typing import Optional
import tempfile

class MarkdownToDocxConverter:
    """Convert Markdown files to DOCX with full LaTeX math support"""
    
    # Mapping of file extensions to Pandoc language identifiers
    LANGUAGE_MAP = {
        '.py': 'python',
        '.ipynb': 'python',  # Jupyter notebooks (will extract code)
        '.js': 'javascript',
        '.jsx': 'javascript',
        '.ts': 'typescript',
        '.tsx': 'typescript',
        '.c': 'c',
        '.cpp': 'cpp',
        '.cc': 'cpp',
        '.cxx': 'cpp',
        '.h': 'c',
        '.hpp': 'cpp',
        '.java': 'java',
        '.go': 'go',
        '.rs': 'rust',
        '.rb': 'ruby',
        '.php': 'php',
        '.swift': 'swift',
        '.kt': 'kotlin',
        '.scala': 'scala',
        '.sh': 'bash',
        '.bash': 'bash',
        '.zsh': 'bash',
        '.ps1': 'powershell',
        '.sql': 'sql',
        '.r': 'r',
        '.R': 'r',
        '.m': 'matlab',
        '.lua': 'lua',
        '.pl': 'perl',
        '.html': 'html',
        '.xml': 'xml',
        '.css': 'css',
        '.scss': 'scss',
        '.sass': 'sass',
        '.json': 'json',
        '.yaml': 'yaml',
        '.yml': 'yaml',
        '.toml': 'toml',
        '.ini': 'ini',
        '.cfg': 'ini',
        '.md': 'markdown',
        '.tex': 'latex',
        '.vim': 'vim',
        '.dart': 'dart',
        '.ex': 'elixir',
        '.exs': 'elixir',
        '.erl': 'erlang',
        '.hrl': 'erlang',
        '.clj': 'clojure',
        '.fs': 'fsharp',
        '.cs': 'csharp',
        '.vb': 'vb',
    }
    
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
    
    def _detect_language(self, file_path: str) -> Optional[str]:
        """Detect programming language from file extension."""
        ext = Path(file_path).suffix.lower()
        return self.LANGUAGE_MAP.get(ext)
    
    def _read_jupyter_notebook(self, notebook_path: str) -> str:
        """Extract code cells from Jupyter notebook."""
        import json
        
        try:
            with open(notebook_path, 'r', encoding='utf-8') as f:
                notebook = json.load(f)
            
            code_blocks = []
            for cell in notebook.get('cells', []):
                if cell.get('cell_type') == 'code':
                    source = cell.get('source', [])
                    if isinstance(source, list):
                        source = ''.join(source)
                    if source.strip():
                        code_blocks.append(source)
            
            return '\n\n# ' + '='*50 + '\n\n'.join(code_blocks)
        except Exception as e:
            print(f"⚠️  Error reading notebook: {e}")
            # Fallback to treating as text
            with open(notebook_path, 'r', encoding='utf-8') as f:
                return f.read()
    
    def _create_markdown_from_code(self, file_path: str) -> str:
        """Create a markdown file with syntax-highlighted code block."""
        language = self._detect_language(file_path)
        file_name = Path(file_path).name
        
        # Read the source code
        if file_path.endswith('.ipynb'):
            code = self._read_jupyter_notebook(file_path)
            language = 'python'
        else:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                code = f.read()
        
        # Create markdown with code block
        markdown_content = f"""---
title: "{file_name}"
---

# {file_name}

```{language if language else ''}
{code}
```
"""
        return markdown_content
    
    def convert(
        self, 
        input_file: str, 
        output_file: Optional[str] = None,
        reference_doc: Optional[str] = None,
        highlight_style: str = 'tango',
        **kwargs
    ) -> str:
        """
        Convert Markdown or source code file to DOCX with syntax highlighting.
        
        Args:
            input_file: Path to input file (.md, .py, .js, .java, etc.)
            output_file: Path to output .docx file (auto-generated if None)
            reference_doc: Path to reference DOCX for custom styling
            highlight_style: Code syntax highlighting style 
                           (tango, pygments, kate, monochrome, breezedark, espresso, etc.)
                           For VS Code-like colors, use: 'tango', 'kate', or 'breezedark'
            **kwargs: Additional pandoc arguments
        
        Returns:
            Path to generated DOCX file
        """
        input_path = Path(input_file)
        
        # Check if input is a source code file (not markdown)
        is_code_file = input_path.suffix.lower() != '.md' and self._detect_language(input_file)
        temp_md_file = None
        
        if is_code_file:
            # Create temporary markdown file with code block
            print(f"🔍 Detected: {self._detect_language(input_file)} source file")
            markdown_content = self._create_markdown_from_code(input_file)
            
            # Write to temporary file
            temp_md = tempfile.NamedTemporaryFile(mode='w', suffix='.md', 
                                                  delete=False, encoding='utf-8')
            temp_md.write(markdown_content)
            temp_md.close()
            temp_md_file = temp_md.name
            input_file = temp_md_file
        
        # Auto-generate output filename if not provided
        if output_file is None:
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
        
        print(f"📄 Converting {input_path.name} → {output_file}")
        if not is_code_file:
            print("   Math support: ENABLED (\\(...\\) and \\[...\\])")
        print(f"   Syntax highlighting: {highlight_style}")
        
        try:
            pypandoc.convert_file(
                input_file,
                'docx',
                outputfile=output_file,
                extra_args=extra_args
            )
            print("✅ Conversion successful!")
            
            # Clean up temporary file if created
            if temp_md_file and os.path.exists(temp_md_file):
                os.unlink(temp_md_file)
            
            return output_file
            
        except Exception as e:
            print(f"❌ Conversion failed: {e}")
            # Clean up temporary file if created
            if temp_md_file and os.path.exists(temp_md_file):
                os.unlink(temp_md_file)
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
    
    # Single file conversion - Markdown
    converter.convert('example.md')
    
    # Convert Python file with syntax highlighting
    # converter.convert('script.py', highlight_style='kate')
    
    # Convert any source code file
    # converter.convert('app.js')
    # converter.convert('main.java')
    # converter.convert('server.go')
    
    # With custom styling
    # converter.convert('example.md', reference_doc='template.docx')
    
    # Batch conversion
    # files = ['doc1.md', 'script.py', 'app.js']
    # converter.batch_convert(files, output_dir='output')


if __name__ == '__main__':
    main()
