"""
Setup script to download Pandoc locally to the project.
Run once: uv run setup_pandoc.py
"""
import os
import sys
from pypandoc.pandoc_download import download_pandoc

def setup_local_pandoc():
    """Download pandoc to local bin directory"""
    bin_dir = os.path.join(os.path.dirname(__file__), 'bin')
    os.makedirs(bin_dir, exist_ok=True)
    
    print(f"📦 Installing Pandoc to {bin_dir}...")
    try:
        download_pandoc(targetfolder=bin_dir)
        print("✅ Pandoc installed successfully!")
        
        # Verify installation
        pandoc_exe = 'pandoc.exe' if sys.platform == 'win32' else 'pandoc'
        pandoc_path = os.path.join(bin_dir, pandoc_exe)
        
        if os.path.exists(pandoc_path):
            print(f"📍 Pandoc location: {pandoc_path}")
            return True
        else:
            print("⚠️  Warning: Pandoc binary not found at expected location")
            return False
            
    except Exception as e:
        print(f"❌ Error installing Pandoc: {e}")
        return False

if __name__ == '__main__':
    success = setup_local_pandoc()
    sys.exit(0 if success else 1)
