import os
import subprocess
import sys

def build_exe():
    # Check if icon exists
    icon_path = 'icon.ico'
    icon_config = f"icon='{icon_path}'," if os.path.exists(icon_path) else "icon=None,"
    
    # Create spec file content
    spec_content = f'''
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['PdfToExcel.py'],
             pathex=['E:\\Java'],
             binaries=[],
             datas=[],
             hiddenimports=['PIL._tkinter_finder'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data,
          cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='PDFTableExtractorPro',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          {icon_config}
          bundle_identifier='com.pdftools.extractor')
'''
    
    # Write spec file
    with open('PdfToExcel.spec', 'w') as f:
        f.write(spec_content)
    
    # Build EXE
    subprocess.call(['pyinstaller', '--clean', 'PdfToExcel.spec'])

if __name__ == '__main__':
    build_exe()
