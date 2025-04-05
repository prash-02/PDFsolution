import os
import shutil
from datetime import datetime

def create_distribution():
    # Create distribution directory
    dist_dir = "PDFTableExtractorPro_Distribution"
    if os.path.exists(dist_dir):
        shutil.rmtree(dist_dir)
    os.makedirs(dist_dir)

    # Copy the executable
    shutil.copy2("dist/PDFTableExtractorPro.exe", dist_dir)

    # Create README
    readme_content = """PDFTableExtractorPro
===================
Version: 1.0
Build Date: {}

Installation
------------
1. Extract all files to a folder
2. Run PDFTableExtractorPro.exe
3. On first run, you'll need to activate the software with your license key

Requirements
-----------
- Windows 10 or later
- 4GB RAM minimum
- Java Runtime Environment (JRE) 8 or later

Support
-------
For support, contact: your-email@example.com
""".format(datetime.now().strftime("%Y-%m-%d"))

    with open(os.path.join(dist_dir, "README.txt"), "w") as f:
        f.write(readme_content)

    # Create final zip
    shutil.make_archive(f"PDFTableExtractorPro_{datetime.now().strftime('%Y%m%d')}", 
                       'zip', dist_dir)
    
    print(f"Distribution package created successfully!")

if __name__ == "__main__":
    create_distribution()
