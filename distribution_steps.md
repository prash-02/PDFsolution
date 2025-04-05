# PDF Table Extractor Pro - Distribution Process

## 1. Preparing Release
1. Update version number in source code
2. Run tests
3. Build executable:
   ```bash
   python build_exe.py
   ```
4. Create distribution package:
   ```bash
   python create_distribution.py
   ```

## 2. License Generation Process
1. Get hardware ID from user's activation screen
2. Run license generator:
   ```bash
   python license_generator.py
   ```
3. Send license file to user

## 3. Distribution Package Contents
- PDFTableExtractorPro.exe
- README.txt
- requirements.txt

## 4. Licensing Types
1. Trial (30 days)
2. Professional (1 year)
3. Enterprise (unlimited)

## 5. Security Measures
- Hardware-locked licensing
- Digital signature verification
- Expiration date enforcement
- Anti-tampering checks

## 6. Support Process
1. User submits hardware ID
2. Generate unique license
3. Send license file to user
4. Track licenses in database
