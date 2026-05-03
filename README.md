# TLCAwardsParser
Trail Life Connect Court of Honor Awards Parser

## Required Packages
pyinstaller
bs4
openpyxl

## Generating Release Version
1. Complete any updates. Modify version number in the manual. Commit and push all changes.
2. Tag the git commit to be used for the release.
3. From a command window, run pyinstaller --onefile your_script.py
4. Save the manual to a pdf in the dist folder.
5. Copy AwardsInventory.xlsx, TL COH ceremony template.docx, and TL_Awards_Parser_settings.xml to the dist folder.
6. Zip the 5 files in the dist folder and post the zip file as a release package on github.