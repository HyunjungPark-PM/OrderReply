<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# p-net Order Reply Tool - Project Instructions

## Project Overview
Personal Python project to automate the comparison of p-net download excel files with factory reply files and generate manual upload files for p-net system.

## Key Requirements
- Language: Python 3.8+
- Dependencies: openpyxl, pandas
- UI: Tkinter-based GUI
- Input: Two Excel files (p-net download and factory reply)
- Output: Combined Excel file in p-net upload format

## File Comparison Logic
- Comparison key: PO# + PO-LINE# (combination)
- Unique identifier: CPO# + CPO-LINE# + LINE SEQ (from p-net file)
- Output includes data from both files merged on matching PO# and LINE#

## Output File Format (in order)
1. PO# (from factory reply)
2. PO-LINE# (from factory reply)
3. Material (from factory reply)
4. CPO QTY (from factory reply)
5. ETD (from factory reply)
6. EX-F (from factory reply)
7. 내부노트 (from factory reply)
8. CPO# (from p-net)
9. CPO-LINE# (from p-net)
10. LINE SEQ (from p-net)

## Development Guidelines
- Maintain modular structure with separate modules for UI and business logic
- Use pandas for Excel processing
- Implement error handling for file reading
- Provide clear status messages in GUI
- Use background threading for file processing to prevent UI freezing
- Korean language support in UI and error messages

## Setup Complete
Project scaffolding completed with:
- ✓ main.py: GUI application
- ✓ excel_processor.py: Excel processing logic
- ✓ requirements.txt: Dependencies
- ✓ README.md: Documentation
