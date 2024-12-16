# Document Automation Tool

A powerful Python-based desktop application for automating document generation from templates. This tool allows you to create multiple personalized documents by combining a Word template with data from Excel or CSV files.

## Features

- **Template Support**: Uses Word documents (.docx) as templates with customizable keywords
- **Multiple Keyword Formats**: Supports various keyword formats including {{keyword}}, $$keyword, ##keyword##, {keyword}, [[keyword]], ((keyword, ||keyword, and @@keyword
- **Bulk Processing**: Process multiple records at once from Excel (.xlsx) or CSV files
- **Custom Formatting**: Customize font type, size, color, and style (bold, italic, underline) for each keyword
- **Multiple Output Formats**: Generate documents in both PDF and DOCX formats
- **Organized Output**: Automatically organizes output files in folders based on document type and custom folder names
- **Auto-matching**: Smart keyword-to-column matching system
- **Live Preview**: Preview formatting changes in real-time
- **User-friendly Interface**: Step-by-step wizard interface for easy navigation

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/document-automation-bot.git
cd document-automation-tool
```

2. Install required dependencies:
```bash
pip install python-docx docx2pdf pandas openpyxl tkinter
```

## How to Use

### Step 1: Upload Files
1. Launch the application
2. Click "Browse" to select your Word template file (.docx)
3. Click "Browse" to select your data file (.xlsx or .csv)

### Step 2: Keyword Matching
1. Review detected keywords from your template
2. Select which keywords you want to process
3. Match keywords with corresponding columns from your data file
4. (Optional) Customize formatting for each keyword:
   - Font type and size
   - Text color
   - Bold, italic, underline options

### Step 3: Output Settings
1. Choose output format(s):
   - PDF
   - Word Document (DOCX)
2. Select save location
3. Click "Generate Files" to process

## Template Creation Guidelines

Your template should include keywords in any of these formats:
- {{keyword}}
- $$keyword
- ##keyword##
- {keyword}
- [[keyword]]
- ((keyword
- ||keyword
- @@keyword

Example template text:
```
Dear {{name}},

Your account number is ##account_number##.
Balance: $$amount
```

## Data File Requirements

Your Excel or CSV file should include:
- Column headers matching or corresponding to template keywords
- A column containing names (e.g., "name", "full name", "fullname")
- (Optional) A "folder" column to specify custom output folders, if not will be declared as default

## Known Limitations

- PDF conversion requires Microsoft Word to be installed on the system
- Maximum file size depends on available system memory
- Template tables and complex formatting may have limited support

## Troubleshooting

1. **PDF Generation Fails**: Ensure Microsoft Word is installed and properly configured
2. **Missing Columns**: Verify column names in your data file match the keywords
3. **Formatting Issues**: Check template document for complex formatting that might interfere

## Contact

For support or feature requests, please contact:
- Sky (skylimsk@hotmail.com)

## License

This project is licensed under the MIT License - see the LICENSE file for details.
