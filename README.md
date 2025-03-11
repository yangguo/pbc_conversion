# Document Converter

A professional tool for creating comprehensive document reports from various file types. The utility scans directories recursively, processes different document formats, and generates a consolidated Word report with an interactive table of contents and internal navigation links.

## Features

- **Multi-format Support**: Processes PDF, Word documents (.doc, .docx), Excel files (.xls, .xlsx), images (.png, .jpg, .jpeg, .gif, .bmp), and text files (.txt, .log, .md, .csv)
- **Interactive Report**: Generated report includes a clickable table of contents and internal hyperlinks
- **Visual Previews**: Creates visual representations of documents and embeds them in the report
- **File Organization**: Organizes files with contextual information about their location
- **Smart Content Extraction**: Extracts relevant content from different file formats
- **Error Handling**: Robust error management with graceful fallbacks for problematic files

## Installation

1. Clone or download this repository
2. Create a virtual environment:
   ```
   python -m venv myenv
   ```
3. Activate the virtual environment:
   - Windows: `myenv\Scripts\activate`
   - macOS/Linux: `source myenv/bin/activate`
4. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

Run the script from the command line:

```
python document_converter.py -i "input_directory" -o "output_file.docx"
```

Arguments:
- `-i, --input`: Input directory containing files to process (required)
- `-o, --output`: Output file name (default: "report.docx")

## Requirements

The project uses the following Python packages:
- PyMuPDF (fitz)
- python-docx
- Pillow
- openpyxl
- pywin32

For a full list of dependencies, see `requirements.txt`.

## How It Works

1. The script recursively scans the input directory for files
2. Each file is processed according to its type:
   - **PDF files**: First page is rendered and added to the report
   - **Word documents**: First page content is extracted or converted via PDF
   - **Excel files**: Data is extracted or converted via PDF
   - **Images**: Added directly to the report
   - **Text/CSV files**: Content is extracted with proper encoding detection
3. A comprehensive report is generated with a table of contents and file index
4. Each entry contains file information, location context, and content preview

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.