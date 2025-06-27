# PDF to Excel Processing Tool

A smart automation tool that extracts data from purchase order PDFs and automatically fills Excel templates. This tool saves hours of manual data entry by using AI to read PDF documents and populate structured Excel files.

## 🎯 Project Overview

This project was built to solve a common business problem: manually copying information from PDF purchase orders into Excel spreadsheets is time-consuming and error-prone. Our tool uses Google's Gemini AI to automatically extract key information from PDFs and populate pre-designed Excel templates.

## ✨ What This Tool Does

- **Reads PDF purchase orders** and extracts important information like client names, order numbers, product details, and specifications
- **Translates content** from various languages into Slovak (where applicable)
- **Fills Excel templates** automatically with the extracted data
- **Processes multiple files** at once - just drop your PDFs in a folder and let it work
- **Handles complex layouts** including multi-page documents and varying PDF formats

## 🚀 Key Features

### Intelligent Data Extraction
- Client information and order numbers
- Product specifications (foil types, heating requirements, etc.)
- Item details (names, quantities, sizes, volumes)
- Special order requirements and embossing data
- Bulk container and microbiological analysis requirements

### Smart Processing
- **Multi-language support** with automatic translation to Slovak
- **Error handling** with retry mechanisms for API rate limits
- **Format preservation** maintains Excel styling and layouts
- **Batch processing** handles multiple PDFs automatically

### Excel Template Integration
- Preserves original formatting and styles
- Automatically adjusts for varying numbers of products
- Maintains proper cell alignment and merged cell structures
- Creates properly formatted output files

## 🛠️ How It Works

1. **PDF Analysis**: The tool uses Google's Gemini AI to read and understand PDF content
2. **Data Extraction**: Extracts 13 specific fields including client info, specifications, and product details
3. **Translation**: Converts relevant text to Slovak while preserving technical terms
4. **Excel Population**: Maps extracted data to the correct cells in your Excel template
5. **File Generation**: Creates new Excel files with all the extracted information

## 📋 Prerequisites

Before you start, make sure you have:

- Python 3.7 or higher installed
- A Google Gemini API key (free to get from Google AI Studio)
- The required Python packages (see installation section)

## 🔧 Installation

1. **Clone or download this project** to your computer

2. **Install required packages** by running this command in your terminal:
```bash
pip install openpyxl google-generativeai python-dotenv
```

3. **Set up your API key**:
   - Get a free Gemini API key from [Google AI Studio](https://makersuite.google.com/app/apikey)
   - Create a file named `.env` in your project folder
   - Add this line to the `.env` file: `GEMINI_API_KEY=your_api_key_here`

4. **Prepare your folders**:
   - The tool will create necessary folders automatically when you first run it

## 📁 Folder Structure

Your project should look like this:
```
your-project/
├── pdf_to_excel_processor.py     # The main script
├── .env                       # Your API key file
├── input_pdfs/               # Put your PDF files here
├── files/                    # Contains the Excel template
│   └── empty file for extraction excel file.xlsx
└── output_excel/             # Generated Excel files appear here
```

## 🚀 How to Use

### Quick Start
1. **Place your PDF files** in the `input_pdfs` folder
2. **Make sure** your Excel template is in the `files` folder
3. **Run the script**: `python pdf_to_excel_processor.py`
4. **Check the results** in the `output_excel` folder

### Step-by-Step Usage

1. **Prepare Your Files**
   - Put all PDF purchase orders you want to process in the `input_pdfs` folder
   - Ensure your Excel template is named correctly and placed in the `files` folder

2. **Run the Processing**
   ```bash
   python pdf_to_excel_processor.py
   ```

3. **Monitor Progress**
   - The tool will show you progress as it processes each file
   - You'll see messages like "Processing file 1 of 5: order_123.pdf"
   - Any errors or issues will be clearly displayed

4. **Review Results**
   - Check the `output_excel` folder for your completed files
   - Each PDF will create a corresponding Excel file with "_filled_order_note" added to the name

### What Gets Extracted

The tool looks for and extracts these key pieces of information:

**Basic Information:**
- Client name
- Order number (formatted as O + numbers)
- Foil specifications
- Container return requirements
- Microbiological analysis needs

**Product Details (for each item):**
- Product name and article numbers
- Sachet dimensions and filling volumes
- Heating requirements
- Embossing data
- Required quantities
- Bulk quantities needed

**Special Requirements:**
- Custom packaging instructions
- Archive requirements
- Product handling specifications
- Any other order-specific notes

## ⚠️ Important Notes

### API Rate Limits
- The tool includes smart retry logic for Google's API rate limits
- If you're processing many files, the tool will automatically pace itself
- You might see "Rate limit exceeded" messages - this is normal and the tool will retry

### File Formats
- **Input**: Only PDF files are supported
- **Output**: Excel files (.xlsx format)
- Make sure your PDFs contain text (not just scanned images)

### Language Handling
- The tool automatically translates most content to Slovak
- Technical terms and product codes remain in their original language
- Item names are kept in their original language for accuracy

## 🐛 Troubleshooting

### Common Issues and Solutions

**"No PDF files found"**
- Make sure your PDF files are in the `input_pdfs` folder
- Check that files have .pdf extension

**"Template file not found"**
- Verify the Excel template is in the `files` folder
- Make sure it's named exactly: `empty file for extraction excel file.xlsx`

**"JSON parsing error"**
- This usually means the PDF format is unusual
- The tool will automatically retry
- If it keeps failing, the PDF might not contain the expected structure

**"Rate limit exceeded"**
- This is normal when processing many files
- The tool will wait and retry automatically
- Just be patient and let it work

**"Error reading PDF"**
- The PDF file might be corrupted or password-protected
- Try opening the PDF manually to verify it works
- Make sure the file isn't being used by another program

### Getting Help

If you run into issues:
1. Check the error messages - they usually explain what's wrong
2. Make sure all your files are in the right folders
3. Verify your API key is set up correctly
4. Try processing one file at a time to isolate problems

## 🔮 Future Improvements

We're constantly working to make this tool better. Planned improvements include:
- Support for more PDF layouts and formats
- Additional language translation options
- Better error recovery and reporting
- GUI interface for easier use
- Integration with cloud storage services

## 🤝 Contributing

Found a bug or have an idea for improvement? We'd love to hear from you! Feel free to:
- Report issues you encounter
- Suggest new features
- Share PDFs that don't process correctly (with sensitive info removed)

## 📄 License

This project is provided as-is for educational and business use. Feel free to modify it for your specific needs.

---

**Happy Processing!** 🎉

*This tool was built to make your work easier. If it saves you time, we've done our job right.*