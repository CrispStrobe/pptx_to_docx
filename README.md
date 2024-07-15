# PPTX to DOCX and RTF Converter

This script extracts text from a PowerPoint (.pptx) file and saves it as both a Word document (.docx) and a Rich Text Format (.rtf) file. 

## Requirements

- Python 3.6+
- `python-pptx`
- `python-docx`
- `pypandoc`

## Installation

1. Clone the repository or download the `pptx_2_docx.py` script.
2. Install the required Python packages:

    ```sh
    pip install python-pptx python-docx pypandoc
    ```

3. Install Pandoc, if not already installed. You can download it from the [Pandoc official site](https://pandoc.org/installing.html).

## Usage

Run the script with the path to your PowerPoint file as an argument:

    ```sh
    python pptx_2_docx.py <input_pptx_file>
    ```

### Example

    ```sh
    python pptx_2_docx.py example.pptx
    ```
    
This command will generate `example.docx` and `example.rtf` in the same directory as your input file.

## Script Explanation

1. **Text Extraction**: The script extracts text from each slide in the PowerPoint file.
2. **Text Cleaning**: It removes non-allowed characters and vertical tab characters from the extracted text.
3. **Save as DOCX**: The cleaned text is saved to a DOCX file using the same base name as the input file.
4. **Save as RTF**: The DOCX file is converted to an RTF file using `pypandoc`.
