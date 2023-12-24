import pdf2docx
import docx2txt
import argparse
import argcomplete

def convert_pdf_to_docx(pdf_file, docx_file):
    try:
        pdf2docx.Converter(pdf_file).convert(docx_file)
        print(f"Conversion successful: {pdf_file} -> {docx_file}")
    except Exception as e:
        print(f"Conversion failed: {e}")

def save_as_txt(docx_file):
    try:
        text = docx2txt.process(docx_file)
        txt_file = docx_file.replace(".docx", ".txt")
        with open(txt_file, "w", encoding="utf-8") as txt:
            txt.write(text)
        print(f"Text saved to: {txt_file}")
    except Exception as e:
        print(f"Text extraction failed: {e}")

def print_banner():
    banner = r"""
    .----------------.  .----------------.  .----------------.  .----------------.   
    | .--------------. || .--------------. || .--------------. || .--------------. |  
    | |   ______     | || |  ____  ____  | || |  ________    | || |  _________   | |  
    | |  |_   __ \   | || | |_  _||_  _| | || | |_   ___ `.  | || | |_   ___  |  | |  
    | |    | |__) |  | || |   \ \  / /   | || |   | |   `. \ | || |   | |_  \_|  | |  
    | |    |  ___/   | || |    \ \/ /    | || |   | |    | | | || |   |  _|      | |  
    | |   _| |_      | || |    _|  |_    | || |  _| |___.' / | || |  _| |_       | |  
    | |  |_____|     | || |   |______|   | || | |________.'  | || | |_____|      | |  
    | |              | || |              | || |              | || |              | |  
    | '--------------' || '--------------' || '--------------' || '--------------' |  
     '----------------'  '----------------'  '----------------'  '----------------'   
    """
    print(banner)

def main():
    print_banner()
    print("Welcome to PyDF - PDF to Word Converter")

    parser = argparse.ArgumentParser(description="Convert PDF to Word document")
    parser.add_argument("pdf_file", help="Name of the PDF file (including extension)")
    parser.add_argument("--docx_file", help="Name for the output Word document (including extension, default is PDF filename with .docx extension)")
    parser.add_argument("-c", "--copy_text", action="store_true", help="Copy text to a separate .txt file")

    argcomplete.autocomplete(parser)
    args = parser.parse_args()

    pdf_file = args.pdf_file
    docx_file = args.docx_file

    if not docx_file:
        docx_file = pdf_file.replace(".pdf", ".docx")

    convert_pdf_to_docx(pdf_file, docx_file)

    if args.copy_text:
        save_as_txt(docx_file)

if __name__ == "__main__":
    main()
