import pdf2docx
import docx2txt

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
    
    pdf_file = input("Enter the name of the PDF file (including extension): ").strip()
    
    if not pdf_file:
        print("No PDF file specified. Exiting.")
        return

    docx_file = input("Enter the name for the output Word document (including extension, press Enter for default): ").strip()

    if not docx_file:
        docx_file = pdf_file.replace(".pdf", ".docx")

    convert_pdf_to_docx(pdf_file, docx_file)

    copy_text_option = input("Copy text to a separate .txt file (y/n)? ").strip().lower()
    if copy_text_option == "y":
        save_as_txt(docx_file)

if __name__ == "__main__":
    main()
