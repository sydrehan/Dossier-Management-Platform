import sys
import os
from pathlib import Path

def convert_pdf_to_docx(input_path: Path, output_path: Path):
    from pdf2docx import Converter
    print(f"Converting PDF: {input_path.name} -> {output_path.name}")
    cv = Converter(str(input_path))
    cv.convert(str(output_path))
    cv.close()

def convert_ppt_to_pdf(input_path: Path, output_path: Path):
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("PowerPoint conversion requires Windows and MS Office (win32com). Not supported on this server.")
    print(f"Converting PowerPoint to PDF: {input_path.name} -> {output_path.name}")
    abs_input = str(input_path.resolve())
    abs_output = str(output_path.resolve())
    
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    try:
        presentation = powerpoint.Presentations.Open(abs_input, WithWindow=False)
        # Format 32 is ppSaveAsPDF
        presentation.SaveAs(abs_output, 32)
        presentation.Close()
    finally:
        try:
            powerpoint.Quit()
        except Exception:
            pass

def main():
    if len(sys.argv) < 4:
        print("Usage: python convert.py <type> <input_path> <output_path>")
        sys.exit(1)
        
    conv_type = sys.argv[1].lower() # 'pdf-to-docx', 'ppt-to-pdf', 'ppt-to-docx'
    input_path = Path(sys.argv[2])
    output_path = Path(sys.argv[3])
    
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        sys.exit(1)
        
    try:
        if conv_type == 'pdf-to-docx':
            convert_pdf_to_docx(input_path, output_path)
        elif conv_type == 'ppt-to-pdf':
            convert_ppt_to_pdf(input_path, output_path)
        elif conv_type == 'ppt-to-docx':
            # PPT to PDF first
            temp_pdf = input_path.with_suffix(".temp.pdf")
            convert_ppt_to_pdf(input_path, temp_pdf)
            # PDF to DOCX
            convert_pdf_to_docx(temp_pdf, output_path)
            # Cleanup temp PDF
            if temp_pdf.exists():
                os.remove(temp_pdf)
        else:
            print(f"Unsupported conversion type: {conv_type}")
            sys.exit(1)
            
        print("Conversion successful.")
        
    except Exception as e:
        print(f"Conversion failed: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
