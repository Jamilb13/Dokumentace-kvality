import os
import fitz  # PyMuPDF
from PIL import Image

def ensure_pdf(file_path):
    """
    Zkontroluje soubor a v případě potřeby jej zkonvertuje na PDF.
    Vrací cestu k PDF (buď původní nebo nový konvertovaný soubor).
    """
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == '.pdf':
        return file_path
        
    pdf_path = os.path.splitext(file_path)[0] + "_converted.pdf"
    
    # Konverze obrázků (PNG, JPG, JPEG, BMP, TIFF, WEBP)
    if ext in ['.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.webp']:
        print(f"Konvertuji obrázek do PDF: {file_path}")
        img = Image.open(file_path)
        if img.mode != 'RGB':
            img = img.convert('RGB')
        img.save(pdf_path, "PDF", resolution=300.0)
        return pdf_path

    # Konverze Office dokumentů (DOCX, XLSX, PPTX) přes win32com pokud je k dispozici
    if ext in ['.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt']:
        try:
            import win32com.client
            print(f"Konvertuji Office dokument do PDF přes MS Office: {file_path}")
            if ext in ['.docx', '.doc']:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(os.path.abspath(file_path))
                doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17) # 17 = wdFormatPDF
                doc.Close()
                word.Quit()
                return pdf_path
            elif ext in ['.xlsx', '.xls']:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                wb = excel.Workbooks.Open(os.path.abspath(file_path))
                wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path)) # 0 = xlTypePDF
                wb.Close(False)
                excel.Quit()
                return pdf_path
        except Exception as e:
            print(f"Varování: Nepodařilo se použít MS Office pro konverzi {file_path}: {e}")

    raise ValueError(f"Nepodporovaný formát souboru pro konverzi na PDF: {ext}")

def handle_protected_pdf(pdf_path):
    """
    Zkontroluje, zda je PDF zamčené nebo chráněné proti úpravám.
    Pokud nejdou vložit úpravy, vyrenderuje jeho stránky na obrázky (300 DPI) a vytvoří čisté nechráněné PDF.
    Vrací fitz.Document objekt.
    """
    doc = fitz.open(pdf_path)
    
    # Pokud je soubor zašifrován heslem nebo chráněn
    if doc.is_encrypted:
        try:
            doc.authenticate("") # zkusit prázdné heslo
        except Exception:
            pass
            
    # Zjistit zda můžeme upravovat
    can_modify = True
    if doc.is_encrypted:
        can_modify = False
        
    if can_modify:
        try:
            # Zkusit otevřít pro úpravy
            return doc
        except Exception:
            can_modify = False

    if not can_modify:
        print(f"Detekováno chráněné/zamčené PDF {pdf_path}. Provádím rasterizaci přes obrázky (300 DPI)...")
        clean_doc = fitz.open()
        for page_num in range(len(doc)):
            page = doc[page_num]
            pix = page.get_pixmap(dpi=300)
            img_pdf = fitz.open("pdf", pix.tobytes("pdf"))
            clean_doc.insert_pdf(img_pdf)
        doc.close()
        return clean_doc
        
    return doc
