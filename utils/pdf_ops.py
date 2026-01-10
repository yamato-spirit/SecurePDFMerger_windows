import os
import io
import sys
import platform
import tempfile
import pandas as pd
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image

# PDF生成用
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from xhtml2pdf import pisa

# Office変換用
try:
    import comtypes.client
    HAS_COM = True
except ImportError:
    HAS_COM = False

try:
    from docx2pdf import convert as docx_convert
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

def get_japanese_font_path():
    system = platform.system()
    candidates = []
    if system == "Windows":
        candidates = [r"C:\Windows\Fonts\msgothic.ttc", r"C:\Windows\Fonts\meiryo.ttc", r"C:\Windows\Fonts\msmincho.ttc"]
    elif system == "Darwin":
        candidates = ["/System/Library/Fonts/Hiragino Sans GB.ttc", "/System/Library/Fonts/AppleGothic.ttf"]
    else:
        candidates = ["/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"]

    for path in candidates:
        if os.path.exists(path): return path.replace('\\', '/')
    return None

def text_to_pdf(source_path, output_path, font_path, font_name, is_landscape=False):
    try:
        page_size = landscape(A4) if is_landscape else portrait(A4)
        c = canvas.Canvas(output_path, pagesize=page_size)
        width, height = page_size
        
        margin = 20 * mm
        y_position = height - margin
        line_height = 14
        
        if font_path:
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                c.setFont(font_name, 10.5)
            except: c.setFont("Helvetica", 10.5)
        else: c.setFont("Helvetica", 10.5)

        try:
            with open(source_path, 'r', encoding='utf-8') as f: lines = f.readlines()
        except:
            with open(source_path, 'r', encoding='cp932', errors='ignore') as f: lines = f.readlines()

        for line in lines:
            line = line.rstrip()
            if y_position < margin:
                c.showPage()
                y_position = height - margin
                if font_path: c.setFont(font_name, 10.5)
                else: c.setFont("Helvetica", 10.5)
            c.drawString(margin, y_position, line)
            y_position -= line_height
        c.save()
        return True
    except Exception as e:
        print(f"Text error: {e}")
        return False

def excel_csv_to_html(source_path, is_landscape=False):
    try:
        if source_path.lower().endswith('.csv'):
            df = pd.read_csv(source_path, encoding='utf-8', on_bad_lines='skip')
        else:
            df = pd.read_excel(source_path)
        
        html = df.to_html(index=False, border=1, classes='table')
        size_css = "size: A4 landscape;" if is_landscape else "size: A4 portrait;"
        style = f"""
        <style>
            @page {{ {size_css} margin: 1cm; }}
            table {{ border-collapse: collapse; width: 100%; }}
            th, td {{ border: 1px solid black; padding: 4px; text-align: left; font-size: 9pt; }}
            th {{ background-color: #f0f0f0; }}
        </style>
        """
        return f"<html><head><meta charset='utf-8'>{style}</head><body>{html}</body></html>"
    except Exception as e:
        return None

def convert_to_pdf(source_path, is_landscape=False):
    base_name = os.path.basename(source_path)
    if base_name.startswith('.') and base_name.count('.') == 1:
        ext = base_name.lower()
    else:
        ext = os.path.splitext(source_path)[1].lower()

    temp_dir = tempfile.gettempdir()
    safe_name = "".join([c for c in base_name if c.isalnum() or c in (' ', '.', '_')]).strip()
    orientation_suffix = "_L" if is_landscape else "_P"
    output_filename = f"converted_{safe_name}{orientation_suffix}.pdf"
    output_path = os.path.join(temp_dir, output_filename)
    
    font_path = get_japanese_font_path()
    font_name = "JapaneseFont"

    try:
        if ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif', '.eps', '.webp', '.ico']:
            image = Image.open(source_path)
            if image.mode != 'RGB': image = image.convert('RGB')
            image.save(output_path, "PDF", resolution=100.0)
            return output_path
        elif ext in ['.txt', '.log', '.md', '.py', '.json', '.xml', '.js', '.css', '.rtf']:
            if text_to_pdf(source_path, output_path, font_path, font_name, is_landscape):
                return output_path
            return None
        elif ext in ['.html', '.htm', '.xlsx', '.xls', '.csv']:
            source_html = ""
            if ext in ['.html', '.htm']:
                try:
                    with open(source_path, 'r', encoding='utf-8') as f: source_html = f.read()
                except:
                    with open(source_path, 'r', encoding='cp932', errors='ignore') as f: source_html = f.read()
                size_css = "size: A4 landscape;" if is_landscape else "size: A4 portrait;"
                if "@page" not in source_html:
                     source_html = f"<style>@page {{ {size_css} }}</style>" + source_html
            else:
                source_html = excel_csv_to_html(source_path, is_landscape)
                if not source_html: return None
            if font_path:
                try:
                    pdfmetrics.registerFont(TTFont(font_name, font_path))
                    css = f"<style>@font-face {{ font-family: '{font_name}'; src: url('{font_path}'); }} body {{ font-family: '{font_name}'; }}</style>"
                    if "<head>" in source_html: source_html = source_html.replace("<head>", f"<head>{css}")
                    else: source_html = f"{css}{source_html}"
                except: pass
            with open(output_path, "wb") as f:
                pisa.CreatePDF(io.StringIO(source_html), dest=f, encoding='utf-8')
            return output_path
        elif ext in ['.docx', '.doc'] and HAS_DOCX:
            docx_convert(source_path, output_path)
            return output_path
        elif ext in ['.pptx', '.ppt'] and HAS_COM:
            try:
                ppt = comtypes.client.CreateObject("Powerpoint.Application")
                deck = ppt.Presentations.Open(source_path, WithWindow=False)
                deck.SaveAs(output_path, 32)
                deck.Close()
                ppt.Quit()
                return output_path
            except: return None
        return None
    except Exception as e:
        print(f"Conversion error: {e}")
        return None

def get_pdf_info(file_path):
    if not os.path.exists(file_path): return None
    try:
        reader = PdfReader(file_path)
        if reader.is_encrypted:
            try: return {"path": file_path, "filename": os.path.basename(file_path), "pages": len(reader.pages), "page_details": []}
            except: return None
        details = []
        for page in reader.pages:
            try:
                w, h = float(page.mediabox.width), float(page.mediabox.height)
                details.append({'is_portrait': h >= w, 'width': w, 'height': h})
            except:
                details.append({'is_portrait': True})
        return {
            "path": file_path,
            "filename": os.path.basename(file_path),
            "pages": len(reader.pages),
            "page_details": details
        }
    except: return None

def merge_pdfs_securely(page_list, output_path=None):
    writer = PdfWriter()
    try:
        open_files = {}
        for item in page_list:
            path = item['path']
            if path not in open_files: open_files[path] = PdfReader(path)
            reader = open_files[path]
            if 0 <= item['page_index'] < len(reader.pages):
                page = reader.pages[item['page_index']]
                if item['rotation'] != 0:
                    page.rotate(item['rotation'])
                writer.add_page(page)
        writer.add_metadata({'/Title': '', '/Author': ''})
        if output_path:
            with open(output_path, "wb") as f: writer.write(f)
            return True
        else:
            out = io.BytesIO()
            writer.write(out)
            out.seek(0)
            return out
    except Exception as e:
        print(f"Merge error: {e}")
        return False if output_path else None

def get_preview_image(pdf_bytes_io, page_num):
    try:
        doc = fitz.open(stream=pdf_bytes_io, filetype="pdf")
        if page_num >= len(doc): return None
        pix = doc.load_page(page_num).get_pixmap(dpi=150)
        return Image.open(io.BytesIO(pix.tobytes("png")))
    except: return None

def get_page_thumbnail(file_path, page_num, rotation=0):
    try:
        doc = fitz.open(file_path)
        if page_num >= len(doc): return None
        page = doc.load_page(page_num)
        if rotation != 0:
            page.set_rotation(rotation)
        pix = page.get_pixmap(dpi=72)
        return Image.open(io.BytesIO(pix.tobytes("png")))
    except Exception as e:
        print(f"Thumbnail error: {e}")
        return None