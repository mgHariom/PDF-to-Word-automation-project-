import fitz  # PyMuPDF

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    lines = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        blocks = page.get_text("dict")["blocks"]
        
        for b in blocks:
            if b['type'] == 0:  # text block
                for line in b["lines"]:
                    line_text = ""
                    max_font_size = 0
                    for span in line["spans"]:
                        line_text += span["text"]
                        if span["size"] > max_font_size:
                            max_font_size = span["size"]
                    line_text = line_text.strip()
                    if not line_text:
                        continue
                    
                    # Define heading threshold based on font size
                    if max_font_size > 15:
                        style = "HEADING_1"
                    elif max_font_size > 12:
                        style = "HEADING_2"
                    else:
                        style = "NORMAL_TEXT"
                    
                    lines.append({
                        "text": line_text,
                        "style": style
                    })
    return lines





import pdfplumber

def extract_content_from_pdf(pdf_path):
    content = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(extra_attrs=["fontname", "size"])
            for word in words:
                text = word['text']
                font_size = word['size']
                font_name = word['fontname']
                
                # Use a simple heuristic for heading detection
                if font_size > 14:
                    style = "HEADING_1"
                elif font_size > 12:
                    style = "HEADING_2"
                else:
                    style = "NORMAL_TEXT"

                content.append({
                    "type": "text",
                    "style": style,
                    "text": text
                })

            # Extract tables (if any)
            tables = page.extract_tables()
            for table in tables:
                content.append({
                    "type": "table",
                    "data": table
                })

    return content

