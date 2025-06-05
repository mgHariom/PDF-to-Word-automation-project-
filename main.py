from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import fitz  # PyMuPDF
import re

ignore_phrases = [
    "Grand Polycoats Co.Pvt.Ltd.",
    "Siddharth Complex, R.C. Dutt Road",
    "Tel.:  +91-265-3064200",
    "Fax:  +91-265-2337022",
    "E-mail:  marketing@grandpolycoats.com",
    "Website:  www.grandpolycoats.com",
    "ISO 9001",
    "ISO 14001",
    "This supersedes all earlier versions",
    "2/2"
]

heading_keywords = [
    "Description", "Recommended Use", "Physical Data",
    "Application Data", "Notes", "Health & Safety",
    "Limitation of Liability", "Application Method",
    "Additional Information", "Surface Preparation"
]

def extract_text_with_styles(pdf_path):
    doc = fitz.open(pdf_path)
    content = []

    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        lines = []

        for block in blocks:
            if block['type'] != 0:
                continue

            for line in block['lines']:
                line_text = ' '.join(span['text'] for span in line['spans']).strip()
                if not line_text or any(phrase.lower() in line_text.lower() for phrase in ignore_phrases):
                    continue

                span = line['spans'][0]
                y_pos = block['bbox'][1]
                font_size = span['size']
                font_name = span['font']
                is_bold = "bold" in font_name.lower()
                is_large = font_size > 12
                is_manual_heading = any(k.lower() == line_text.strip().lower() for k in heading_keywords)
                style = 'HEADING_1' if (is_manual_heading or is_large or is_bold) else 'NORMAL_TEXT'

                lines.append({
                    'text': line_text,
                    'y': y_pos,
                    'style': style
                })

        lines.sort(key=lambda l: l['y'])

        paragraph = ""
        current_style = None
        previous_y = None

        for line in lines:
            y = line['y']
            text = line['text']
            style = line['style']

            if current_style != style or (previous_y is not None and abs(y - previous_y) > 5):
                if paragraph:
                    content.append({'text': paragraph.strip(), 'style': current_style})
                paragraph = text
                current_style = style
            else:
                paragraph += " " + text

            previous_y = y

        if paragraph:
            content.append({'text': paragraph.strip(), 'style': current_style})

    return content

def extract_table_blocks(content_lines):
    table_data = []
    pattern = re.compile(r"^([^:]+):\s*(.+)$")
    buffer = []

    for line in content_lines:
        matches = pattern.findall(line["text"])
        if matches:
            for key, value in matches:
                buffer.append((key.strip(), value.strip()))
        else:
            if buffer:
                key, value = buffer[-1]
                buffer[-1] = (key, value + " " + line["text"].strip())

    return buffer

def clear_doc_content(service, document_id):
    doc = service.documents().get(documentId=document_id).execute()
    content = doc.get('body').get('content', [])
    if len(content) < 2:
        print("Document is empty or no content to delete.")
        return

    start_index = 1
    end_index = content[-1]['endIndex']

    # Check last character to adjust end_index
    last_paragraph = content[-1].get('paragraph', {})
    elements = last_paragraph.get('elements', [])
    if elements and 'textRun' in elements[-1]:
        text_content = elements[-1]['textRun']['content']
        if text_content.endswith('\n'):
            end_index -= 1

    if end_index <= start_index:
        print("No content to delete, delete range is empty.")
        return

    requests = [
        {
            "deleteContentRange": {
                "range": {
                    "startIndex": start_index,
                    "endIndex": end_index
                }
            }
        }
    ]

    service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()
    print(f"Deleted content from index {start_index} to {end_index}")


def insert_table_from_lines(service, document_id, insert_index, table_data):
    num_rows = len(table_data)
    if num_rows == 0:
        return
    requests = [{"insertTable": {"rows": num_rows, "columns": 2, "location": {"index": insert_index}}}]
    insert_offset = insert_index + 1
    for key, value in table_data:
        for cell_text in [key, value]:
            requests.append({
                "insertText": {
                    "text": cell_text,
                    "location": {"index": insert_offset}
                }
            })
            insert_offset += len(cell_text)
    service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()

def get_doc_end_index(service, document_id):
    doc = service.documents().get(documentId=document_id).execute()
    content = doc.get('body').get('content', [])
    return content[-1]['endIndex'] if content else 1

def insert_content_to_docs(service, document_id, content_lines):
    # Join all text with newlines
    full_text = ""
    for line in content_lines:
        text = line["text"]
        if line["style"] == "HEADING_1":
            text = text.upper()
        full_text += text + "\n"

    requests = [
        {
            "insertText": {
                "location": {"index": 1},
                "text": full_text
            }
        }
    ]

    # You can optionally add paragraph style update requests after insert if needed
    # But be aware indexes must be correct — this is simpler for initial insert

    service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()


def main():
    creds = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/documents'])
    service = build('docs', 'v1', credentials=creds)

    pdf_path = input("Enter path to PDF: ").strip()
    document_id = "1NRBaUvQCRiNyQZNu6gtLKM3qbkDCSySCGak3OQkb2kY"

    print("Extracting content from PDF...")
    content_lines = extract_text_with_styles(pdf_path)

    print("Clearing document content...")
    clear_doc_content(service, document_id)

    print("Inserting content into Google Doc...")
    insert_content_to_docs(service, document_id, content_lines)

    print("✅ Document updated successfully.")

if __name__ == "__main__":
    main()
