from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import fitz  # PyMuPDF

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

def extract_text_with_styles(pdf_path):
    doc = fitz.open(pdf_path)
    content = []

    heading_keywords = [
        "Description", "Recommended Use", "Physical Data" ,
        "Application Data", "Notes",
        "Health & Safety", "Limitation of Liability", "Application Method",
        "Additional Information"
    ]

    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        lines = []

        for block in blocks:
            if block['type'] != 0:
                continue

            for line in block['lines']:
                line_text = ' '.join(span['text'] for span in line['spans']).strip()
                if not line_text:
                    continue

                # Check if line contains any ignored phrase (case-insensitive)
                if any(phrase.lower() in line_text.lower() for phrase in ignore_phrases):
                    continue  # Skip this line completely

                span = line['spans'][0]
                y_pos = block['bbox'][1]
                font_size = span['size']
                font_name = span['font']
                is_bold = "bold" in font_name.lower()
                is_large = font_size > 12
                is_manual_heading = any(k.lower() in line_text.lower() for k in heading_keywords)

                style = 'HEADING_1' if (is_manual_heading or is_large or is_bold) else 'NORMAL_TEXT'

                lines.append({
                    'text': line_text,
                    'y': y_pos,
                    'style': style
                })

        # Sort lines top to bottom
        lines.sort(key=lambda l: l['y'])

        # Group into paragraphs
        paragraph = ""
        current_style = None
        previous_y = None

        for line in lines:
            y = line['y']
            text = line['text']
            style = line['style']

            if current_style != style or (previous_y is not None and abs(y - previous_y) > 5):  # New paragraph
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

# def insert_content_to_docs(service, document_id, content_lines):
    requests = []
    index = 1
    for line in content_lines:
        text = line["text"] + "\n"
        length = len(text)
        requests.append({
            "insertText": {
                "location": {"index": index},
                "text": text
            }
        })
        requests.append({
            "updateParagraphStyle": {
                "range": {
                    "startIndex": index,
                    "endIndex": index + length
                },
                "paragraphStyle": {
                    "namedStyleType": line["style"]
                },
                "fields": "namedStyleType"
            }
        })
        index += length

    service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()
def insert_content_to_docs(service, document_id, content_lines):
    requests = []
    index = 1
    for line in content_lines:
        text = line["text"] + "\n"
        length = len(text)
        requests.append({
            "insertText": {
                "location": {"index": index},
                "text": text
            }
        })
        requests.append({
            "updateParagraphStyle": {
                "range": {
                    "startIndex": index,
                    "endIndex": index + length
                },
                "paragraphStyle": {
                    "namedStyleType": line["style"]
                },
                "fields": "namedStyleType"
            }
        })
        index += length

    service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()

def clear_doc_content(service, document_id):
    doc = service.documents().get(documentId=document_id).execute()
    content = doc.get('body').get('content')

    if not content or len(content) < 2:
        print("Document is empty or no content to delete.")
        return

    start_index = 1
    end_index = content[-1]['endIndex']

    # Fetch the document text to check last char
    text_content = ""
    for element in content:
        if 'paragraph' in element:
            for elem in element['paragraph']['elements']:
                if 'textRun' in elem:
                    text_content += elem['textRun']['content']

    # If last character is newline, reduce end_index by 1
    if text_content.endswith('\n'):
        end_index -= 1

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

def insert_text(service, document_id, text):
    requests = [
        {
            "insertText": {
                "location": {"index": 1},
                "text": text
            }
        }
    ]
    service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()

def rewrite_doc_with_text(service, document_id, content_lines):
    print("Clearing existing document content...")
    clear_doc_content(service, document_id)

    # Join all lines into one string, preserve line breaks
    text = "\n".join(line['text'] for line in content_lines)

    print("Inserting new text content...")
    insert_text(service, document_id, text)
    print("Done rewriting the document.")

def insert_content_to_docs(service, document_id, content_lines):
    requests = []
    index = 1
    for line in content_lines:
        text = line["text"]
        style = line["style"]

        if style == "HEADING_1":
            text = text.upper()  # Convert heading to uppercase

        text += "\n"
        length = len(text)

        # Insert text
        requests.append({
            "insertText": {
                "location": {"index": index},
                "text": text
            }
        })

        # Update paragraph style (e.g. HEADING_1 or NORMAL_TEXT)
        requests.append({
            "updateParagraphStyle": {
                "range": {
                    "startIndex": index,
                    "endIndex": index + length
                },
                "paragraphStyle": {
                    "namedStyleType": style
                },
                "fields": "namedStyleType"
            }
        })

        # For headings, also make text bold
        if style == "HEADING_1":
            requests.append({
                "updateTextStyle": {
                    "range": {
                        "startIndex": index,
                        "endIndex": index + length - 1  # exclude newline from bold
                    },
                    "textStyle": {
                        "bold": True
                    },
                    "fields": "bold"
                }
            })

        index += length

    service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()


def main():
    creds = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/documents'])
    service = build('docs', 'v1', credentials=creds)

    pdf_path = input("Enter path to PDF: ")
    document_id = "18E7OyGAomyzKZ-xv1HN8NjRAK4NG-iOi6Y0ds2k5_iw"

    content_lines = extract_text_with_styles(pdf_path)
    
    # Either insert with styles (if you implement that function)
    insert_content_to_docs(service, document_id, content_lines)
    
    # Or just rewrite plain text (recommended for simplicity)
    # rewrite_doc_with_text(service, document_id, content_lines)

    print("✅ PDF content successfully inserted into Google Doc.")

if __name__ == "__main__":
    main()
