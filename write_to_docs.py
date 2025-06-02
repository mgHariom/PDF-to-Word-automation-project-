from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

# Your authentication function, return service
def get_docs_service():
    creds = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/documents'])
    service = build('docs', 'v1', credentials=creds)
    return service

# Create a blank doc and return its ID
def create_doc_with_content():
    service = get_docs_service()
    doc = service.documents().create(body={'title': 'PDF Converted Document'}).execute()
    return doc['documentId']

def insert_content_with_headings(service, document_id, content_lines):
    requests = []
    index = 1  # Google Docs index starts at 1
    
    for line in content_lines:
        requests.append({
            "insertText": {
                "location": {"index": index},
                "text": line["text"] + "\n"
            }
        })
        
        requests.append({
            "updateParagraphStyle": {
                "range": {
                    "startIndex": index,
                    "endIndex": index + len(line["text"]) + 1
                },
                "paragraphStyle": {
                    "namedStyleType": line["style"]
                },
                "fields": "namedStyleType"
            }
        })
        
        index += len(line["text"]) + 1

    service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()


