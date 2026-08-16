from msal import ConfidentialClientApplication
import os
from dotenv import load_dotenv
import requests
import json
from dataclasses import dataclass
from io import BytesIO
from docx import Document
from pypdf import PdfReader
from openpyxl import load_workbook
from pptx import Presentation

@dataclass
class SharePointFile:
    id: str
    site_id: str
    drive_id: str
    name: str
    mime_type: str
    size: int
    web_url: str
    created_at: str
    modified_at: str
    parent_id: str

@dataclass
class DocumentText:
    file_id: str
    filename: str
    text: str

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")


AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

result = app.acquire_token_for_client(scopes=SCOPES)

if "access_token" in result:
    access_token = result["access_token"]
else:
    raise Exception("No Access Token found")

headers = {
    "Authorization": f"Bearer {access_token}",
}

response = requests.get(
    "https://graph.microsoft.com/v1.0/sites?search=*",
    headers=headers
)

print(response.status_code)
#print(response.text)

it_ressort_site_id = "wi3.sharepoint.com,ee22c8d5-9932-421e-baa5-7d7c7af97e38,a59f0326-2de6-4360-8a55-1d0c89783f0e"

response = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives",
    headers=headers
)
print(response.status_code)
#print(response.text)

it_ressort_drive_id = "b!1cgi7jKZHkK6pX18evl-OCYDn6XmLWBDilUdDIl4Pw5jmHe4bh3yQpgHh_5Nhor2"

response = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{it_ressort_drive_id}/root/children",
    headers=headers
)
print(response.status_code)
#print(response.text)
'''
strategy_days_item_id = "01HODMHCWHUWZRJC72YNF2WTWTRQIOKKTJ"
response = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{it_ressort_drive_id}/items/{strategy_days_item_id}",
    headers=headers
)
print(response.status_code)
#print(response.text)

data = response.json()
#print(data.keys())

strategy_meta = SharePointFile(
    id=data["id"],
    site_id=data["parentReference"]["siteId"],
    drive_id=data["parentReference"]["driveId"],
    name=data["name"],
    mime_type=data["file"]["mimeType"],
    size=data["size"],
    web_url=data["webUrl"],
    created_at=data["createdDateTime"],
    modified_at=data["lastModifiedDateTime"],
    parent_id=data["parentReference"]["id"]
)
#print(strategy_meta)

response = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{it_ressort_drive_id}/items/{strategy_days_item_id}/content",
    headers=headers
)

print(response.status_code)

document = Document(BytesIO(response.content))

paragraphs = []
for paragraph in document.paragraphs:
    if paragraph.text.strip():
        paragraphs.append(paragraph.text)

text = "\n".join(paragraphs)

strategy_doc = DocumentText(
    file_id=strategy_meta.id,
    filename=strategy_meta.name,
    text=text
)
print(strategy_doc)
'''
def process_root(rootId):
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{rootId}/root/children",
        headers=headers
    )
    data = response.json()["value"]

    for item in data:
        if "folder" in item:
            print("Ordner:", item["name"])
            process_folder(rootId, item["id"])

        elif "file" in item:
            print("Datei:", item["name"])
    
def process_folder(rootId, folderId):
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{rootId}/items/{folderId}/children",
        headers=headers
    )
    data = response.json()["value"]
    for item in data:
        if "folder" in item:
            print("process folder", item["name"])
            process_folder(rootId, item["id"])
            
        elif "file" in item:
            print("process file", item["name"])
            #dprocess_file(item["id"])


def process_file(rootId, item):
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{rootId}/items/{item}",
        headers=headers
    )
    data = response.json()
    metaData = SharePointFile(
        id=data["id"],
        site_id=data["parentReference"]["siteId"],
        drive_id=data["parentReference"]["driveId"],
        name=data["name"],
        mime_type=data["file"]["mimeType"],
        size=data["size"],
        web_url=data["webUrl"],
        created_at=data["createdDateTime"],
        modified_at=data["lastModifiedDateTime"],
        parent_id=data["parentReference"]["id"]
    )
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{rootId}/items/{item}/content",
        headers=headers
    )

    text = ""

    match metaData.mime_type:
        case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            document = Document(BytesIO(response.content))
            paragraphs = []
            for paragraph in document.paragraphs:
                if paragraph.text.strip():
                    paragraphs.append(paragraph.text)
            text = "\n".join(paragraphs)

        case "application/pdf":
            document = PdfReader(BytesIO(response.content))
            text_parts = []
            for page in document.pages:
                page_text = page.extract_text()
                if page_text:
                    text_parts.append(page_text)
            text = "\n".join(text_parts)
    
        case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            document = load_workbook(BytesIO(response.content))
            text_parts = []
            for sheet in document.worksheets:
                text_parts.append(f"Sheet: {sheet.title}")
                rows = sheet.iter_rows(values_only=True)
                col_headers = next(rows, None)

                if not col_headers:
                    continue
                for row in rows:
                    values = []

                    for header, value in zip(col_headers,row):
                        if value is not None:
                            values.append(f"{header}: {value}")
                    if values:
                        text_parts.append(" | ".join(values))
                text_parts.append("")
            text = "\n".join(text_parts)

        case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
            document = Presentation(BytesIO(response.content))
            text_parts = []
            for slide_number, slide in enumerate(document.slides, start=1):
                text_parts.append(f"Slide: {slide_number}")

                for shape in slide.shapes:
                    if shape.has_text_frame:
                        shape_text = shape.text.strip()

                        if shape_text:
                            text_parts.append(shape_text)
                text_parts.append("")
            text = "\n".join(text_parts)
        case _:
            return
            

    contentData = DocumentText(
        file_id=metaData.id,
        filename=metaData.name,
        text=text
    )
    return metaData, contentData

process_root(it_ressort_drive_id)