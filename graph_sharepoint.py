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
import sqlite3

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
    text: str

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")


AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

connection = sqlite3.connect("sharpoint.db")

cursor = connection.cursor()

cursor.execute("""
    CREATE TABLE IF NOT EXISTS documents (
        id TEXT PRIMARY KEY,
        site_id TEXT NOT NULL,
        drive_id TEXT NOT NULL,
        name TEXT NOT NULL,
        mime_type TEXT,
        size INTEGER,
        web_url TEXT,
        created_at TEXT,
        modified_at TEXT,
        parent_id TEXT,
        text TEXT
    )
""")

connection.commit()



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
            process_file(rootId, item["id"])
    
def process_folder(rootId, folderId):
    meta_Data = []
    docuemtn_Content = []
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{rootId}/items/{folderId}/children",
        headers=headers
    )
    data = response.json()["value"]
    for item in data:
        if "folder" in item:
            print("process folder", item["name"])
            sub_meta, sub_content = process_folder(rootId, item["id"])

            meta_Data.extend(sub_meta)
            docuemtn_Content.extend(sub_content)
            
        elif "file" in item:
            print("process file", item["name"])
            result = process_file(rootId, item["id"])
            if result:
                meta, content = result
                meta_Data.append(meta)
                docuemtn_Content.append(content)

    return meta_Data, docuemtn_Content


def process_file(rootId, item):
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{rootId}/items/{item}",
        headers=headers
    )

    metaData = response.json()

    existing = cursor.execute("""SELECT modified_at FROM documents WHERE id=?""", (metaData["id"],)).fetchone()
    if existing and existing[0] == metaData["lastModifiedDateTime"]:
        return
    
    response = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{it_ressort_site_id}/drives/{rootId}/items/{item}/content",
    headers=headers
    )
    
    text = ""

    match metaData["file"]["mimeType"]:
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
        
    cursor.execute("""
        INSERT INTO documents(
            id,
            site_id,
            drive_id,
            name,
            mime_type,
            size,
            web_url,
            created_at,
            modified_at,
            parent_id,
            text
        )
        VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)

        ON CONFLICT(id) DO UPDATE SET
        site_id=excluded.site_id,
        drive_id=excluded.drive_id,
        name=excluded.name,
        mime_type=excluded.mime_type,
        size=excluded.size,
        web_url=excluded.web_url,
        created_at=excluded.created_at,
        modified_at=excluded.modified_at,
        parent_id=excluded.parent_id,
        text=excluded.text
    """, (
        metaData["id"],
        metaData["parentReference"]["siteId"],
        metaData["parentReference"]["driveId"],
        metaData["name"],
        metaData["file"]["mimeType"],
        metaData["size"],
        metaData["webUrl"],
        metaData["createdDateTime"],
        metaData["lastModifiedDateTime"],
        metaData["parentReference"]["id"],
        text
    ))
    connection.commit()
    return
'''
meta, content = process_folder(it_ressort_drive_id, "01HODMHCXI62P3O77VIRCZ7O5ZZXHFUADK")
for i in range(len(meta)):
    print(meta[i].name)
    print(content[i].text)
'''
file = process_file(it_ressort_drive_id, "01HODMHCWHUWZRJC72YNF2WTWTRQIOKKTJ")



connection.close()