import sqlite3

connection = sqlite3.connect("sharpoint.db")
cursor = connection.cursor()

cursor.execute("""
    CREATE TABLE IF NOT EXISTS chunks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id TEXT NOT NULL,
        chunk_index INTEGER NOT NULL,
        text TEXT NOT NULL,
        filename TEXT,
        web_url TEXT,
        mime_type TEXT,
        modified_at TEXT,
        FOREIGN KEY (document_id) REFERENCES documents(id)
    )
""")

connection.commit()


def chunk_text(text, chunk_size=100, overlap=40):
    words = text.split()

    chunks=[]
    start = 0

    while start < len(words):
        end = start + chunk_size
        chunks.append(" ".join(words[start:end]))
        start += chunk_size - overlap

    return chunks

def process_documents():
    documents = cursor.execute("""
        SELECT id, text, name, web_url, mime_type, modified_at from documents
    """).fetchall()

    for document_id, text, name, web_url, mime_type, modified_at in documents:
        cursor.execute("""
            DELETE FROM chunks
            WHERE document_id = ?
        """, (document_id,))
            
        chunks = chunk_text(text)
        for i in range(len(chunks)):
            cursor.execute("""
            INSERT INTO chunks(
                document_id,
                chunk_index,
                text,
                filename,
                web_url,
                mime_type,
                modified_at 
            )
            VALUES(?, ?, ?, ?, ?, ?, ?)
            """, (
                document_id,
                i,
                chunks[i],
                name,
                web_url,
                mime_type,
                modified_at
            ))
        connection.commit()
process_documents()

connection.close()
