import sqlite3

connection = sqlite3.connect("sharpoint.db")
curosr = connection.cursor()

curosr.execute("""
    CREATE TABLE IF NOT EXISTS chunks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id TEXT NOT NULL,
        chunk_index INTEGER NOT NULL,
        text TEXT NOT NULL,
        FOREIGN KEY (document_id) REFERENCES documents(id)
    )
""")

connection.commit()


def chunk_text(text, chunk_size=200, overlap=40):
    words = text.split()

    chunks=[]
    start = 0

    while start < len(words):
        end = start + chunk_size
        chunks.append(" ".join(words[start:end]))
        start += chunk_size - overlap

    return chunks

def process_documents():
    documents = curosr.execute("""
        SELECT id, text from documents
    """).fetchall()

    for document_id, text in documents:
        curosr.execute("""
            DELETE FROM chunks
            WHERE document_id = ?
        """, (document_id,))
            
        chunks = chunk_text(text)
        for i in range(len(chunks)):
            curosr.execute("""
            INSERT INTO chunks(
                document_id,
                chunk_index,
                text 
            )
            VALUES(?, ?, ?)
            """, (
                document_id,
                i,
                chunks[i]
            ))
        connection.commit()

process_documents()