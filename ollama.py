from qdrant_client import QdrantClient
from qdrant_client.models import Distance, VectorParams, PointStruct
import sqlite3
import requests

connection = sqlite3.connect("sharpoint.db")
cursor = connection.cursor()

client = QdrantClient(url="http://localhost:6333")

if not client.collection_exists("wi3_docs"):
    client.create_collection(
        collection_name="wi3_docs",
        vectors_config=VectorParams(size=1024, distance=Distance.COSINE)
    )


def get_embedding(text, model="bge-m3"):
    response = requests.post(
        "http://localhost:11434/api/embeddings",
        json={"model": model, "prompt": text},
        timeout=60
    )
    response.raise_for_status()
    return response.json()["embedding"]

def embedd_chunks():
    chunks = cursor.execute("""SELECT * FROM chunks""").fetchall()

    BATCH_SIZE = 100
    batch = []
    for chunk_id, document_id, chunk_index, text, filename, web_url, mime_type, modified_at in chunks:
        chunk_embedding = get_embedding(text)
        point = PointStruct(id=chunk_id, vector=chunk_embedding, payload={"text": text, 
                                                                          "document_id": document_id, 
                                                                          "chunk_index": chunk_index, 
                                                                          "filename": filename, 
                                                                          "web_url": web_url,
                                                                          "mime_type": mime_type, 
                                                                          "modified_at": modified_at})
        batch.append(point)

        if len(batch) >= BATCH_SIZE:
            client.upsert(collection_name="wi3_docs", points=batch)
            batch = []
    if batch:        
        client.upsert(collection_name="wi3_docs", points=batch)

def search(query, limit=5):
    query_vector = get_embedding(query)
    results = client.query_points(
        collection_name="wi3_docs",
        query=query_vector,
        limit=limit
    )
    seen_filenames = set()
    for r in results.points:
        filename = r.payload["filename"]
        if filename in seen_filenames:
            continue
        seen_filenames.add(filename)
        print(f"Score: {r.score:.3f} | {r.payload['filename']}")
        print(r.payload['text'][:200])
        print("-----")

search("Aufgaben qm ressort")
connection.close()
