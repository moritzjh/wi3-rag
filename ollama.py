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
        vector_config=VectorParams(size=1024, distance=Distance.COSINE)
    )


def get_embedding(text, model="bge-m3"):
    response = requests.post(
        "http://localhost:11434/api/embeddings",
        json={"model": model, "prompt": text}
    )
    response.raise_for_status()
    response.json()["embedding"]

def embedd_chunks():
    chunks = cursor.execute("""SELECT * FROM chunks""").fetchall()

    BATCH_SIZE = 100
    batch = []
    for row in chunks:
        



