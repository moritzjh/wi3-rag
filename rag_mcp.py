from mcp.server.mcpserver import MCPServer
from qdrant_client import QdrantClient
import requests

mcp = MCPServer("wi3-knowledge")
client = QdrantClient(url="http://localhost:6333")

def get_embedding(text, model="bge-m3"):
    response = requests.post(
        "http://localhost:11434/api/embeddings",
        json={"model": model, "prompt": text},
        timeout=60
    )
    response.raise_for_status()
    return response.json()["embedding"]

@mcp.tool()
def search_wi3_knowledge(query: str) -> str:
    """Durchsucht die wi3 consulting Wissensdatenbank (Sharepoint-Dokumente) nach relevanten Informationen zu einer Frage."""
    query_vector = get_embedding(query)
    results = client.query_points(
        collection_name="wi3_docs",
        query=query_vector,
        limit=5
    )

    if not results:
        return "Keine relevanten Dokumente gefunden"

    parts = []
    seen_filenames = set()
    for r in results.points:
        filename = r.payload["filename"]
        if filename in seen_filenames:
            continue
        seen_filenames.add(filename)
        parts.append(f"[{r.payload['filename']}]\n{r.payload['text']}")
    return "\n\n---\n\n".join(parts)

if __name__ == "__main__":
    mcp.run()
