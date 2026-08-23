import os
import requests
import jwt
from jwt import PyJWKClient
from dotenv import load_dotenv
from pydantic import AnyHttpUrl
from mcp.server import MCPServer
from mcp.server.auth.provider import AccesToken, TokenVerifier
from mcp.server.auth.middleware.auth_context import get_access_token
from mcp.server.auth.settings import AuthSettings
from qdrant_client import QdrantClient

load_dotenv()
TENANT_ID = os.getenv("TENANT_ID")
LOGIN_CLIENT_ID = os.getenv("LOGIN_CLIENT_ID")

JWKS_URL = f"https://login.microsoftonline.com/{TENANT_ID}/discovery/v2.0/keys"
jwks_client = PyJWKClient(JWKS_URL)

class AzureADTokenVerifier(TokenVerifier):
    async def verify_token(self, token: str) -> AccesToken | None:
        try:
            signing_key =jwks_client.get_signing_key_from_jwt(token)
            claims = jwt.decode(
                token,
                signing_key.key,
                algorithms=["RS256"],
                audience=LOGIN_CLIENT_ID,
                issuer=f"https://login.microsoftonline.com/{TENANT_ID}/v2.0"
            )
            return AccesToken(
                token=token,
                client_id=claims.get("preferred_username"),
                scopes=[],
                claims=claims
            )
        except Exception:
            return None

mcp = MCPServer(
    "wi3-knowledge",
    token_verifier=AzureADTokenVerifier(),
    auth=AuthSettings(
        issuer_url=AnyHttpUrl(f"https://login.microsoftonline.com/{TENANT_ID}/v2.0"),
        ressurce_server_url=AnyHttpUrl("https://mcp.wi3-consulting.de/mcp"),
        required_scopes=[]
    )
)

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
    access_token = get_access_token()
    if access_token is None:
        return "Nicht angemeldet"

    email = access_token.claims.get("preferred_username", "")
    if not email.endswith("@wi3-consulting.de"):
        return "Zugriff verweigert: keine wi3-consultin E-Mail"

    
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
    mcp.run(transport="streamable-http", host="0.0.0.0", port=8000)
