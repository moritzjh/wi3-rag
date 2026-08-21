# wi3-rag

A small retrieval-augmented-generation pipeline that crawls SharePoint documents via
Microsoft Graph, chunks and embeds them, and exposes semantic search over the result
through an MCP server.

## Pipeline

1. **`graph_sharepoint.py`** – Authenticates against Microsoft Graph, recursively
   walks a fixed list of SharePoint sites/drives, extracts text from supported
   documents (`.docx`, `.pdf`, `.xlsx`, `.pptx`), and stores each document in the
   `documents` table of a local SQLite database (`sharpoint.db`). Skips documents
   whose `lastModifiedDateTime` hasn't changed since the last run. At the end it
   calls into `chunking.py`.
2. **`chunking.py`** – Splits each document's text into overlapping word chunks and
   stores them in the `chunks` table of `sharpoint.db`.
3. **`ollama.py`** – Embeds each chunk with a local Ollama model (`bge-m3` by
   default) and upserts the vectors into a Qdrant collection (`wi3_docs`). Also
   contains a simple `search()` helper for ad-hoc querying from the CLI.
4. **`rag_mcp.py`** – An MCP server that exposes a `search_wi3_knowledge` tool,
   letting an MCP-compatible client (e.g. Claude) query the embedded knowledge base
   in Qdrant.

`test_login.py` is a standalone script for interactively testing the delegated
(user) login flow against Azure AD; it's not part of the main pipeline.

## Requirements

- Python 3.13+
- A running [Ollama](https://ollama.com) instance with the embedding model pulled
  (`ollama pull bge-m3`)
- A running [Qdrant](https://qdrant.tech) instance (defaults to
  `http://localhost:6333`, e.g. via `docker run -p 6333:6333 qdrant/qdrant`)
- An Azure AD app registration with Microsoft Graph access to the target SharePoint
  sites

## Setup

```bash
python -m venv venv
venv\Scripts\activate        # on Windows
pip install msal requests python-dotenv python-docx pypdf openpyxl python-pptx qdrant-client mcp

cp .env.example .env         # then fill in your Azure AD values
```

## Usage

```bash
# 1. Crawl SharePoint, extract text, and chunk it
python graph_sharepoint.py

# 2. Embed the chunks and upsert them into Qdrant (also runs a sample search)
python ollama.py

# 3. Start the MCP server
python rag_mcp.py
```

## Notes

- `sharpoint.db` (the local SQLite database) and `__pycache__/` are git-ignored and
  generated locally — they are not part of the repository.
- The list of SharePoint site IDs to crawl is currently hard-coded in
  `graph_sharepoint.py`.
