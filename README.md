# Word MCP Server

A FastAPI-based Word document generation server with full MCP (Model Context Protocol) support.

## Features

- **REST API** for document creation and management
- **MCP Protocol Support** for integration with MCP-compatible clients
- **Markdown Support** in documents
- **Multiple Templates** (standard, report, memo, letter)
- **Chat Export** functionality to convert conversations to Word documents
- **Health Check** endpoint for monitoring

## API Endpoints

### Health Check
- `GET /health` - Service health status

### Document Operations
- `POST /word/create` - Create a new Word document
- `POST /word/create-from-chat` - Create a document from chat messages
- `GET /word/list` - List all documents
- `GET /word/download/{filename}` - Download a document
- `DELETE /word/delete/{filename}` - Delete a document

### MCP Protocol
- `POST /mcp` - MCP JSON-RPC 2.0 endpoint

### Service Info
- `GET /` - Root endpoint with service information
- `GET /tools` - Get available tools

## Deployment

### Docker Compose

```bash
cd word-mcp-server
docker-compose build
docker-compose up -d
```

The service will be available at `https://word.mac.mndambuki.me.ke` via Traefik.

### Local Development

```bash
pip install -r requirements.txt
python server.py
```

The server will run on `http://localhost:8004`

## MCP Protocol Support

The server implements the MCP JSON-RPC 2.0 protocol with the following methods:

- `initialize` - Initialize the MCP connection
- `tools/list` - List available tools
- `tools/call` - Call a tool
- `resources/list` - List available resources
- `resources/read` - Read a resource

## Example Requests

### Create a Document (REST)

```bash
curl -X POST http://localhost:8004/word/create \
  -H "Content-Type: application/json" \
  -d '{
    "title": "My Document",
    "content": "# Heading\n\nThis is a paragraph.",
    "author": "John Doe",
    "template": "standard"
  }'
```

### Create a Document (MCP Protocol)

```bash
curl -X POST http://localhost:8004/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "method": "tools/call",
    "params": {
      "name": "create_document",
      "arguments": {
        "title": "My Document",
        "content": "# Heading\n\nThis is a paragraph."
      }
    },
    "id": 1
  }'
```

## Environment Variables

- `TZ` - Timezone (default: Africa/Nairobi)
- `PORT` - Server port (default: 8004)

## Requirements

- Python 3.11+
- Docker (for containerized deployment)
- External Traefik network (for Docker Compose)

## Configuration

Documents are stored in `/app/documents` inside the container. In Docker Compose, this can be mounted to a local directory.

## License

MIT
