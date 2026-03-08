# mcp-server-sharepoint

[![CI](https://github.com/fidesit/mcp-server-sharepoint/actions/workflows/ci.yml/badge.svg)](https://github.com/fidesit/mcp-server-sharepoint/actions/workflows/ci.yml)
[![Maven Central](https://img.shields.io/maven-central/v/com.utisha/mcp-server-sharepoint.svg)](https://central.sonatype.com/artifact/com.utisha/mcp-server-sharepoint)
[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://www.apache.org/licenses/LICENSE-2.0.txt)
[![Java 21+](https://img.shields.io/badge/Java-21%2B-orange.svg)](https://openjdk.org/)
[![Kotlin](https://img.shields.io/badge/Kotlin-2.2.0-purple.svg)](https://kotlinlang.org/)

Talk to your SharePoint — browse sites, manage documents, search files, all through natural language via any MCP-compatible AI assistant.

Built with Kotlin, Microsoft Graph SDK v6, and the official MCP SDK.

## Highlights

* 🛠️ 10 tools covering full document lifecycle
* 🔄 Automatic retry with exponential backoff (429/503)
* 📄 Transparent pagination (follows @odata.nextLink automatically)
* 📤 Chunked uploads for large files (>4MB auto-switches to resumable)
* 🔍 Structured error messages from Graph API
* 🤖 Works with Claude Desktop, Cursor, and any MCP client
* 📦 Published to Maven Central

## Quick Start

**Prerequisites:** JDK 21+, an Azure AD app registration with SharePoint permissions.

```bash
# 1. Clone and build
git clone https://github.com/fidesit/mcp-server-sharepoint.git
cd mcp-server-sharepoint
./gradlew build

# 2. Configure (copy and fill in your Azure AD credentials)
cp .env.example .env

# 3. Run
java -jar build/libs/mcp-server-sharepoint-0.1.0.jar
```

The server communicates over stdio (stdin/stdout) using JSON-RPC, as defined by the MCP specification. Connect it to any MCP-compatible client.

## Usage with MCP Clients

### Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "java",
      "args": ["-jar", "/absolute/path/to/mcp-server-sharepoint-0.1.0.jar"],
      "env": {
        "SHAREPOINT_TENANT_ID": "your-tenant-id",
        "SHAREPOINT_CLIENT_ID": "your-client-id",
        "SHAREPOINT_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### Cursor

Add to your Cursor MCP settings (`.cursor/mcp.json`):

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "java",
      "args": ["-jar", "/absolute/path/to/mcp-server-sharepoint-0.1.0.jar"],
      "env": {
        "SHAREPOINT_TENANT_ID": "your-tenant-id",
        "SHAREPOINT_CLIENT_ID": "your-client-id",
        "SHAREPOINT_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### Claude Code

```bash
claude mcp add sharepoint -- java -jar /path/to/mcp-server-sharepoint-0.1.0.jar
```

### Other Clients

Any MCP client that supports the stdio transport can use this server. Point it at the JAR with the required environment variables.

## Available Tools

### Navigation

| Tool | Description | Required Params | Optional Params |
|------|-------------|-----------------|-----------------|
| `sharepoint_list_sites` | List or search accessible SharePoint sites | — | `search` |
| `sharepoint_list_drives` | List document libraries for a site | `siteId` | — |
| `sharepoint_list_files` | List files and folders in a drive path | `driveId` | `path` |

### Read & Search

| Tool | Description | Required Params | Optional Params |
|------|-------------|-----------------|-----------------|
| `sharepoint_get_file_content` | Download file content as UTF-8 text | `driveId`, `itemId` | — |
| `sharepoint_search_files` | Search files in a drive by text query | `driveId`, `query` | — |

### Write & Organize

| Tool | Description | Required Params | Optional Params |
|------|-------------|-----------------|-----------------|
| `sharepoint_upload_file` | Upload a file (base64-encoded content) | `driveId`, `parentPath`, `fileName`, `content` | — |
| `sharepoint_create_folder` | Create a folder | `driveId`, `parentPath`, `folderName` | — |
| `sharepoint_copy_item` | Copy item to a new location | `driveId`, `itemId`, `destinationPath` | — |
| `sharepoint_move_item` | Move item to a new location | `driveId`, `itemId`, `destinationPath` | — |
| `sharepoint_delete_item` | Delete a file or folder (permanent) | `driveId`, `itemId` | — |

## Example Workflow

**User:** Find the marketing team's SharePoint site.
**Assistant:** *Calls `sharepoint_list_sites(search="marketing")`*

**User:** What document libraries does it have?
**Assistant:** *Calls `sharepoint_list_drives(siteId="...")`*

**User:** Show me the files in the Reports folder.
**Assistant:** *Calls `sharepoint_list_files(driveId="...", path="Reports")`*

**User:** Read the Q4 summary.
**Assistant:** *Calls `sharepoint_get_file_content(driveId="...", itemId="...")`*

## Configuration

### Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `SHAREPOINT_TENANT_ID` | Yes | — | Azure AD tenant ID |
| `SHAREPOINT_CLIENT_ID` | Yes | — | Azure AD application (client) ID |
| `SHAREPOINT_CLIENT_SECRET` | Yes | — | Azure AD client secret |
| `SHAREPOINT_SITE_URL` | No | — | Restrict operations to a specific SharePoint site URL |
| `SHAREPOINT_MAX_UPLOAD_BYTES` | No | `0` | Hard limit on upload file size in bytes. 0 = unlimited |
| `SHAREPOINT_MAX_PAGINATION_RESULTS` | No | `500` | Safety cap for paginated list operations |

### Azure AD Setup

1. Go to [Azure Portal](https://portal.azure.com/) > **App registrations** > **New registration**
2. Name it (e.g., `mcp-sharepoint-server`), select **Single tenant**, register
3. Note the **Application (client) ID** and **Directory (tenant) ID**
4. Go to **Certificates & secrets** > **New client secret** — copy the secret value
5. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Application permissions**, add:
   - `Sites.Read.All`
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
6. Click **Grant admin consent** (requires Azure AD admin)

## Building from Source

```bash
# Build and run unit tests
./gradlew build

# Run unit tests only
./gradlew test

# Run integration tests (requires SharePoint credentials in environment)
./gradlew integrationTest

# Build fat JAR only (skip tests)
./gradlew shadowJar
```

The fat JAR is produced at `build/libs/mcp-server-sharepoint-<version>.jar`.

## Project Structure

```
src/main/kotlin/com/utisha/mcp/sharepoint/
  Main.kt                  # Entry point — stdio transport, env-based config
  SharePointConfig.kt      # Configuration with validation and env var parsing
  SharePointGraphClient.kt # Microsoft Graph API client (all SharePoint operations)
  SharePointMcpServer.kt   # MCP tool definitions and request handling
```

## Architecture

The server is built in three layers:

1. **SharePointGraphClient** — Wraps Microsoft Graph SDK v6 (Kiota-based). Handles authentication, pagination, retries, and chunked uploads. This is a pure SharePoint client with no MCP dependency.

2. **SharePointMcpServer** — Registers 10 MCP tools and maps incoming tool calls to `SharePointGraphClient` methods. Transport-agnostic (receives an `McpServerTransportProvider`).

3. **Main** — Wires stdio transport to the server with environment-based configuration.

Key behaviors:
- **Automatic retry** with exponential backoff for Graph API throttling (429) and transient errors (503)
- **Transparent pagination** — list operations follow `@odata.nextLink` automatically, capped by `maxPaginationResults`
- **Chunked upload** — files larger than 4MB are uploaded using Graph's resumable upload session
- **Structured errors** — Graph API `ODataError` responses are extracted into readable messages

## License

[Apache License 2.0](LICENSE)
