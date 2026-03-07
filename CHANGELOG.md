# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-03-07

### Added

- Initial release of mcp-server-sharepoint
- 10 MCP tools for SharePoint document operations:
  - `sharepoint_list_sites` — List or search accessible SharePoint sites
  - `sharepoint_list_drives` — List document libraries for a site
  - `sharepoint_list_files` — List files and folders in a drive path
  - `sharepoint_get_file_content` — Download file content as text (UTF-8)
  - `sharepoint_upload_file` — Upload file (base64-encoded, auto-chunked for >4MB)
  - `sharepoint_create_folder` — Create a folder
  - `sharepoint_delete_item` — Delete a file or folder
  - `sharepoint_search_files` — Search files in a drive
  - `sharepoint_copy_item` — Copy item to new location
  - `sharepoint_move_item` — Move item to new location
- Azure AD client credentials authentication (client_credentials flow)
- Automatic retry with exponential backoff for Graph API rate limits (429/503)
- Transparent pagination with configurable safety cap
- Auto-switching between simple upload (<4MB) and chunked resumable upload
- Structured error extraction from Graph API ODataError responses
- Optional site URL filtering via `SHAREPOINT_SITE_URL`
- Configurable upload size limits and pagination caps
- Runnable fat JAR with stdio transport (compatible with Claude Desktop, Cursor, etc.)
- Unit tests with MockK (30+ tests)
- Integration tests against real SharePoint tenants
