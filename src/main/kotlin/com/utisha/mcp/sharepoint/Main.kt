package com.utisha.mcp.sharepoint

import com.fasterxml.jackson.databind.ObjectMapper
import io.modelcontextprotocol.json.jackson.JacksonMcpJsonMapper
import io.modelcontextprotocol.server.transport.StdioServerTransportProvider

/**
 * Entry point for the SharePoint MCP server.
 *
 * Reads configuration from environment variables and starts an MCP server
 * over stdio (stdin/stdout), compatible with any MCP client (Claude Desktop,
 * Cursor, Windsurf, etc.).
 *
 * ## Required Environment Variables
 *
 * - `SHAREPOINT_TENANT_ID` — Azure AD tenant ID
 * - `SHAREPOINT_CLIENT_ID` — Azure AD application (client) ID
 * - `SHAREPOINT_CLIENT_SECRET` — Azure AD client secret
 *
 * ## Optional Environment Variables
 *
 * - `SHAREPOINT_SITE_URL` — Restrict operations to a specific SharePoint site
 * - `SHAREPOINT_MAX_UPLOAD_BYTES` — Hard limit on upload file size (0 = unlimited)
 * - `SHAREPOINT_MAX_PAGINATION_RESULTS` — Safety cap for list operations (default: 500)
 */
fun main() {
  val config = SharePointConfig.fromEnv()

  val jsonMapper = JacksonMcpJsonMapper(ObjectMapper())
  val transportProvider = StdioServerTransportProvider(jsonMapper)
  val server = SharePointMcpServer.create(transportProvider, config)

  // Keep the server running until the process is terminated.
  // The MCP client controls the lifecycle by closing stdin or killing the process.
  Runtime.getRuntime().addShutdownHook(Thread {
    server.close()
  })

  Thread.currentThread().join()
}
