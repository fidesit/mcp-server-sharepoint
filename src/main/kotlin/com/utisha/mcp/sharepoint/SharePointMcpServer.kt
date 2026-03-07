package com.utisha.mcp.sharepoint

import com.fasterxml.jackson.databind.ObjectMapper
import com.fasterxml.jackson.databind.SerializationFeature
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule
import io.modelcontextprotocol.server.McpServer
import io.modelcontextprotocol.server.McpSyncServer
import io.modelcontextprotocol.spec.McpSchema
import io.modelcontextprotocol.spec.McpSchema.CallToolResult
import io.modelcontextprotocol.spec.McpSchema.TextContent
import io.modelcontextprotocol.spec.McpServerTransportProvider
import org.slf4j.LoggerFactory
import java.util.Base64

/**
 * Factory for creating an MCP server that exposes SharePoint operations as tools.
 *
 * The server is transport-agnostic — callers provide an [McpServerTransportProvider]
 * (e.g., [StdioServerTransportProvider] for CLI use, or a custom transport for embedding).
 *
 * Tools registered:
 * 1. sharepoint_list_sites — List/search SharePoint sites
 * 2. sharepoint_list_drives — List document libraries for a site
 * 3. sharepoint_list_files — List files/folders in a drive path
 * 4. sharepoint_get_file_content — Download file content as text
 * 5. sharepoint_upload_file — Upload file (base64-encoded content)
 * 6. sharepoint_create_folder — Create a folder
 * 7. sharepoint_delete_item — Delete a file or folder
 * 8. sharepoint_search_files — Search files in a drive
 * 9. sharepoint_copy_item — Copy item to new location
 * 10. sharepoint_move_item — Move item to new location
 */
object SharePointMcpServer {

  private val logger = LoggerFactory.getLogger(SharePointMcpServer::class.java)
  private val objectMapper = ObjectMapper()
    .registerModule(JavaTimeModule())
    .disable(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS)

  /**
   * Create and start an MCP server with SharePoint tools.
   *
   * @param transportProvider Transport provider for MCP communication
   * @param config SharePoint connection configuration
   * @return Running [McpSyncServer] instance
   */
  fun create(transportProvider: McpServerTransportProvider, config: SharePointConfig): McpSyncServer {
    val graphClient = SharePointGraphClient(config)

    val server = McpServer.sync(transportProvider)
      .serverInfo("mcp-server-sharepoint", "0.1.0")
      // ── List Sites ──
      .toolCall(
        buildTool("sharepoint_list_sites",
          "List or search SharePoint sites accessible by the configured Azure AD application. " +
            "This is typically the FIRST tool to call — use the returned site ID for subsequent operations. " +
            "Without a search query, returns all accessible sites. " +
            "Example: search for 'marketing' to find the marketing team site.",
          mapOf("search" to mapOf("type" to "string", "description" to "Optional search query to filter sites by name (e.g., 'finance', 'marketing'). Omit to list all accessible sites.")),
          emptyList()
        )
      ) { _, request ->
        handleToolCall("sharepoint_list_sites", request) { args ->
          toJson(graphClient.listSites(args.str("search")))
        }
      }
      // ── List Drives ──
      .toolCall(
        buildTool("sharepoint_list_drives",
          "List document libraries (drives) for a SharePoint site. " +
            "Requires a site ID from sharepoint_list_sites. " +
            "Most sites have a default 'Documents' drive. Use the returned drive ID for file operations.",
          mapOf("siteId" to mapOf("type" to "string", "description" to "SharePoint site ID (from sharepoint_list_sites, e.g., 'contoso.sharepoint.com,guid1,guid2')")),
          listOf("siteId")
        )
      ) { _, request ->
        handleToolCall("sharepoint_list_drives", request) { args ->
          toJson(graphClient.listDrives(args.require("siteId")))
        }
      }
      // ── List Files ──
      .toolCall(
        buildTool("sharepoint_list_files",
          "List files and folders in a document library drive. " +
            "Requires a drive ID from sharepoint_list_drives. " +
            "Optionally specify a folder path to list contents of a specific folder. " +
            "Returns file metadata including item IDs needed for download, delete, move, and copy operations.",
          mapOf(
            "driveId" to mapOf("type" to "string", "description" to "Drive ID (from sharepoint_list_drives)"),
            "path" to mapOf("type" to "string", "description" to "Optional folder path (e.g., 'Documents/Reports'). Omit or use '/' for root.")
          ),
          listOf("driveId")
        )
      ) { _, request ->
        handleToolCall("sharepoint_list_files", request) { args ->
          toJson(graphClient.listFiles(args.require("driveId"), args.str("path")))
        }
      }
      // ── Get File Content ──
      .toolCall(
        buildTool("sharepoint_get_file_content",
          "Download file content as text (UTF-8). Suitable for text-based files only (e.g., .txt, .csv, .json, .md). " +
            "Requires a drive ID and item ID from sharepoint_list_files. " +
            "For binary files (images, PDFs), this will return garbled content.",
          mapOf(
            "driveId" to mapOf("type" to "string", "description" to "Drive ID (from sharepoint_list_drives)"),
            "itemId" to mapOf("type" to "string", "description" to "File item ID (from sharepoint_list_files)")
          ),
          listOf("driveId", "itemId")
        )
      ) { _, request ->
        handleToolCall("sharepoint_get_file_content", request) { args ->
          graphClient.getFileContent(args.require("driveId"), args.require("itemId"))
        }
      }
      // ── Upload File ──
      .toolCall(
        buildTool("sharepoint_upload_file",
          "Upload a file to a document library drive. Content must be base64-encoded. " +
            "Automatically uses chunked resumable upload for files larger than ~4MB. " +
            "If a file with the same name exists, it will be overwritten.",
          mapOf(
            "driveId" to mapOf("type" to "string", "description" to "Drive ID (from sharepoint_list_drives)"),
            "parentPath" to mapOf("type" to "string", "description" to "Parent folder path (e.g., 'Documents/Reports'). Use '/' for root."),
            "fileName" to mapOf("type" to "string", "description" to "Name of the file to create (e.g., 'report.pdf')"),
            "content" to mapOf("type" to "string", "description" to "File content, base64-encoded")
          ),
          listOf("driveId", "parentPath", "fileName", "content")
        )
      ) { _, request ->
        handleToolCall("sharepoint_upload_file", request) { args ->
          val contentBytes = Base64.getDecoder().decode(args.require("content"))
          toJson(graphClient.uploadFile(args.require("driveId"), args.require("parentPath"), args.require("fileName"), contentBytes))
        }
      }
      // ── Create Folder ──
      .toolCall(
        buildTool("sharepoint_create_folder",
          "Create a new folder in a document library drive at the specified parent path. " +
            "Use '/' as parentPath to create a folder at the root of the drive.",
          mapOf(
            "driveId" to mapOf("type" to "string", "description" to "Drive ID (from sharepoint_list_drives)"),
            "parentPath" to mapOf("type" to "string", "description" to "Parent folder path (use '/' for root)"),
            "folderName" to mapOf("type" to "string", "description" to "Name of the folder to create")
          ),
          listOf("driveId", "parentPath", "folderName")
        )
      ) { _, request ->
        handleToolCall("sharepoint_create_folder", request) { args ->
          toJson(graphClient.createFolder(args.require("driveId"), args.require("parentPath"), args.require("folderName")))
        }
      }
      // ── Delete Item ──
      .toolCall(
        buildTool("sharepoint_delete_item",
          "Permanently delete a file or folder by item ID. This action cannot be undone. " +
            "Requires a drive ID and item ID from sharepoint_list_files. " +
            "Deleting a folder also deletes all its contents.",
          mapOf(
            "driveId" to mapOf("type" to "string", "description" to "Drive ID (from sharepoint_list_drives)"),
            "itemId" to mapOf("type" to "string", "description" to "Item ID of the file or folder to delete (from sharepoint_list_files)")
          ),
          listOf("driveId", "itemId")
        )
      ) { _, request ->
        handleToolCall("sharepoint_delete_item", request) { args ->
          graphClient.deleteItem(args.require("driveId"), args.require("itemId"))
          """{"success": true, "message": "Item deleted"}"""
        }
      }
      // ── Search Files ──
      .toolCall(
        buildTool("sharepoint_search_files",
          "Search for files within a document library drive using a text query. " +
            "Searches file names and content. Returns matching files with their item IDs. " +
            "Requires a drive ID from sharepoint_list_drives.",
          mapOf(
            "driveId" to mapOf("type" to "string", "description" to "Drive ID (from sharepoint_list_drives)"),
            "query" to mapOf("type" to "string", "description" to "Search query (e.g., 'quarterly report', 'budget 2024')")
          ),
          listOf("driveId", "query")
        )
      ) { _, request ->
        handleToolCall("sharepoint_search_files", request) { args ->
          toJson(graphClient.searchFiles(args.require("driveId"), args.require("query")))
        }
      }
      // ── Copy Item ──
      .toolCall(
        buildTool("sharepoint_copy_item",
          "Copy a file or folder to a new location within the same drive. " +
            "The copy operation is asynchronous server-side — the response may indicate 'copy in progress'. " +
            "Requires item ID from sharepoint_list_files and a destination folder path.",
          mapOf(
            "driveId" to mapOf("type" to "string", "description" to "Drive ID (from sharepoint_list_drives)"),
            "itemId" to mapOf("type" to "string", "description" to "Item ID to copy (from sharepoint_list_files)"),
            "destinationPath" to mapOf("type" to "string", "description" to "Destination folder path (e.g., 'Archive/2024')")
          ),
          listOf("driveId", "itemId", "destinationPath")
        )
      ) { _, request ->
        handleToolCall("sharepoint_copy_item", request) { args ->
          toJson(graphClient.copyItem(args.require("driveId"), args.require("itemId"), args.require("destinationPath")))
        }
      }
      // ── Move Item ──
      .toolCall(
        buildTool("sharepoint_move_item",
          "Move a file or folder to a new location within the same drive. " +
            "The item is removed from its current location. " +
            "Requires item ID from sharepoint_list_files and a destination folder path.",
          mapOf(
            "driveId" to mapOf("type" to "string", "description" to "Drive ID (from sharepoint_list_drives)"),
            "itemId" to mapOf("type" to "string", "description" to "Item ID to move (from sharepoint_list_files)"),
            "destinationPath" to mapOf("type" to "string", "description" to "Destination folder path (e.g., 'Archive/2024')")
          ),
          listOf("driveId", "itemId", "destinationPath")
        )
      ) { _, request ->
        handleToolCall("sharepoint_move_item", request) { args ->
          toJson(graphClient.moveItem(args.require("driveId"), args.require("itemId"), args.require("destinationPath")))
        }
      }
      .build()

    logger.info("SharePoint MCP server created with 10 tools")
    return server
  }

  // ── Helpers ───────────────────────────────────────────────────────────

  private fun buildTool(
    name: String,
    description: String,
    properties: Map<String, Map<String, String>>,
    required: List<String>
  ): McpSchema.Tool {
    return McpSchema.Tool.builder()
      .name(name)
      .description(description)
      .inputSchema(McpSchema.JsonSchema(
        "object",
        properties.mapValues { (_, v) -> v as Any },
        required,
        false,
        null,
        null
      ))
      .build()
  }

  private fun handleToolCall(toolName: String, request: McpSchema.CallToolRequest, block: (Map<String, Any>) -> String): CallToolResult {
    @Suppress("UNCHECKED_CAST")
    val args = request.arguments() as? Map<String, Any> ?: emptyMap()
    return try {
      val result = block(args)
      CallToolResult.builder()
        .content(listOf(TextContent(result)))
        .isError(false)
        .build()
    } catch (e: Exception) {
      logger.error("Tool '{}' failed: {}", toolName, e.message, e)
      CallToolResult.builder()
        .content(listOf(TextContent("Error: ${e.message}")))
        .isError(true)
        .build()
    }
  }

  private fun Map<String, Any>.str(key: String): String? {
    return this[key]?.toString()?.takeIf { it.isNotBlank() }
  }

  private fun Map<String, Any>.require(key: String): String {
    return this[key]?.toString()
      ?: throw IllegalArgumentException("Required parameter '$key' is missing")
  }

  private fun toJson(value: Any): String {
    return objectMapper.writeValueAsString(value)
  }
}
