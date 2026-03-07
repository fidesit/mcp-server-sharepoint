package com.utisha.mcp.sharepoint

import com.azure.identity.ClientSecretCredentialBuilder
import com.microsoft.graph.core.models.IProgressCallback
import com.microsoft.graph.core.tasks.LargeFileUploadTask
import com.microsoft.graph.drives.item.items.item.copy.CopyPostRequestBody
import com.microsoft.graph.drives.item.items.item.createuploadsession.CreateUploadSessionPostRequestBody
import com.microsoft.graph.models.DriveItem
import com.microsoft.graph.models.DriveItemUploadableProperties
import com.microsoft.graph.models.Folder
import com.microsoft.graph.models.ItemReference
import com.microsoft.graph.models.odataerrors.ODataError
import com.microsoft.graph.serviceclient.GraphServiceClient
import org.slf4j.LoggerFactory
import java.io.ByteArrayInputStream
import java.net.URLEncoder
import java.nio.charset.StandardCharsets
import java.time.OffsetDateTime

/**
 * Wraps Microsoft Graph SDK for SharePoint document operations.
 *
 * Uses Azure AD client_credentials flow via [ClientSecretCredentialBuilder].
 * All operations target SharePoint sites and their document libraries (drives).
 *
 * Features:
 * - Automatic retry with exponential backoff (429/503 handling via SDK middleware)
 * - Transparent pagination for list operations (capped by [SharePointConfig.maxPaginationResults])
 * - Auto-switching between simple upload (<4MB) and chunked resumable upload
 * - Structured error extraction from Graph API ODataError responses
 * - Optional site filtering via [SharePointConfig.siteUrl]
 *
 * Required Azure AD API permissions (Application):
 * - Sites.Read.All — list sites, drives, files
 * - Sites.ReadWrite.All — create folders, upload/delete/move/copy files
 * - Files.ReadWrite.All — file content operations
 */
class SharePointGraphClient(private val config: SharePointConfig) {

  private val logger = LoggerFactory.getLogger(SharePointGraphClient::class.java)

  private val graphClient: GraphServiceClient by lazy {
    val credential = ClientSecretCredentialBuilder()
      .tenantId(config.tenantId)
      .clientId(config.clientId)
      .clientSecret(config.clientSecret)
      .build()

    // SDK v6 default constructor includes retry middleware (RetryHandler) for 429/503.
    GraphServiceClient(credential, "https://graph.microsoft.com/.default")
  }

  // ── Data classes ──────────────────────────────────────────────────────

  data class SiteInfo(
    val id: String,
    val name: String,
    val displayName: String,
    val webUrl: String,
    val description: String?
  )

  data class DriveInfo(
    val id: String,
    val name: String,
    val driveType: String?,
    val webUrl: String?,
    val totalSize: Long?,
    val usedSize: Long?
  )

  data class FileInfo(
    val id: String,
    val name: String,
    val webUrl: String?,
    val size: Long?,
    val isFolder: Boolean,
    val lastModified: OffsetDateTime?,
    val createdBy: String?,
    val mimeType: String?
  )

  // ── Operations ────────────────────────────────────────────────────────

  /**
   * List or search SharePoint sites accessible by the app.
   *
   * When [search] is provided, filters sites by name. Otherwise, returns all accessible sites.
   * If [SharePointConfig.siteUrl] is configured, results are filtered to only include
   * sites matching that URL.
   * Results are paginated transparently up to [SharePointConfig.maxPaginationResults].
   */
  fun listSites(search: String? = null): List<SiteInfo> {
    return try {
      val sites = mutableListOf<SiteInfo>()

      val firstPage = if (!search.isNullOrBlank()) {
        graphClient.sites().get { cfg ->
          cfg.queryParameters.search = search
        }
      } else {
        graphClient.sites().get()
      }

      firstPage?.value?.forEach { site ->
        if (sites.size >= config.maxPaginationResults) return@forEach
        sites.add(site.toSiteInfo())
      }

      // Follow pagination links
      var nextLink = firstPage?.odataNextLink
      while (nextLink != null && sites.size < config.maxPaginationResults) {
        val nextPage = graphClient.sites().withUrl(nextLink).get()
        nextPage?.value?.forEach { site ->
          if (sites.size >= config.maxPaginationResults) return@forEach
          sites.add(site.toSiteInfo())
        }
        nextLink = nextPage?.odataNextLink
      }

      // Apply siteUrl filter if configured
      if (!config.siteUrl.isNullOrBlank()) {
        sites.filter { it.webUrl.equals(config.siteUrl, ignoreCase = true) }
      } else {
        sites
      }
    } catch (e: Exception) {
      throw graphException("list sites", e, "search" to search)
    }
  }

  /**
   * List document libraries (drives) for a site.
   *
   * Results are paginated transparently up to [SharePointConfig.maxPaginationResults].
   */
  fun listDrives(siteId: String): List<DriveInfo> {
    return try {
      val drives = mutableListOf<DriveInfo>()

      val firstPage = graphClient.sites().bySiteId(siteId).drives().get()
      firstPage?.value?.forEach { drive ->
        if (drives.size >= config.maxPaginationResults) return@forEach
        drives.add(DriveInfo(
          id = drive.id ?: "",
          name = drive.name ?: "",
          driveType = drive.driveType,
          webUrl = drive.webUrl,
          totalSize = drive.quota?.total,
          usedSize = drive.quota?.used
        ))
      }

      var nextLink = firstPage?.odataNextLink
      while (nextLink != null && drives.size < config.maxPaginationResults) {
        val nextPage = graphClient.sites().bySiteId(siteId).drives().withUrl(nextLink).get()
        nextPage?.value?.forEach { drive ->
          if (drives.size >= config.maxPaginationResults) return@forEach
          drives.add(DriveInfo(
            id = drive.id ?: "",
            name = drive.name ?: "",
            driveType = drive.driveType,
            webUrl = drive.webUrl,
            totalSize = drive.quota?.total,
            usedSize = drive.quota?.used
          ))
        }
        nextLink = nextPage?.odataNextLink
      }

      drives
    } catch (e: Exception) {
      throw graphException("list drives", e, "siteId" to siteId)
    }
  }

  /**
   * List files and folders in a drive, optionally at a specific path.
   *
   * Results are paginated transparently up to [SharePointConfig.maxPaginationResults].
   */
  fun listFiles(driveId: String, path: String? = null): List<FileInfo> {
    return try {
      val parentId = resolveItemId(driveId, path)
      val files = mutableListOf<FileInfo>()

      val firstPage = graphClient.drives().byDriveId(driveId)
        .items().byDriveItemId(parentId)
        .children()
        .get()

      firstPage?.value?.forEach { item ->
        if (files.size >= config.maxPaginationResults) return@forEach
        files.add(item.toFileInfo())
      }

      var nextLink = firstPage?.odataNextLink
      while (nextLink != null && files.size < config.maxPaginationResults) {
        val nextPage = graphClient.drives().byDriveId(driveId)
          .items().byDriveItemId(parentId)
          .children().withUrl(nextLink)
          .get()
        nextPage?.value?.forEach { item ->
          if (files.size >= config.maxPaginationResults) return@forEach
          files.add(item.toFileInfo())
        }
        nextLink = nextPage?.odataNextLink
      }

      files
    } catch (e: Exception) {
      throw graphException("list files", e, "driveId" to driveId, "path" to path)
    }
  }

  /**
   * Download file content as text (UTF-8).
   *
   * Note: Only suitable for text-based files. Binary files should use a different approach.
   */
  fun getFileContent(driveId: String, itemId: String): String {
    return try {
      val stream = graphClient.drives().byDriveId(driveId)
        .items().byDriveItemId(itemId)
        .content()
        .get()

      stream?.bufferedReader()?.use { it.readText() }
        ?: throw RuntimeException("No content returned for item $itemId")
    } catch (e: Exception) {
      throw graphException("get file content", e, "driveId" to driveId, "itemId" to itemId)
    }
  }

  /**
   * Upload a file to a drive at the specified path.
   *
   * Automatically selects the upload strategy based on file size:
   * - Files <= [SharePointConfig.maxSimpleUploadBytes]: simple single-request upload
   * - Files > threshold: chunked resumable upload via [LargeFileUploadTask]
   *
   * Chunk size for large uploads is configured via [SharePointConfig.uploadChunkBytes].
   */
  fun uploadFile(driveId: String, parentPath: String, fileName: String, content: ByteArray): FileInfo {
    return try {
      if (config.maxUploadBytes > 0 && content.size > config.maxUploadBytes) {
        throw IllegalArgumentException(
          "File size ${content.size} bytes exceeds maximum allowed ${config.maxUploadBytes} bytes"
        )
      }
      if (content.size <= config.maxSimpleUploadBytes) {
        simpleUpload(driveId, parentPath, fileName, content)
      } else {
        chunkedUpload(driveId, parentPath, fileName, content)
      }
    } catch (e: Exception) {
      throw graphException("upload file", e, "driveId" to driveId, "path" to "$parentPath/$fileName", "size" to content.size)
    }
  }

  /**
   * Create a folder in a drive at the specified parent path.
   */
  fun createFolder(driveId: String, parentPath: String, folderName: String): FileInfo {
    return try {
      val parentId = resolveItemId(driveId, parentPath)

      val folder = DriveItem().apply {
        name = folderName
        folder = Folder()
      }

      val result = graphClient.drives().byDriveId(driveId)
        .items().byDriveItemId(parentId)
        .children()
        .post(folder)

      result?.toFileInfo()
        ?: throw RuntimeException("Create folder returned null for $folderName")
    } catch (e: Exception) {
      throw graphException("create folder", e, "driveId" to driveId, "path" to "$parentPath/$folderName")
    }
  }

  /**
   * Delete a file or folder by item ID.
   */
  fun deleteItem(driveId: String, itemId: String) {
    try {
      graphClient.drives().byDriveId(driveId)
        .items().byDriveItemId(itemId)
        .delete()
    } catch (e: Exception) {
      throw graphException("delete item", e, "driveId" to driveId, "itemId" to itemId)
    }
  }

  /**
   * Search for files within a drive.
   *
   * Results are paginated transparently up to [SharePointConfig.maxPaginationResults].
   */
  fun searchFiles(driveId: String, query: String): List<FileInfo> {
    return try {
      val files = mutableListOf<FileInfo>()

      val firstPage = graphClient.drives().byDriveId(driveId)
        .searchWithQ(query)
        .get()

      firstPage?.value?.forEach { item ->
        if (files.size >= config.maxPaginationResults) return@forEach
        files.add(item.toFileInfo())
      }

      var nextLink = firstPage?.odataNextLink
      while (nextLink != null && files.size < config.maxPaginationResults) {
        val nextPage = graphClient.drives().byDriveId(driveId)
          .searchWithQ(query).withUrl(nextLink)
          .get()
        nextPage?.value?.forEach { item ->
          if (files.size >= config.maxPaginationResults) return@forEach
          files.add(item.toFileInfo())
        }
        nextLink = nextPage?.odataNextLink
      }

      files
    } catch (e: Exception) {
      throw graphException("search files", e, "driveId" to driveId, "query" to query)
    }
  }

  /**
   * Copy an item to a new location within the same drive.
   *
   * Graph API copy is async — returns a DriveItem or null (with a monitoring URL header).
   */
  fun copyItem(driveId: String, itemId: String, destinationPath: String): FileInfo {
    return try {
      val destId = resolveItemId(driveId, destinationPath)
      val parentRef = ItemReference().apply {
        this.driveId = driveId
        this.id = destId
      }

      val copyBody = CopyPostRequestBody().apply {
        this.parentReference = parentRef
      }

      val result = graphClient.drives().byDriveId(driveId)
        .items().byDriveItemId(itemId)
        .copy()
        .post(copyBody)

      result?.toFileInfo() ?: FileInfo(
        id = itemId,
        name = "(copy in progress)",
        webUrl = null,
        size = null,
        isFolder = false,
        lastModified = null,
        createdBy = null,
        mimeType = null
      )
    } catch (e: Exception) {
      throw graphException("copy item", e, "driveId" to driveId, "itemId" to itemId, "destination" to destinationPath)
    }
  }

  /**
   * Move an item to a new location within the same drive.
   */
  fun moveItem(driveId: String, itemId: String, destinationPath: String): FileInfo {
    return try {
      val destId = resolveItemId(driveId, destinationPath)
      val parentRef = ItemReference().apply {
        this.driveId = driveId
        this.id = destId
      }

      val patchBody = DriveItem().apply {
        this.parentReference = parentRef
      }

      val result = graphClient.drives().byDriveId(driveId)
        .items().byDriveItemId(itemId)
        .patch(patchBody)

      result?.toFileInfo()
        ?: throw RuntimeException("Move returned null for item $itemId")
    } catch (e: Exception) {
      throw graphException("move item", e, "driveId" to driveId, "itemId" to itemId, "destination" to destinationPath)
    }
  }

  // ── Upload strategies ────────────────────────────────────────────────

  /**
   * Simple single-request upload for files <= 4MB.
   */
  private fun simpleUpload(driveId: String, parentPath: String, fileName: String, content: ByteArray): FileInfo {
    val fullPath = buildPath(parentPath, fileName)
    val encodedPath = encodePath(fullPath)
    val inputStream = ByteArrayInputStream(content)

    val uploadUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/root:/$encodedPath:/content"
    val result = graphClient.drives().byDriveId(driveId)
      .items().byDriveItemId("root")
      .content()
      .withUrl(uploadUrl)
      .put(inputStream)

    return result?.toFileInfo()
      ?: throw RuntimeException("Simple upload returned null result for $fullPath")
  }

  /**
   * Chunked resumable upload for files > 4MB using [LargeFileUploadTask].
   *
   * Creates an upload session, then uploads in chunks of [SharePointConfig.uploadChunkBytes].
   * Upload URLs are pre-authenticated — uses [AnonymousAuthenticationProvider] to avoid 401.
   */
  private fun chunkedUpload(driveId: String, parentPath: String, fileName: String, content: ByteArray): FileInfo {
    val fullPath = buildPath(parentPath, fileName)
    val encodedPath = encodePath(fullPath)

    // Path-based item ID for createUploadSession: "root:/{path}:"
    val pathBasedItemId = "root:/$encodedPath:"

    val sessionRequest = CreateUploadSessionPostRequestBody().apply {
      item = DriveItemUploadableProperties().apply {
        additionalData["@microsoft.graph.conflictBehavior"] = "replace"
      }
    }

    val uploadSession = graphClient.drives().byDriveId(driveId)
      .items().byDriveItemId(pathBasedItemId)
      .createUploadSession()
      .post(sessionRequest)
      ?: throw RuntimeException("Failed to create upload session for $fullPath")

    val inputStream = ByteArrayInputStream(content)
    val uploadTask = LargeFileUploadTask(
      graphClient.requestAdapter,
      uploadSession,
      inputStream,
      content.size.toLong(),
      config.uploadChunkBytes,
      DriveItem::createFromDiscriminatorValue
    )

    val callback = IProgressCallback { current, max ->
      logger.debug("Upload progress for {}: {} / {} bytes", fullPath, current, max)
    }

    val result = uploadTask.upload(3, callback)

    if (!result.isUploadSuccessful) {
      throw RuntimeException("Chunked upload failed for $fullPath")
    }

    return result.itemResponse?.toFileInfo()
      ?: throw RuntimeException("Chunked upload returned null result for $fullPath")
  }

  // ── Helpers ───────────────────────────────────────────────────────────

  /**
   * Resolve a path to a drive item ID.
   *
   * For root or null/blank paths, returns the literal "root" item ID
   * (Graph API accepts "root" as a well-known ID — no extra API call needed).
   * For sub-paths, uses URL-based path addressing to resolve the item.
   */
  private fun resolveItemId(driveId: String, path: String?): String {
    if (path.isNullOrBlank() || path == "/") {
      return "root"
    }

    val normalizedPath = path.trimStart('/').trimEnd('/')
    val encodedPath = encodePath(normalizedPath)
    val pathUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/root:/$encodedPath"
    val item = graphClient.drives().byDriveId(driveId)
      .items().byDriveItemId("root")
      .withUrl(pathUrl)
      .get()
      ?: throw RuntimeException("Failed to resolve path '$path' in drive $driveId")
    return item.id ?: throw RuntimeException("Item at path '$path' has no ID in drive $driveId")
  }

  /**
   * URL-encode each segment of a path individually, preserving '/' separators.
   *
   * Graph API requires path segments to be percent-encoded but '/' must remain literal.
   * Handles spaces, unicode, and special characters (#, ?, %, &, etc.).
   */
  private fun encodePath(path: String): String {
    return path.split("/")
      .filter { it.isNotEmpty() }
      .joinToString("/") { segment ->
        URLEncoder.encode(segment, StandardCharsets.UTF_8)
          .replace("+", "%20") // URLEncoder uses + for spaces, Graph API expects %20
      }
  }

  private fun com.microsoft.graph.models.Site.toSiteInfo(): SiteInfo {
    return SiteInfo(
      id = this.id ?: "",
      name = this.name ?: "",
      displayName = this.displayName ?: "",
      webUrl = this.webUrl ?: "",
      description = this.description
    )
  }

  private fun DriveItem.toFileInfo(): FileInfo {
    return FileInfo(
      id = this.id ?: "",
      name = this.name ?: "",
      webUrl = this.webUrl,
      size = this.size,
      isFolder = this.folder != null,
      lastModified = this.lastModifiedDateTime,
      createdBy = this.createdBy?.user?.displayName,
      mimeType = this.file?.mimeType
    )
  }

  private fun buildPath(parentPath: String, fileName: String): String {
    val normalized = parentPath.trimStart('/').trimEnd('/')
    return if (normalized.isEmpty()) fileName else "$normalized/$fileName"
  }

  // ── Error handling ───────────────────────────────────────────────────

  /**
   * Extract structured error details from Graph API exceptions.
   *
   * [ODataError] contains a nested `error` object with `code` (e.g., "accessDenied",
   * "itemNotFound") and `message` (human-readable). This method extracts those details
   * and wraps them in a [RuntimeException] with actionable context.
   */
  private fun graphException(operation: String, cause: Exception, vararg context: Pair<String, Any?>): RuntimeException {
    val contextStr = context
      .filter { it.second != null }
      .joinToString(", ") { "${it.first}=${it.second}" }

    return if (cause is ODataError) {
      val code = cause.error?.code ?: "unknown"
      val message = cause.error?.message ?: cause.message ?: "Unknown error"
      logger.error("Graph API error during {} [{}]: code={}, message={}", operation, contextStr, code, message, cause)
      RuntimeException("Failed to $operation [$contextStr] (code=$code): $message", cause)
    } else if (cause is RuntimeException && cause !is ODataError) {
      // Re-throw RuntimeExceptions from our own code (e.g., resolveItemId failures)
      // without double-wrapping
      logger.error("Failed to {} [{}]: {}", operation, contextStr, cause.message, cause)
      cause
    } else {
      logger.error("Failed to {} [{}]: {}", operation, contextStr, cause.message, cause)
      RuntimeException("Failed to $operation [$contextStr]: ${cause.message}", cause)
    }
  }
}
