package com.utisha.mcp.sharepoint

/**
 * Configuration for connecting to SharePoint via Microsoft Graph API.
 *
 * Uses Azure AD app registration with client credentials (client_credentials grant).
 * Required Azure AD permissions: Sites.Read.All, Sites.ReadWrite.All, Files.ReadWrite.All
 *
 * @param tenantId Azure AD tenant ID
 * @param clientId Azure AD application (client) ID
 * @param clientSecret Azure AD client secret
 * @param siteUrl Optional: restrict operations to a specific SharePoint site URL.
 *        When set, `listSites()` will only return sites matching this URL.
 * @param maxSimpleUploadBytes Max file size in bytes for simple (single-request) upload.
 *        Files larger than this use chunked resumable upload via LargeFileUploadTask.
 *        Graph API hard limit for simple upload is 4MB. Default: 4,000,000 bytes.
 * @param uploadChunkBytes Chunk size for large file uploads. Must be a multiple of
 *        320 KiB (327,680 bytes). Default: 3,276,800 bytes (3.2 MB = 320 KiB x 10).
 * @param maxUploadBytes Hard limit on upload file size in bytes. Files exceeding this
 *        are rejected before upload. 0 = no limit. Default: 0 (unlimited).
 * @param maxPaginationResults Safety cap for paginated list operations to prevent
 *        runaway pagination. Default: 500 items.
 */
data class SharePointConfig(
  val tenantId: String,
  val clientId: String,
  val clientSecret: String,
  val siteUrl: String? = null,
  val maxUploadBytes: Long = 0L,
  val maxSimpleUploadBytes: Long = 4_000_000L,
  val uploadChunkBytes: Long = DEFAULT_CHUNK_SIZE,
  val maxPaginationResults: Int = 500
) {
  companion object {
    /** 320 KiB — the required alignment unit for Graph API upload chunks. */
    const val CHUNK_ALIGNMENT = 327_680L

    /** Default chunk size: 320 KiB x 10 = 3.2 MB */
    const val DEFAULT_CHUNK_SIZE = CHUNK_ALIGNMENT * 10

    /**
     * Create a [SharePointConfig] from environment variables.
     *
     * Required:
     * - `SHAREPOINT_TENANT_ID`
     * - `SHAREPOINT_CLIENT_ID`
     * - `SHAREPOINT_CLIENT_SECRET`
     *
     * Optional:
     * - `SHAREPOINT_SITE_URL` — restrict to a specific site
     * - `SHAREPOINT_MAX_UPLOAD_BYTES` — hard upload limit (default: 0 = unlimited)
     * - `SHAREPOINT_MAX_PAGINATION_RESULTS` — pagination cap (default: 500)
     *
     * @throws IllegalStateException if required environment variables are missing
     */
    fun fromEnv(): SharePointConfig {
      val tenantId = requireEnv("SHAREPOINT_TENANT_ID")
      val clientId = requireEnv("SHAREPOINT_CLIENT_ID")
      val clientSecret = requireEnv("SHAREPOINT_CLIENT_SECRET")

      return SharePointConfig(
        tenantId = tenantId,
        clientId = clientId,
        clientSecret = clientSecret,
        siteUrl = System.getenv("SHAREPOINT_SITE_URL")?.takeIf { it.isNotBlank() },
        maxUploadBytes = System.getenv("SHAREPOINT_MAX_UPLOAD_BYTES")
          ?.toLongOrNull() ?: 0L,
        maxPaginationResults = System.getenv("SHAREPOINT_MAX_PAGINATION_RESULTS")
          ?.toIntOrNull() ?: 500
      )
    }

    private fun requireEnv(name: String): String {
      return System.getenv(name)?.takeIf { it.isNotBlank() }
        ?: throw IllegalStateException(
          "Required environment variable '$name' is not set. " +
            "See README.md for configuration instructions."
        )
    }
  }

  init {
    require(maxUploadBytes >= 0) {
      "maxUploadBytes must be >= 0 (0 = unlimited), got $maxUploadBytes"
    }
    require(uploadChunkBytes > 0 && uploadChunkBytes % CHUNK_ALIGNMENT == 0L) {
      "uploadChunkBytes must be a positive multiple of 320 KiB ($CHUNK_ALIGNMENT bytes), got $uploadChunkBytes"
    }
    require(maxSimpleUploadBytes in 0..4_000_000L) {
      "maxSimpleUploadBytes must be between 0 and 4,000,000 (Graph API hard limit), got $maxSimpleUploadBytes"
    }
    require(maxPaginationResults in 1..5000) {
      "maxPaginationResults must be between 1 and 5000, got $maxPaginationResults"
    }
  }
}
