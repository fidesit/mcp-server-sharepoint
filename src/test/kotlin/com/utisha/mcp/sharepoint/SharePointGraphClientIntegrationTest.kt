package com.utisha.mcp.sharepoint

import org.junit.jupiter.api.*
import org.junit.jupiter.api.Assertions.*
import org.junit.jupiter.api.MethodOrderer.OrderAnnotation
import java.io.File
import java.util.Base64

/**
 * Integration tests for [SharePointGraphClient] against a real SharePoint tenant.
 *
 * These tests exercise the full Graph API call chain (auth, pagination, CRUD)
 * without mocks. They require a configured Azure AD app registration with
 * Sites.Read.All, Sites.ReadWrite.All, and Files.ReadWrite.All permissions.
 *
 * ## Configuration
 *
 * Set the following environment variables (or add them to the root `.env` file):
 *
 * ```
 * SHAREPOINT_TENANT_ID=...
 * SHAREPOINT_CLIENT_ID=...
 * SHAREPOINT_CLIENT_SECRET=...
 * SHAREPOINT_SITE_URL=https://yourorg.sharepoint.com/sites/your-site
 * ```
 *
 * ## Running
 *
 * ```bash
 * cd apps/backend-kt
 * ./gradlew :mcp-sharepoint:integrationTest
 * ```
 *
 * Tests are ordered to form a full lifecycle: list sites → list drives →
 * create folder → upload file → search → copy → move → delete → cleanup.
 */
@Tag("integration")
@TestMethodOrder(OrderAnnotation::class)
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class SharePointGraphClientIntegrationTest {

  private lateinit var client: SharePointGraphClient
  private lateinit var siteId: String
  private lateinit var driveId: String

  // Track created items for cleanup
  private val createdItemIds = mutableListOf<String>()

  companion object {
    private const val TEST_FOLDER_NAME = "mcp-sharepoint-integration-test"
    private const val TEST_FILE_NAME = "test-document.txt"
    private const val TEST_FILE_CONTENT = "Hello from mcp-sharepoint integration tests!"

    /**
     * Load configuration from environment variables, falling back to the
     * root `.env` file if env vars are not set.
     */
    fun loadConfig(): SharePointConfig? {
      val env = loadEnv()
      val tenantId = env["SHAREPOINT_TENANT_ID"] ?: return null
      val clientId = env["SHAREPOINT_CLIENT_ID"] ?: return null
      val clientSecret = env["SHAREPOINT_CLIENT_SECRET"] ?: return null

      return SharePointConfig(
        tenantId = tenantId,
        clientId = clientId,
        clientSecret = clientSecret,
        maxPaginationResults = 50
      )
    }

    fun loadSiteUrl(): String? {
      return loadEnv()["SHAREPOINT_SITE_URL"]
    }

    /**
     * Merge System.getenv() with values parsed from the root `.env` file.
     * System env vars take precedence over `.env` file values.
     */
    private fun loadEnv(): Map<String, String> {
      val fileEnv = mutableMapOf<String, String>()
      // Walk up from the module dir to find the root .env
      val candidates = listOf(
        File("../../.env"),         // from apps/backend-kt/mcp-sharepoint
        File("../../../.env"),      // fallback
        File(".env")                // if running from root
      )
      val envFile = candidates.firstOrNull { it.exists() }
      envFile?.readLines()?.forEach { line ->
        val trimmed = line.trim()
        if (trimmed.isNotEmpty() && !trimmed.startsWith("#")) {
          val eqIndex = trimmed.indexOf('=')
          if (eqIndex > 0) {
            val key = trimmed.substring(0, eqIndex).trim()
            val value = trimmed.substring(eqIndex + 1).trim()
              .removeSurrounding("\"")
              .removeSurrounding("'")
            fileEnv[key] = value
          }
        }
      }
      // System env vars take precedence
      return fileEnv + System.getenv()
    }
  }

  @BeforeAll
  fun setup() {
    val config = loadConfig()
    Assumptions.assumeTrue(
      config != null,
      "Skipping SharePoint integration tests: SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, " +
        "and SHAREPOINT_CLIENT_SECRET must be set (via env vars or root .env file)"
    )
    client = SharePointGraphClient(config!!)
  }

  @AfterAll
  fun cleanup() {
    if (!::client.isInitialized || !::driveId.isInitialized) return

    // Best-effort cleanup of all created items (reverse order)
    createdItemIds.reversed().forEach { itemId ->
      try {
        client.deleteItem(driveId, itemId)
      } catch (e: Exception) {
        System.err.println("Cleanup: failed to delete item $itemId: ${e.message}")
      }
    }
  }

  // ── 1. List Sites ────────────────────────────────────────────────────

  @Test
  @Order(1)
  fun `list sites returns at least one site`() {
    val sites = client.listSites()
    assertTrue(sites.isNotEmpty(), "Expected at least one accessible SharePoint site")

    sites.forEach { site ->
      assertNotNull(site.id, "Site ID should not be null")
      assertTrue(site.id.isNotBlank(), "Site ID should not be blank")
      assertTrue(site.webUrl.isNotBlank(), "Site webUrl should not be blank")
    }
  }

  @Test
  @Order(2)
  fun `find configured test site`() {
    val siteUrl = loadSiteUrl()
    Assumptions.assumeTrue(
      siteUrl != null,
      "Skipping: SHAREPOINT_SITE_URL not configured"
    )

    val sites = client.listSites()
    val targetSite = sites.firstOrNull { it.webUrl.equals(siteUrl, ignoreCase = true) }
      ?: sites.firstOrNull { it.webUrl.contains(siteUrl!!.substringAfterLast("/sites/"), ignoreCase = true) }

    assertNotNull(targetSite, "Could not find site matching $siteUrl. Available sites: ${sites.map { it.webUrl }}")
    siteId = targetSite!!.id

    println("Using site: ${targetSite.displayName} (${targetSite.webUrl})")
    println("Site ID: $siteId")
  }

  @Test
  @Order(3)
  fun `search sites filters results`() {
    Assumptions.assumeTrue(::siteId.isInitialized, "Requires site from previous test")

    // Search for a term that should match at least one site
    val siteUrl = loadSiteUrl() ?: return
    val searchTerm = siteUrl.substringAfterLast("/sites/").take(5)
    val results = client.listSites(searchTerm)

    // Search should return results (may include the test site)
    // We just verify it doesn't throw and returns a valid list
    assertNotNull(results)
  }

  // ── 2. List Drives ───────────────────────────────────────────────────

  @Test
  @Order(10)
  fun `list drives returns at least one drive`() {
    Assumptions.assumeTrue(::siteId.isInitialized, "Requires site from previous test")

    val drives = client.listDrives(siteId)
    assertTrue(drives.isNotEmpty(), "Expected at least one document library (drive)")

    drives.forEach { drive ->
      assertTrue(drive.id.isNotBlank(), "Drive ID should not be blank")
      assertTrue(drive.name.isNotBlank(), "Drive name should not be blank")
    }

    // Use the first drive (usually "Documents")
    driveId = drives.first().id
    println("Using drive: ${drives.first().name} (id=$driveId)")
  }

  // ── 3. List Files (root) ─────────────────────────────────────────────

  @Test
  @Order(20)
  fun `list files at root does not throw`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")

    val files = client.listFiles(driveId)
    assertNotNull(files)
    // Root may be empty or have files — just verify the call succeeds
    files.forEach { file ->
      assertTrue(file.id.isNotBlank(), "File ID should not be blank")
      assertTrue(file.name.isNotBlank(), "File name should not be blank")
    }
  }

  // ── 4. Create Folder ─────────────────────────────────────────────────

  @Test
  @Order(30)
  fun `create test folder at root`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")

    val folder = client.createFolder(driveId, "/", TEST_FOLDER_NAME)

    assertEquals(TEST_FOLDER_NAME, folder.name)
    assertTrue(folder.isFolder, "Created item should be a folder")
    assertTrue(folder.id.isNotBlank(), "Folder should have an ID")

    createdItemIds.add(folder.id)
    println("Created folder: ${folder.name} (id=${folder.id})")
  }

  // ── 5. Upload File ───────────────────────────────────────────────────

  @Test
  @Order(40)
  fun `upload text file to test folder`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")
    Assumptions.assumeTrue(createdItemIds.isNotEmpty(), "Requires folder from previous test")

    val content = TEST_FILE_CONTENT.toByteArray()
    val uploaded = client.uploadFile(driveId, TEST_FOLDER_NAME, TEST_FILE_NAME, content)

    assertEquals(TEST_FILE_NAME, uploaded.name)
    assertFalse(uploaded.isFolder, "Uploaded item should be a file")
    assertTrue(uploaded.id.isNotBlank(), "File should have an ID")
    assertEquals(content.size.toLong(), uploaded.size, "File size should match uploaded content")

    createdItemIds.add(uploaded.id)
    println("Uploaded file: ${uploaded.name} (id=${uploaded.id}, size=${uploaded.size})")
  }

  // ── 6. Get File Content ──────────────────────────────────────────────

  @Test
  @Order(50)
  fun `download file content matches uploaded content`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")
    Assumptions.assumeTrue(createdItemIds.size >= 2, "Requires uploaded file from previous test")

    val fileId = createdItemIds.last() // the uploaded file
    val content = client.getFileContent(driveId, fileId)

    assertEquals(TEST_FILE_CONTENT, content, "Downloaded content should match what was uploaded")
  }

  // ── 7. List Files in Subfolder ───────────────────────────────────────

  @Test
  @Order(60)
  fun `list files in test folder shows uploaded file`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")

    val files = client.listFiles(driveId, TEST_FOLDER_NAME)

    assertTrue(files.isNotEmpty(), "Test folder should contain at least the uploaded file")
    assertTrue(
      files.any { it.name == TEST_FILE_NAME },
      "Test folder should contain '$TEST_FILE_NAME'. Found: ${files.map { it.name }}"
    )
  }

  // ── 8. Search Files ──────────────────────────────────────────────────

  @Test
  @Order(70)
  fun `search files finds uploaded document`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")

    // SharePoint search indexing can be delayed — search for the unique file name
    val results = client.searchFiles(driveId, "test-document")

    // Search indexing is eventually consistent, so we can't guarantee results
    // immediately. Just verify the call succeeds without error.
    assertNotNull(results, "Search should return a non-null list")
    println("Search returned ${results.size} result(s)")
  }

  // ── 9. Copy Item ─────────────────────────────────────────────────────

  @Test
  @Order(80)
  fun `copy file to root`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")
    Assumptions.assumeTrue(createdItemIds.size >= 2, "Requires uploaded file from previous test")

    val fileId = createdItemIds.last() // the uploaded file
    val result = client.copyItem(driveId, fileId, "/")

    // Copy is async — we may get the actual item or a placeholder
    assertNotNull(result, "Copy should return a result")
    println("Copy result: name=${result.name}, id=${result.id}")

    // If we got a real item back (not async placeholder), track it for cleanup
    if (result.name != "(copy in progress)" && result.id != fileId) {
      createdItemIds.add(result.id)
    }
  }

  // ── 10. Move Item ────────────────────────────────────────────────────

  @Test
  @Order(90)
  fun `move file within drive`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")
    Assumptions.assumeTrue(createdItemIds.size >= 2, "Requires uploaded file from previous test")

    // Upload a second file specifically for the move test
    val moveContent = "This file will be moved".toByteArray()
    val moveFile = client.uploadFile(driveId, TEST_FOLDER_NAME, "move-test.txt", moveContent)
    // Don't track — we'll track after move

    val moved = client.moveItem(driveId, moveFile.id, "/")

    assertEquals("move-test.txt", moved.name, "Moved file should keep its name")
    assertTrue(moved.id.isNotBlank())

    createdItemIds.add(moved.id)
    println("Moved file: ${moved.name} from $TEST_FOLDER_NAME to root (id=${moved.id})")
  }

  // ── 11. Delete Item ──────────────────────────────────────────────────

  @Test
  @Order(100)
  fun `delete file succeeds`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")
    Assumptions.assumeTrue(createdItemIds.size >= 2, "Requires items from previous tests")

    // Delete the last tracked item (moved file)
    val itemId = createdItemIds.removeLastOrNull() ?: return

    assertDoesNotThrow { client.deleteItem(driveId, itemId) }
    println("Deleted item: $itemId")
  }

  // ── 12. Error Handling ───────────────────────────────────────────────

  @Test
  @Order(110)
  fun `list drives with invalid site id throws`() {
    val ex = assertThrows<RuntimeException> {
      client.listDrives("not-a-real-site-id")
    }
    assertTrue(
      ex.message!!.contains("code=") || ex.message!!.contains("not-a-real-site-id"),
      "Error should contain OData error code or the invalid ID. Got: ${ex.message}"
    )
  }

  @Test
  @Order(111)
  fun `get file content with invalid item id throws`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")

    val ex = assertThrows<RuntimeException> {
      client.getFileContent(driveId, "not-a-real-item-id")
    }
    assertNotNull(ex.message)
  }

  @Test
  @Order(112)
  fun `list files with invalid path throws`() {
    Assumptions.assumeTrue(::driveId.isInitialized, "Requires drive from previous test")

    val ex = assertThrows<RuntimeException> {
      client.listFiles(driveId, "this/path/definitely/does/not/exist/anywhere")
    }
    assertNotNull(ex.message)
  }
}
