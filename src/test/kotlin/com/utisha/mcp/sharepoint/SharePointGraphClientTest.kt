package com.utisha.mcp.sharepoint

import com.microsoft.graph.core.models.UploadResult
import com.microsoft.graph.drives.DrivesRequestBuilder
import com.microsoft.graph.drives.item.DriveItemRequestBuilder
import com.microsoft.graph.drives.item.items.ItemsRequestBuilder
import com.microsoft.graph.drives.item.items.item.DriveItemItemRequestBuilder
import com.microsoft.graph.drives.item.items.item.children.ChildrenRequestBuilder
import com.microsoft.graph.drives.item.items.item.content.ContentRequestBuilder
import com.microsoft.graph.drives.item.items.item.copy.CopyRequestBuilder
import com.microsoft.graph.drives.item.items.item.createuploadsession.CreateUploadSessionRequestBuilder
import com.microsoft.graph.drives.item.root.RootRequestBuilder
import com.microsoft.graph.drives.item.searchwithq.SearchWithQGetResponse
import com.microsoft.graph.drives.item.searchwithq.SearchWithQRequestBuilder
import com.microsoft.graph.models.DriveCollectionResponse
import com.microsoft.graph.models.DriveItem
import com.microsoft.graph.models.DriveItemCollectionResponse
import com.microsoft.graph.models.File
import com.microsoft.graph.models.Folder
import com.microsoft.graph.models.IdentitySet
import com.microsoft.graph.models.Identity
import com.microsoft.graph.models.Quota
import com.microsoft.graph.models.Site
import com.microsoft.graph.models.SiteCollectionResponse
import com.microsoft.graph.models.UploadSession
import com.microsoft.graph.models.odataerrors.MainError
import com.microsoft.graph.models.odataerrors.ODataError
import com.microsoft.graph.serviceclient.GraphServiceClient
import com.microsoft.graph.sites.SitesRequestBuilder
import com.microsoft.graph.sites.item.SiteItemRequestBuilder
import com.microsoft.graph.sites.item.drives.DrivesRequestBuilder as SiteDrivesRequestBuilder
import io.mockk.*
import org.junit.jupiter.api.AfterEach
import org.junit.jupiter.api.Assertions.*
import org.junit.jupiter.api.BeforeEach
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream
import java.time.OffsetDateTime

/**
 * Unit tests for [SharePointGraphClient].
 *
 * Mocks the Microsoft Graph SDK v6 (Kiota) fluent builder chain to test
 * each operation in isolation without making real Graph API calls.
 *
 * IMPORTANT: Graph SDK uses deep fluent chains (drives → byDriveId → items → byDriveItemId →
 * children/content/copy). All mocks for a single test must share the SAME root chain
 * (drivesBuilder, driveBuilder) because `resolveItemId()` and the operation itself both
 * call `mockGraphClient.drives()`. Using separate mock chains would cause MockK to return
 * the LAST-registered stub, breaking the earlier call.
 */
class SharePointGraphClientTest {

  private val config = SharePointConfig(
    tenantId = "test-tenant",
    clientId = "test-client",
    clientSecret = "test-secret",
    siteUrl = null
  )

  private val mockGraphClient = mockk<GraphServiceClient>()
  private lateinit var client: SharePointGraphClient

  @BeforeEach
  fun setup() {
    clearAllMocks()
    client = SharePointGraphClient(config)

    // Replace the lazy graphClient with our mock via reflection
    val field = SharePointGraphClient::class.java.getDeclaredField("graphClient\$delegate")
    field.isAccessible = true
    field.set(client, lazy { mockGraphClient })
  }

  @AfterEach
  fun teardown() {
    clearAllMocks()
  }

  // ── listSites ────────────────────────────────────────────────────────

  @Test
  fun `listSites - returns sites from graph api`() {
    val site = Site().apply {
      id = "site-1"
      name = "TestSite"
      displayName = "Test Site"
      webUrl = "https://contoso.sharepoint.com/sites/TestSite"
      description = "A test site"
    }

    val response = SiteCollectionResponse().apply { value = listOf(site) }
    val sitesBuilder = mockk<SitesRequestBuilder>()
    every { mockGraphClient.sites() } returns sitesBuilder
    every { sitesBuilder.get() } returns response

    val result = client.listSites()

    assertEquals(1, result.size)
    assertEquals("site-1", result[0].id)
    assertEquals("TestSite", result[0].name)
    assertEquals("Test Site", result[0].displayName)
    assertEquals("https://contoso.sharepoint.com/sites/TestSite", result[0].webUrl)
    assertEquals("A test site", result[0].description)
  }

  @Test
  fun `listSites - with search query passes search parameter`() {
    val response = SiteCollectionResponse().apply { value = emptyList() }
    val sitesBuilder = mockk<SitesRequestBuilder>()
    every { mockGraphClient.sites() } returns sitesBuilder
    every { sitesBuilder.get(any()) } returns response

    val result = client.listSites("finance")

    assertEquals(0, result.size)
    verify { sitesBuilder.get(any()) }
  }

  @Test
  fun `listSites - null response returns empty list`() {
    val sitesBuilder = mockk<SitesRequestBuilder>()
    every { mockGraphClient.sites() } returns sitesBuilder
    every { sitesBuilder.get() } returns null

    assertTrue(client.listSites().isEmpty())
  }

  @Test
  fun `listSites - graph api error throws RuntimeException`() {
    val sitesBuilder = mockk<SitesRequestBuilder>()
    every { mockGraphClient.sites() } returns sitesBuilder
    every { sitesBuilder.get() } throws RuntimeException("Graph API error")

    val ex = assertThrows(RuntimeException::class.java) { client.listSites() }
    assertTrue(ex.message!!.contains("Graph API error"))
  }

  @Test
  fun `listSites - ODataError extracts error code and message`() {
    val odataError = ODataError().apply {
      error = MainError().apply {
        code = "accessDenied"
        message = "Insufficient privileges to complete the operation."
      }
    }

    val sitesBuilder = mockk<SitesRequestBuilder>()
    every { mockGraphClient.sites() } returns sitesBuilder
    every { sitesBuilder.get() } throws odataError

    val ex = assertThrows(RuntimeException::class.java) { client.listSites() }
    assertTrue(ex.message!!.contains("code=accessDenied"), "Error should contain OData code, got: ${ex.message}")
    assertTrue(ex.message!!.contains("Insufficient privileges"), "Error should contain OData message, got: ${ex.message}")
  }

  // ── listDrives ───────────────────────────────────────────────────────

  @Test
  fun `listDrives - returns drives for site`() {
    val drive = com.microsoft.graph.models.Drive().apply {
      id = "drive-1"; name = "Documents"; driveType = "documentLibrary"
      webUrl = "https://contoso.sharepoint.com/sites/TestSite/Documents"
      quota = Quota().apply { total = 1073741824L; used = 536870912L }
    }
    val driveResponse = DriveCollectionResponse().apply { value = listOf(drive) }

    val sitesBuilder = mockk<SitesRequestBuilder>()
    val siteItemBuilder = mockk<SiteItemRequestBuilder>()
    val siteDrivesBuilder = mockk<SiteDrivesRequestBuilder>()
    every { mockGraphClient.sites() } returns sitesBuilder
    every { sitesBuilder.bySiteId("site-1") } returns siteItemBuilder
    every { siteItemBuilder.drives() } returns siteDrivesBuilder
    every { siteDrivesBuilder.get() } returns driveResponse

    val result = client.listDrives("site-1")

    assertEquals(1, result.size)
    assertEquals("drive-1", result[0].id)
    assertEquals("Documents", result[0].name)
    assertEquals("documentLibrary", result[0].driveType)
    assertEquals(1073741824L, result[0].totalSize)
    assertEquals(536870912L, result[0].usedSize)
  }

  @Test
  fun `listDrives - null response returns empty list`() {
    val sitesBuilder = mockk<SitesRequestBuilder>()
    val siteItemBuilder = mockk<SiteItemRequestBuilder>()
    val siteDrivesBuilder = mockk<SiteDrivesRequestBuilder>()
    every { mockGraphClient.sites() } returns sitesBuilder
    every { sitesBuilder.bySiteId("site-1") } returns siteItemBuilder
    every { siteItemBuilder.drives() } returns siteDrivesBuilder
    every { siteDrivesBuilder.get() } returns null

    assertTrue(client.listDrives("site-1").isEmpty())
  }

  @Test
  fun `listDrives - graph api error throws RuntimeException with context`() {
    val sitesBuilder = mockk<SitesRequestBuilder>()
    val siteItemBuilder = mockk<SiteItemRequestBuilder>()
    val siteDrivesBuilder = mockk<SiteDrivesRequestBuilder>()
    every { mockGraphClient.sites() } returns sitesBuilder
    every { sitesBuilder.bySiteId("site-1") } returns siteItemBuilder
    every { siteItemBuilder.drives() } returns siteDrivesBuilder
    every { siteDrivesBuilder.get() } throws RuntimeException("Not found")

    val ex = assertThrows(RuntimeException::class.java) { client.listDrives("site-1") }
    assertTrue(ex.message!!.contains("Not found"))
  }

  // ── listFiles ────────────────────────────────────────────────────────

  @Test
  fun `listFiles - at root lists children using literal root id`() {
    // resolveItemId for root/null should return "root" without making an API call
    val childItem = buildDriveItem("item-1", "report.docx", isFolder = false, size = 1024)
    val childrenResponse = DriveItemCollectionResponse().apply { value = listOf(childItem) }

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val childrenBuilder = mockk<ChildrenRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    // listFiles → items().byDriveItemId("root").children().get()
    // No root().get() call — resolveItemId returns "root" directly for root path
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("root") } returns itemBuilder
    every { itemBuilder.children() } returns childrenBuilder
    every { childrenBuilder.get() } returns childrenResponse

    val result = client.listFiles("drive-1")

    assertEquals(1, result.size)
    assertEquals("item-1", result[0].id)
    assertEquals("report.docx", result[0].name)
    assertFalse(result[0].isFolder)
    assertEquals(1024L, result[0].size)
  }

  @Test
  fun `listFiles - at subpath resolves path to item id first`() {
    val pathItem = DriveItem().apply { id = "folder-id" }
    val childItem = buildDriveItem("item-2", "notes.txt", isFolder = false, size = 256)
    val childrenResponse = DriveItemCollectionResponse().apply { value = listOf(childItem) }

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val rootItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val pathItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val folderItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val childrenBuilder = mockk<ChildrenRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    // resolveItemId(driveId, "Documents/Reports") → items().byDriveItemId("root").withUrl(...).get()
    every { itemsBuilder.byDriveItemId("root") } returns rootItemBuilder
    every { rootItemBuilder.withUrl(any()) } returns pathItemBuilder
    every { pathItemBuilder.get() } returns pathItem
    // listFiles → items().byDriveItemId("folder-id").children().get()
    every { itemsBuilder.byDriveItemId("folder-id") } returns folderItemBuilder
    every { folderItemBuilder.children() } returns childrenBuilder
    every { childrenBuilder.get() } returns childrenResponse

    val result = client.listFiles("drive-1", "Documents/Reports")

    assertEquals(1, result.size)
    assertEquals("item-2", result[0].id)
  }

  // ── getFileContent ───────────────────────────────────────────────────

  @Test
  fun `getFileContent - returns file content as text`() {
    val stream = ByteArrayInputStream("Hello, SharePoint!".toByteArray())

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val contentBuilder = mockk<ContentRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("item-1") } returns itemBuilder
    every { itemBuilder.content() } returns contentBuilder
    every { contentBuilder.get() } returns stream

    assertEquals("Hello, SharePoint!", client.getFileContent("drive-1", "item-1"))
  }

  @Test
  fun `getFileContent - null stream throws RuntimeException`() {
    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val contentBuilder = mockk<ContentRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("item-1") } returns itemBuilder
    every { itemBuilder.content() } returns contentBuilder
    every { contentBuilder.get() } returns null

    val ex = assertThrows(RuntimeException::class.java) { client.getFileContent("drive-1", "item-1") }
    assertTrue(ex.message!!.contains("No content returned"))
  }

  // ── uploadFile (simple) ─────────────────────────────────────────────

  @Test
  fun `uploadFile - small file uses simple upload`() {
    val content = "Hello World".toByteArray()
    val uploadedItem = buildDriveItem("uploaded-1", "test.txt", isFolder = false, size = content.size.toLong())

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val contentBuilder = mockk<ContentRequestBuilder>()
    val urlContentBuilder = mockk<ContentRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("root") } returns itemBuilder
    every { itemBuilder.content() } returns contentBuilder
    every { contentBuilder.withUrl(any()) } returns urlContentBuilder
    every { urlContentBuilder.put(any()) } returns uploadedItem

    val result = client.uploadFile("drive-1", "/", "test.txt", content)

    assertEquals("uploaded-1", result.id)
    assertEquals("test.txt", result.name)
    verify { urlContentBuilder.put(any()) }
  }

  @Test
  fun `uploadFile - simple upload null result throws RuntimeException`() {
    val content = "data".toByteArray()

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val contentBuilder = mockk<ContentRequestBuilder>()
    val urlContentBuilder = mockk<ContentRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("root") } returns itemBuilder
    every { itemBuilder.content() } returns contentBuilder
    every { contentBuilder.withUrl(any()) } returns urlContentBuilder
    every { urlContentBuilder.put(any()) } returns null

    val ex = assertThrows(RuntimeException::class.java) {
      client.uploadFile("drive-1", "/", "test.txt", content)
    }
    assertTrue(ex.message!!.contains("Simple upload returned null"))
  }

  @Test
  fun `uploadFile - simple upload with path containing special characters encodes URL`() {
    val content = "data".toByteArray()
    val uploadedItem = buildDriveItem("uploaded-2", "report #1.txt", isFolder = false, size = content.size.toLong())

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val contentBuilder = mockk<ContentRequestBuilder>()
    val urlContentBuilder = mockk<ContentRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("root") } returns itemBuilder
    every { itemBuilder.content() } returns contentBuilder
    every { contentBuilder.withUrl(match { it.contains("%23") && it.contains("%20") }) } returns urlContentBuilder
    every { urlContentBuilder.put(any()) } returns uploadedItem

    val result = client.uploadFile("drive-1", "My Folder", "report #1.txt", content)

    assertEquals("uploaded-2", result.id)
    // Verify the URL contained encoded characters
    verify { contentBuilder.withUrl(match { it.contains("My%20Folder") && it.contains("report%20%231.txt") }) }
  }

  @Test
  fun `uploadFile - graph api error throws RuntimeException with context`() {
    val content = "data".toByteArray()

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val contentBuilder = mockk<ContentRequestBuilder>()
    val urlContentBuilder = mockk<ContentRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("root") } returns itemBuilder
    every { itemBuilder.content() } returns contentBuilder
    every { contentBuilder.withUrl(any()) } returns urlContentBuilder
    every { urlContentBuilder.put(any()) } throws RuntimeException("Request entity too large")

    val ex = assertThrows(RuntimeException::class.java) {
      client.uploadFile("drive-1", "/", "big.bin", content)
    }
    assertTrue(ex.message!!.contains("Request entity too large"))
  }

  // ── createFolder ─────────────────────────────────────────────────────

  @Test
  fun `createFolder - creates folder at root`() {
    val createdFolder = buildDriveItem("folder-new", "NewFolder", isFolder = true, size = 0)

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val childrenBuilder = mockk<ChildrenRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    // resolveItemId for "/" returns "root" directly — no root().get() call
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("root") } returns itemBuilder
    every { itemBuilder.children() } returns childrenBuilder
    every { childrenBuilder.post(any()) } returns createdFolder

    val result = client.createFolder("drive-1", "/", "NewFolder")

    assertEquals("folder-new", result.id)
    assertEquals("NewFolder", result.name)
    assertTrue(result.isFolder)
  }

  // ── deleteItem ───────────────────────────────────────────────────────

  @Test
  fun `deleteItem - deletes item by id`() {
    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("item-to-delete") } returns itemBuilder
    every { itemBuilder.delete() } just Runs

    assertDoesNotThrow { client.deleteItem("drive-1", "item-to-delete") }
    verify { itemBuilder.delete() }
  }

  @Test
  fun `deleteItem - graph api error throws RuntimeException`() {
    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("item-x") } returns itemBuilder
    every { itemBuilder.delete() } throws RuntimeException("Not found")

    val ex = assertThrows(RuntimeException::class.java) { client.deleteItem("drive-1", "item-x") }
    assertTrue(ex.message!!.contains("Not found"))
  }

  // ── searchFiles ──────────────────────────────────────────────────────

  @Test
  fun `searchFiles - returns matching files`() {
    val matchingItem = buildDriveItem("found-1", "budget.xlsx", isFolder = false, size = 4096)
    val searchResponse = SearchWithQGetResponse().apply { value = listOf(matchingItem) }

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val searchBuilder = mockk<SearchWithQRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.searchWithQ("budget") } returns searchBuilder
    every { searchBuilder.get() } returns searchResponse

    val result = client.searchFiles("drive-1", "budget")

    assertEquals(1, result.size)
    assertEquals("found-1", result[0].id)
    assertEquals("budget.xlsx", result[0].name)
  }

  // ── moveItem ─────────────────────────────────────────────────────────

  @Test
  fun `moveItem - moves item to new location`() {
    val destItem = DriveItem().apply { id = "dest-folder-id" }
    val movedItem = buildDriveItem("item-1", "report.docx", isFolder = false, size = 1024)

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val rootItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val pathItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val targetItemBuilder = mockk<DriveItemItemRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    // resolveItemId(driveId, "Archive") → items().byDriveItemId("root").withUrl(...).get()
    every { itemsBuilder.byDriveItemId("root") } returns rootItemBuilder
    every { rootItemBuilder.withUrl(any()) } returns pathItemBuilder
    every { pathItemBuilder.get() } returns destItem
    // moveItem → items().byDriveItemId("item-1").patch()
    every { itemsBuilder.byDriveItemId("item-1") } returns targetItemBuilder
    every { targetItemBuilder.patch(any()) } returns movedItem

    val result = client.moveItem("drive-1", "item-1", "Archive")

    assertEquals("item-1", result.id)
    verify { targetItemBuilder.patch(any()) }
  }

  // ── copyItem ─────────────────────────────────────────────────────────

  @Test
  fun `copyItem - copies item to destination`() {
    val destItem = DriveItem().apply { id = "dest-folder-id" }
    val copiedItem = buildDriveItem("item-copy", "report-copy.docx", isFolder = false, size = 1024)

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val rootItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val pathItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val targetItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val copyBuilder = mockk<CopyRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    // resolveItemId
    every { itemsBuilder.byDriveItemId("root") } returns rootItemBuilder
    every { rootItemBuilder.withUrl(any()) } returns pathItemBuilder
    every { pathItemBuilder.get() } returns destItem
    // copyItem
    every { itemsBuilder.byDriveItemId("item-1") } returns targetItemBuilder
    every { targetItemBuilder.copy() } returns copyBuilder
    every { copyBuilder.post(any()) } returns copiedItem

    val result = client.copyItem("drive-1", "item-1", "Backup")

    assertEquals("item-copy", result.id)
  }

  @Test
  fun `copyItem - null result returns placeholder with copy in progress`() {
    val destItem = DriveItem().apply { id = "dest-folder-id" }

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val rootItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val pathItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val targetItemBuilder = mockk<DriveItemItemRequestBuilder>()
    val copyBuilder = mockk<CopyRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    // resolveItemId
    every { itemsBuilder.byDriveItemId("root") } returns rootItemBuilder
    every { rootItemBuilder.withUrl(any()) } returns pathItemBuilder
    every { pathItemBuilder.get() } returns destItem
    // copyItem returns null (async copy)
    every { itemsBuilder.byDriveItemId("item-1") } returns targetItemBuilder
    every { targetItemBuilder.copy() } returns copyBuilder
    every { copyBuilder.post(any()) } returns null

    val result = client.copyItem("drive-1", "item-1", "Backup")

    assertEquals("item-1", result.id)
    assertEquals("(copy in progress)", result.name)
  }

  // ── SharePointConfig validation ─────────────────────────────────────

  // ── uploadFile (max upload bytes enforcement) ────────────────────────

  @Test
  fun `uploadFile - exceeds maxUploadBytes throws IllegalArgumentException`() {
    val restrictedConfig = SharePointConfig(
      tenantId = "test-tenant",
      clientId = "test-client",
      clientSecret = "test-secret",
      maxUploadBytes = 10L // 10-byte hard limit
    )
    val restrictedClient = SharePointGraphClient(restrictedConfig)

    // Replace the lazy graphClient with our mock via reflection
    val field = SharePointGraphClient::class.java.getDeclaredField("graphClient\$delegate")
    field.isAccessible = true
    field.set(restrictedClient, lazy { mockGraphClient })

    val content = "This content is definitely longer than 10 bytes".toByteArray()

    val ex = assertThrows(IllegalArgumentException::class.java) {
      restrictedClient.uploadFile("drive-1", "/", "big-file.txt", content)
    }
    assertTrue(ex.message!!.contains("exceeds maximum allowed 10 bytes"), "Error message should mention limit, got: ${ex.message}")
  }

  @Test
  fun `uploadFile - exactly at maxUploadBytes is allowed`() {
    val content = "1234567890".toByteArray() // exactly 10 bytes
    val restrictedConfig = SharePointConfig(
      tenantId = "test-tenant",
      clientId = "test-client",
      clientSecret = "test-secret",
      maxUploadBytes = 10L
    )
    val restrictedClient = SharePointGraphClient(restrictedConfig)

    val field = SharePointGraphClient::class.java.getDeclaredField("graphClient\$delegate")
    field.isAccessible = true
    field.set(restrictedClient, lazy { mockGraphClient })

    val uploadedItem = buildDriveItem("uploaded-1", "exact.txt", isFolder = false, size = content.size.toLong())

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val contentBuilder = mockk<ContentRequestBuilder>()
    val urlContentBuilder = mockk<ContentRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("root") } returns itemBuilder
    every { itemBuilder.content() } returns contentBuilder
    every { contentBuilder.withUrl(any()) } returns urlContentBuilder
    every { urlContentBuilder.put(any()) } returns uploadedItem

    val result = restrictedClient.uploadFile("drive-1", "/", "exact.txt", content)
    assertEquals("uploaded-1", result.id)
  }

  @Test
  fun `uploadFile - maxUploadBytes zero means unlimited`() {
    // Default config has maxUploadBytes = 0 (unlimited)
    val content = "data".toByteArray()
    val uploadedItem = buildDriveItem("uploaded-1", "test.txt", isFolder = false, size = content.size.toLong())

    val drivesBuilder = mockk<DrivesRequestBuilder>()
    val driveBuilder = mockk<DriveItemRequestBuilder>()
    val itemsBuilder = mockk<ItemsRequestBuilder>()
    val itemBuilder = mockk<DriveItemItemRequestBuilder>()
    val contentBuilder = mockk<ContentRequestBuilder>()
    val urlContentBuilder = mockk<ContentRequestBuilder>()

    every { mockGraphClient.drives() } returns drivesBuilder
    every { drivesBuilder.byDriveId("drive-1") } returns driveBuilder
    every { driveBuilder.items() } returns itemsBuilder
    every { itemsBuilder.byDriveItemId("root") } returns itemBuilder
    every { itemBuilder.content() } returns contentBuilder
    every { contentBuilder.withUrl(any()) } returns urlContentBuilder
    every { urlContentBuilder.put(any()) } returns uploadedItem

    // Should not throw — maxUploadBytes=0 means unlimited
    val result = client.uploadFile("drive-1", "/", "test.txt", content)
    assertEquals("uploaded-1", result.id)
  }

  // ── SharePointConfig validation ─────────────────────────────────────

  @Test
  fun `config - negative maxUploadBytes throws IllegalArgumentException`() {
    assertThrows(IllegalArgumentException::class.java) {
      SharePointConfig("t", "c", "s", maxUploadBytes = -1L)
    }
  }

  @Test
  fun `config - invalid chunk size throws IllegalArgumentException`() {
    assertThrows(IllegalArgumentException::class.java) {
      SharePointConfig("t", "c", "s", uploadChunkBytes = 100_000L)
    }
  }

  @Test
  fun `config - maxSimpleUploadBytes over 4MB throws IllegalArgumentException`() {
    assertThrows(IllegalArgumentException::class.java) {
      SharePointConfig("t", "c", "s", maxSimpleUploadBytes = 5_000_000L)
    }
  }

  @Test
  fun `config - valid custom chunk size accepted`() {
    val config = SharePointConfig("t", "c", "s", uploadChunkBytes = 327_680L * 5)
    assertEquals(327_680L * 5, config.uploadChunkBytes)
  }

  // ── Helper methods ───────────────────────────────────────────────────

  private fun buildDriveItem(id: String, name: String, isFolder: Boolean, size: Long): DriveItem {
    return DriveItem().apply {
      this.id = id
      this.name = name
      this.size = size
      if (isFolder) {
        this.folder = Folder()
      } else {
        this.file = File().apply { mimeType = "application/octet-stream" }
      }
      this.webUrl = "https://contoso.sharepoint.com/items/$id"
      this.lastModifiedDateTime = OffsetDateTime.now()
      this.createdBy = IdentitySet().apply {
        user = Identity().apply { displayName = "Test User" }
      }
    }
  }
}
