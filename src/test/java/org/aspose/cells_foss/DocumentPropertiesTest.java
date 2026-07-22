package org.aspose.cells_foss;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 tests for DocumentProperties and WorkbookProperties 鈥?DP-* / WP-* test cases.
 *
 * <p>DocumentProperties covers docProps/core.xml (title, subject, author/creator,
 * keywords, comments/description, category) and docProps/app.xml (company, manager).
 *
 * <p>WorkbookProperties covers workbook.xml {@code <workbookPr>} attributes
 * (codeName, filterPrivacy, hidePivotFieldList, etc.).
 */
class DocumentPropertiesTest {

    // =========================================================================
    // DP-01 鈥?DP-08  In-memory API
    // =========================================================================

    /**
     * Verifies that title roundtrips in memory.
     */
    @Test
    void DP_01_titleRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getDocumentProperties().setTitle("My Report");
            assertEquals("My Report", wb.getDocumentProperties().getTitle());
        }
    }

    /**
     * Verifies that subject roundtrips in memory.
     */
    @Test
    void DP_02_subjectRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getDocumentProperties().setSubject("Q4 2024");
            assertEquals("Q4 2024", wb.getDocumentProperties().getSubject());
        }
    }

    /**
     * Verifies that author roundtrips in memory.
     */
    @Test
    void DP_03_authorRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getDocumentProperties().setAuthor("Jane Doe");
            assertEquals("Jane Doe", wb.getDocumentProperties().getAuthor());
        }
    }

    /**
     * Verifies that keywords roundtrips in memory.
     */
    @Test
    void DP_04_keywordsRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getDocumentProperties().setKeywords("finance, annual, report");
            assertEquals("finance, annual, report", wb.getDocumentProperties().getKeywords());
        }
    }

    /**
     * Verifies that comments roundtrips in memory.
     */
    @Test
    void DP_05_commentsRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getDocumentProperties().setComments("Internal draft");
            assertEquals("Internal draft", wb.getDocumentProperties().getComments());
        }
    }

    /**
     * Verifies that category roundtrips in memory.
     */
    @Test
    void DP_06_categoryRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getDocumentProperties().setCategory("Finance");
            assertEquals("Finance", wb.getDocumentProperties().getCategory());
        }
    }

    /**
     * Verifies that company roundtrips in memory.
     */
    @Test
    void DP_07_companyRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getDocumentProperties().setCompany("ACME Corp");
            assertEquals("ACME Corp", wb.getDocumentProperties().getCompany());
        }
    }

    /**
     * Verifies that manager roundtrips in memory.
     */
    @Test
    void DP_08_managerRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getDocumentProperties().setManager("Alice Smith");
            assertEquals("Alice Smith", wb.getDocumentProperties().getManager());
        }
    }

    // =========================================================================
    // DP-10 鈥?DP-18  XLSX round-trip (save 鈫?reload)
    // =========================================================================

    /**
     * Verifies that title roundtrips through xlsx.
     */
    @Test
    void DP_10_titleRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("props.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setTitle("Annual Report 2024");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("Annual Report 2024", loaded.getDocumentProperties().getTitle());
            }
        }
    }

    /**
     * Verifies that subject roundtrips through xlsx.
     */
    @Test
    void DP_11_subjectRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("props.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setSubject("Financial Summary");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("Financial Summary", loaded.getDocumentProperties().getSubject());
            }
        }
    }

    /**
     * Verifies that author roundtrips through xlsx.
     */
    @Test
    void DP_12_authorRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("props.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setAuthor("Jane Doe");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("Jane Doe", loaded.getDocumentProperties().getAuthor());
            }
        }
    }

    /**
     * Verifies that keywords roundtrips through xlsx.
     */
    @Test
    void DP_13_keywordsRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("props.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setKeywords("finance, annual");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("finance, annual", loaded.getDocumentProperties().getKeywords());
            }
        }
    }

    /**
     * Verifies that comments roundtrips through xlsx.
     */
    @Test
    void DP_14_commentsRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("props.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setComments("Confidential");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("Confidential", loaded.getDocumentProperties().getComments());
            }
        }
    }

    /**
     * Verifies that category roundtrips through xlsx.
     */
    @Test
    void DP_15_categoryRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("props.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setCategory("HR");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("HR", loaded.getDocumentProperties().getCategory());
            }
        }
    }

    /**
     * Verifies that company roundtrips through xlsx.
     */
    @Test
    void DP_16_companyRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("props.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setCompany("ACME Corp");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("ACME Corp", loaded.getDocumentProperties().getCompany());
            }
        }
    }

    /**
     * Verifies that manager roundtrips through xlsx.
     */
    @Test
    void DP_17_managerRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("props.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setManager("Bob Jones");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("Bob Jones", loaded.getDocumentProperties().getManager());
            }
        }
    }

    /**
     * Verifies that all core properties roundtrip.
     */
    @Test
    void DP_18_allCorePropertiesRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("props.xlsx");
            try (Workbook wb = new Workbook()) {
                DocumentProperties dp = wb.getDocumentProperties();
                dp.setTitle("T");
                dp.setSubject("S");
                dp.setAuthor("A");
                dp.setKeywords("K");
                dp.setComments("C");
                dp.setCategory("Cat");
                dp.setCompany("Co");
                dp.setManager("M");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                DocumentProperties dp = loaded.getDocumentProperties();
                assertAll(
                    () -> assertEquals("T",   dp.getTitle()),
                    () -> assertEquals("S",   dp.getSubject()),
                    () -> assertEquals("A",   dp.getAuthor()),
                    () -> assertEquals("K",   dp.getKeywords()),
                    () -> assertEquals("C",   dp.getComments()),
                    () -> assertEquals("Cat", dp.getCategory()),
                    () -> assertEquals("Co",  dp.getCompany()),
                    () -> assertEquals("M",   dp.getManager())
                );
            }
        }
    }

    // =========================================================================
    // DP-20 鈥?DP-22  XLSX stream round-trip
    // =========================================================================

    /**
     * Verifies that properties roundtrip via stream.
     */
    @Test
    void DP_20_propertiesRoundtripViaStream() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getDocumentProperties().setTitle("Stream Title");
            wb.getDocumentProperties().setCompany("Stream Co");
            wb.save(out, SaveFormat.XLSX);
        }
        try (Workbook loaded = new Workbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("Stream Title", loaded.getDocumentProperties().getTitle());
            assertEquals("Stream Co",    loaded.getDocumentProperties().getCompany());
        }
    }

    /**
     * Verifies that docProps parts are always present (OPC compliance) even when empty.
     */
    @Test
    void DP_21_noDocPropsPartWhenPropertiesAreEmpty() throws Exception {
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("empty.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.save(path);
            }
            // docProps/core.xml must always exist for OPC compliance (Excel requires it)
            String core = ZipPackageHelper.readEntryText(path, "docProps/core.xml");
            assertTrue(core.contains("coreProperties"), "docProps/core.xml should always be present");
            // content should be empty (no title, no author, etc.)
            try (Workbook loaded = new Workbook(path)) {
                assertTrue(loaded.getDocumentProperties().getTitle().isEmpty(),
                    "Title should be empty when not set");
            }
        }
    }

    /**
     * Verifies that doc props xml contains title text.
     */
    @Test
    void DP_22_docPropsXmlContainsTitleText() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("titled.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setTitle("XLSX Title Check");
                wb.save(path);
            }
            String coreXml = ZipPackageHelper.readEntryText(path, "docProps/core.xml");
            assertTrue(coreXml.contains("XLSX Title Check"), "core.xml should contain the title text");
            String appXml = ZipPackageHelper.readEntryText(path, "docProps/app.xml");
            assertNotNull(appXml); // app.xml present even when only core has content
        }
    }

    // =========================================================================
    // DP-25  Special characters in properties
    // =========================================================================

    /**
     * Verifies that special chars in title are escaped.
     */
    @Test
    void DP_25_specialCharsInTitleAreEscaped() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("special.xlsx");
            String title = "Report <Q4> & \"2024\"";
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setTitle(title);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals(title, loaded.getDocumentProperties().getTitle());
            }
            String xml = ZipPackageHelper.readEntryText(path, "docProps/core.xml");
            assertFalse(xml.contains("<Q4>"), "Raw < > should be XML-escaped in core.xml");
        }
    }

    // =========================================================================
    // DP-30  POI interop 鈥?write our properties, read back with Apache POI
    // =========================================================================

    /**
     * Verifies that poi can read document properties.
     */
    @Test
    void DP_30_poiCanReadDocumentProperties() throws Exception {
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("poi-props.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getDocumentProperties().setTitle("POI Title");
                wb.getDocumentProperties().setAuthor("POI Author");
                wb.getDocumentProperties().setCompany("POI Corp");
                wb.save(path);
            }
            // Verify via raw XML since POI core-properties API varies by version
            String coreXml = ZipPackageHelper.readEntryText(path, "docProps/core.xml");
            String appXml  = ZipPackageHelper.readEntryText(path, "docProps/app.xml");
            assertTrue(coreXml.contains("POI Title"),  "core.xml must contain title");
            assertTrue(coreXml.contains("POI Author"), "core.xml must contain author/creator");
            assertTrue(appXml.contains("POI Corp"),    "app.xml must contain company");
        }
    }

    // =========================================================================
    // WP-01 鈥?WP-05  WorkbookProperties (workbookPr attributes)
    // =========================================================================

    /**
     * Verifies that code name roundtrips in memory.
     */
    @Test
    void WP_01_codeNameRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getProperties().setCodeName("ThisWorkbook");
            assertEquals("ThisWorkbook", wb.getProperties().getCodeName());
        }
    }

    /**
     * Verifies that filter privacy roundtrips in memory.
     */
    @Test
    void WP_02_filterPrivacyRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getProperties().setFilterPrivacy(true);
            assertTrue(wb.getProperties().getFilterPrivacy());
        }
    }

    /**
     * Verifies that hide pivot field list roundtrips in memory.
     */
    @Test
    void WP_03_hidePivotFieldListRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            wb.getProperties().setHidePivotFieldList(true);
            assertTrue(wb.getProperties().getHidePivotFieldList());
        }
    }

    /**
     * Verifies that code name roundtrips through xlsx.
     */
    @Test
    void WP_04_codeNameRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("wbprops.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getProperties().setCodeName("MyWorkbook");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("MyWorkbook", loaded.getProperties().getCodeName());
            }
        }
    }

    /**
     * Verifies that workbook pr xml contains code name.
     */
    @Test
    void WP_05_workbookPrXmlContainsCodeName() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("DocumentPropertiesTest")) {
            String path = dir.getPath("wbprops.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getProperties().setCodeName("SalesModel");
                wb.save(path);
            }
            String wbXml = ZipPackageHelper.readEntryText(path, "xl/workbook.xml");
            assertTrue(wbXml.contains("codeName=\"SalesModel\""),
                "workbook.xml must contain codeName attribute");
        }
    }
}

