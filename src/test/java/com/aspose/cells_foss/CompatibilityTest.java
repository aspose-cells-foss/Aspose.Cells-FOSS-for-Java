package com.aspose.cells_foss;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Compatibility tests for the cells-foss Java API.
 * Mirrors C# Aspose.Cells_FOSS.CompatibilityTests.
 * Uses Apache POI (XSSFWorkbook) for XLSX output verification instead of OpenXML SDK.
 */
class CompatibilityTest {

    // -------------------------------------------------------------------------
    // Cell API tests
    // -------------------------------------------------------------------------

    /**
     * Verifies that formula setter accepts with or without leading equal.
     */
    @Test
    void formulaSetterAcceptsWithOrWithoutLeadingEqual() {
        Workbook workbook = new Workbook();
        Cell cell = workbook.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(10);
        cell.setFormula("B1+C1");
        assertEquals("=B1+C1", cell.getFormula());

        cell.setFormula("=D1+E1");
        assertEquals("=D1+E1", cell.getFormula());
    }

    /**
     * Verifies that value property setter matches supported scalar behavior.
     */
    @Test
    void valuePropertySetterMatchesSupportedScalarBehavior() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);
        sheet.getCells().get("A1").setValue("alpha");
        sheet.getCells().get("B1").setValue(12);
        sheet.getCells().get("C1").setValue(true);
        sheet.getCells().get("D1").setValue(LocalDateTime.of(2024, 1, 2, 3, 4, 0));
        sheet.getCells().get("E1").setValue(null);

        assertEquals("alpha", sheet.getCells().get("A1").getValue());
        assertEquals(12, sheet.getCells().get("B1").getValue());
        assertEquals(true, sheet.getCells().get("C1").getValue());
        assertEquals(CellValueType.DATE_TIME, sheet.getCells().get("D1").getType());
        assertEquals("", sheet.getCells().get("E1").getDisplayStringValue());
    }

    /**
     * Verifies that public type mapping matches cell value types.
     */
    @Test
    void publicTypeMappingMatchesCellValueTypes() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);
        sheet.getCells().get("A1").putValue("Hello");
        sheet.getCells().get("B1").putValue(123);
        sheet.getCells().get("C1").putValue(true);
        sheet.getCells().get("D1").putValue(12.5d);
        sheet.getCells().get("F1").putValue(LocalDateTime.of(2024, 5, 6, 7, 8, 9));
        sheet.getCells().get("G1").setFormula("=B1*2");

        assertEquals(CellValueType.STRING, sheet.getCells().get("A1").getType());
        assertEquals(CellValueType.NUMBER, sheet.getCells().get("B1").getType());
        assertEquals(CellValueType.BOOLEAN, sheet.getCells().get("C1").getType());
        assertEquals(CellValueType.NUMBER, sheet.getCells().get("D1").getType());
        assertEquals(CellValueType.BLANK, sheet.getCells().get("E1").getType());
        assertEquals(CellValueType.DATE_TIME, sheet.getCells().get("F1").getType());
        assertEquals(CellValueType.FORMULA, sheet.getCells().get("G1").getType());
    }

    /**
     * Verifies that string value formats correctly.
     */
    @Test
    void stringValueFormatsCorrectly() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);
        sheet.getCells().get("A1").putValue("Hello");
        sheet.getCells().get("B1").putValue(123);
        sheet.getCells().get("C1").putValue(true);

        assertEquals("Hello", sheet.getCells().get("A1").getStringValue());
        assertEquals("123", sheet.getCells().get("B1").getStringValue());
        assertEquals("TRUE", sheet.getCells().get("C1").getStringValue());
        assertEquals("", sheet.getCells().get("D1").getStringValue());
    }

    // -------------------------------------------------------------------------
    // Workbook settings tests
    // -------------------------------------------------------------------------

    /**
     * Verifies that workbook date 1904 setting round trips.
     */
    @Test
    void workbookDate1904SettingRoundTrips() {
        Workbook workbook = new Workbook();
        assertFalse(workbook.getSettings().getDate1904());
        workbook.getSettings().setDate1904(true);
        assertTrue(workbook.getSettings().getDate1904());
        workbook.getSettings().setDate1904(false);
        assertFalse(workbook.getSettings().getDate1904());
    }

    // -------------------------------------------------------------------------
    // Worksheet collection tests
    // -------------------------------------------------------------------------

    /**
     * Verifies that worksheet collection add and get.
     */
    @Test
    void worksheetCollectionAddAndGet() {
        Workbook workbook = new Workbook();
        assertEquals(1, workbook.getWorksheets().getCount());

        int idx = workbook.getWorksheets().add();
        assertEquals(2, workbook.getWorksheets().getCount());

        workbook.getWorksheets().get(idx).setName("Report");
        assertEquals("Report", workbook.getWorksheets().get(idx).getName());
        assertEquals("Report", workbook.getWorksheets().get("Report").getName());
    }

    /**
     * Verifies that worksheet collection active sheet name.
     */
    @Test
    void worksheetCollectionActiveSheetName() {
        Workbook workbook = new Workbook();
        int idx = workbook.getWorksheets().add();
        workbook.getWorksheets().get(idx).setName("Report");
        workbook.getWorksheets().setActiveSheetName("Report");
        assertEquals(idx, workbook.getWorksheets().getActiveSheetIndex());
        assertEquals("Report", workbook.getWorksheets().getActiveSheetName());
    }

    /**
     * Verifies that worksheet not found throws cells exception.
     */
    @Test
    void worksheetNotFoundThrowsCellsException() {
        Workbook workbook = new Workbook();
        assertThrows(CellsException.class, () -> workbook.getWorksheets().get("missing"));
    }

    // -------------------------------------------------------------------------
    // Exception mapping tests
    // -------------------------------------------------------------------------

    /**
     * Verifies that exception mapping uses cells exception types.
     */
    @Test
    void exceptionMappingUsesCellsExceptionTypes() {
        assertThrows(CellsException.class, () -> new Workbook().getWorksheets().get("missing"));
        assertThrows(CellsException.class, () -> new Workbook().getWorksheets().get(0).getCells().get("1A"));
        assertThrows(InvalidFileFormatException.class,
                () -> new Workbook(new ByteArrayInputStream(new byte[]{1, 2, 3, 4})));
    }

    // -------------------------------------------------------------------------
    // Worksheet settings tests
    // -------------------------------------------------------------------------

    /**
     * Verifies that worksheet view members follow supported patterns.
     */
    @Test
    void worksheetViewMembersFollowSupportedPatterns() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);

        sheet.setTabColor(Color.fromArgb(255, 34, 68, 102));
        sheet.setShowGridlines(false);
        sheet.setShowRowColumnHeaders(false);
        sheet.setShowZeros(false);
        sheet.setRightToLeft(true);
        sheet.setZoom(85);

        assertEquals(Color.fromArgb(255, 34, 68, 102), sheet.getTabColor());
        assertFalse(sheet.getShowGridlines());
        assertFalse(sheet.getShowRowColumnHeaders());
        assertFalse(sheet.getShowZeros());
        assertTrue(sheet.getRightToLeft());
        assertEquals(85, sheet.getZoom());
    }

    /**
     * Verifies that worksheet protection members follow supported patterns.
     */
    @Test
    void worksheetProtectionMembersFollowSupportedPatterns() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);

        sheet.protect();
        sheet.getProtection().setAllowEditingObject(false);
        sheet.getProtection().setAllowEditingScenario(false);
        sheet.getProtection().setAllowFiltering(false);
        sheet.getProtection().setAllowSelectingLockedCell(false);
        sheet.getProtection().setAllowSelectingUnlockedCell(false);

        assertTrue(sheet.getProtection().isProtected());
        assertFalse(sheet.getProtection().getAllowEditingObject());
        assertFalse(sheet.getProtection().getAllowEditingScenario());
        assertFalse(sheet.getProtection().getAllowFiltering());
        assertFalse(sheet.getProtection().getAllowSelectingLockedCell());
        assertFalse(sheet.getProtection().getAllowSelectingUnlockedCell());
    }

    /**
     * Verifies that worksheet settings scenario in memory.
     */
    @Test
    void worksheetSettingsScenarioInMemory() {
        Workbook workbook = WorksheetScenarioFactory.createWorksheetSettingsWorkbook();
        WorksheetScenarioFactory.assertWorksheetSettings(workbook);
        WorksheetScenarioFactory.assertWorksheetSettingsScenarioHasVisibleSheet(workbook);
    }

    // -------------------------------------------------------------------------
    // PageSetup tests
    // -------------------------------------------------------------------------

    /**
     * Verifies that page setup members round trip in memory.
     */
    @Test
    void pageSetupMembersRoundTripInMemory() {
        Workbook workbook = PageSetupScenarioFactory.createPageSetupWorkbook();
        PageSetupScenarioFactory.assertPageSetup(workbook);
    }

    // -------------------------------------------------------------------------
    // Hyperlink tests
    // -------------------------------------------------------------------------

    /**
     * Verifies that hyperlink add and retrieve.
     */
    @Test
    void hyperlinkAddAndRetrieve() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);
        sheet.getCells().get("A1").putValue("Docs");

        int idx = sheet.getHyperlinks().add("A1", 1, 1, "https://example.com/docs");
        Hyperlink link = sheet.getHyperlinks().get(idx);
        link.setTextToDisplay("Docs");
        link.setScreenTip("External docs");

        assertEquals(1, sheet.getHyperlinks().getCount());
        assertEquals("https://example.com/docs", sheet.getHyperlinks().get(0).getAddress());
        assertEquals("External docs", sheet.getHyperlinks().get(0).getScreenTip());
        assertEquals("Docs", sheet.getHyperlinks().get(0).getTextToDisplay());
        assertEquals(TargetModeType.EXTERNAL, sheet.getHyperlinks().get(0).getLinkType());
    }

    /**
     * Verifies that hyperlink delete reduces count.
     */
    @Test
    void hyperlinkDeleteReducesCount() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);
        int idx = sheet.getHyperlinks().add(0, 0, 1, 1, "https://example.com/docs");
        assertEquals(1, sheet.getHyperlinks().getCount());
        sheet.getHyperlinks().get(idx).delete();
        assertEquals(0, sheet.getHyperlinks().getCount());
    }

    // -------------------------------------------------------------------------
    // AutoFilter in-memory tests
    // -------------------------------------------------------------------------

    /**
     * Verifies that auto filter members follow supported patterns.
     */
    @Test
    void autoFilterMembersFollowSupportedPatterns() {
        Workbook workbook = AutoFilterScenarioFactory.createAutoFilterWorkbook();
        AutoFilterScenarioFactory.assertAutoFilter(workbook);
    }

    // -------------------------------------------------------------------------
    // DefinedName tests
    // -------------------------------------------------------------------------

    /**
     * Verifies that defined name add and retrieve.
     */
    @Test
    void definedNameAddAndRetrieve() {
        Workbook workbook = new Workbook();
        workbook.getWorksheets().get(0).setName("Data");

        int idx = workbook.getDefinedNames().add("MyRange", "Data!$A$1:$B$10");
        assertEquals(1, workbook.getDefinedNames().getCount());
        assertEquals("MyRange", workbook.getDefinedNames().get(idx).getName());
        assertEquals("Data!$A$1:$B$10", workbook.getDefinedNames().get(idx).getFormula());
    }

    // -------------------------------------------------------------------------
    // XLSX save tests (verified via Apache POI)
    // -------------------------------------------------------------------------

    /**
     * Verifies that save then read with poi returns correct cell values.
     */
    @Test
    void saveThenReadWithPoiReturnsCorrectCellValues() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CompatibilityTest")) {
            String path = tempDir.getPath("basic.xlsx");

            try (Workbook workbook = new Workbook()) {
            Worksheet sheet = workbook.getWorksheets().get(0);
            sheet.setName("Data");
            sheet.getCells().get("A1").putValue("Hello");
            sheet.getCells().get("B1").putValue(42);
            sheet.getCells().get("C1").putValue(true);
            workbook.save(path);
            }

            // Verify with Apache POI
            try (org.apache.poi.ss.usermodel.Workbook poiWorkbook = WorkbookFactory.create(new File(path))) {
                org.apache.poi.ss.usermodel.Sheet poiSheet = poiWorkbook.getSheet("Data");
                assertNotNull(poiSheet, "Sheet 'Data' must exist in saved XLSX");
                org.apache.poi.ss.usermodel.Row row = poiSheet.getRow(0);
                assertNotNull(row);
                assertEquals("Hello", row.getCell(0).getStringCellValue());
                assertEquals(42.0, row.getCell(1).getNumericCellValue(), 0.001);
                assertTrue(row.getCell(2).getBooleanCellValue());
            }
        }
    }

    /**
     * Verifies that save to stream then reload produces same values.
     */
    @Test
    void saveToStreamThenReloadProducesSameValues() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CompatibilityTest")) {
            String path = tempDir.getPath("roundtrip.xlsx");
            Workbook original = WorkbookScenarioFactory.createMixedCellWorkbook();
            original.save(path, SaveFormat.XLSX);

            try (InputStream stream = new FileInputStream(path)) {
                Workbook loaded = new Workbook(stream);
                WorkbookScenarioFactory.assertWorkbookDataEquals(original, loaded);
            }
        }
    }

    /**
     * Verifies that save file and stream produce equivalent workbooks.
     */
    @Test
    void saveFileAndStreamProduceEquivalentWorkbooks() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CompatibilityTest")) {
            String filePath = tempDir.getPath("file.xlsx");
            Workbook workbook = WorkbookScenarioFactory.createMixedCellWorkbook();
            workbook.save(filePath, SaveFormat.XLSX);

            ByteArrayOutputStream streamBuffer = new ByteArrayOutputStream();
            workbook.save(streamBuffer, SaveFormat.XLSX);

            Workbook fromFile = new Workbook(filePath);
            Workbook fromStream = new Workbook(new ByteArrayInputStream(streamBuffer.toByteArray()));
            WorkbookScenarioFactory.assertWorkbookDataEquals(fromFile, fromStream);
        }
    }

    /**
     * Verifies that worksheet settings round trip via file.
     */
    @Test
    void worksheetSettingsRoundTripViaFile() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CompatibilityTest")) {
            String path = tempDir.getPath("ws-settings.xlsx");
            Workbook workbook = WorksheetScenarioFactory.createWorksheetSettingsWorkbook();
            workbook.save(path);

            Workbook loaded = new Workbook(path);
            WorksheetScenarioFactory.assertWorksheetSettings(loaded);
            WorksheetScenarioFactory.assertWorksheetSettingsScenarioHasVisibleSheet(loaded);
        }
    }

    /**
     * Verifies that page setup round trip via file.
     */
    @Test
    void pageSetupRoundTripViaFile() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CompatibilityTest")) {
            String path = tempDir.getPath("page-setup.xlsx");
            Workbook workbook = PageSetupScenarioFactory.createPageSetupWorkbook();
            workbook.save(path);

            Workbook loaded = new Workbook(path);
            PageSetupScenarioFactory.assertPageSetup(loaded);
        }
    }

    /**
     * Verifies that auto filter round trip via file.
     */
    @Test
    void autoFilterRoundTripViaFile() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CompatibilityTest")) {
            String path = tempDir.getPath("autofilter.xlsx");
            Workbook workbook = AutoFilterScenarioFactory.createAutoFilterWorkbook();
            workbook.save(path);

            Workbook loaded = new Workbook(path);
            AutoFilterScenarioFactory.assertAutoFilter(loaded);
        }
    }

    /**
     * Verifies that worksheet xml contains auto filter element.
     */
    @Test
    void worksheetXmlContainsAutoFilterElement() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CompatibilityTest")) {
            String path = tempDir.getPath("autofilter-xml.xlsx");
            try (Workbook workbook = new Workbook()) {
                workbook.getWorksheets().get(0).getAutoFilter().setRange("A1:C5");
                workbook.save(path);
            }

            String worksheetXml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(worksheetXml.contains("autoFilter"),
                    "worksheet XML should contain <autoFilter> element");
        }
    }

    /**
     * Verifies that workbook xml contains sheet references.
     */
    @Test
    void workbookXmlContainsSheetReferences() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CompatibilityTest")) {
            String path = tempDir.getPath("multi-sheet.xlsx");
            try (Workbook workbook = new Workbook()) {
                workbook.getWorksheets().get(0).setName("Alpha");
                workbook.getWorksheets().add("Beta");
                workbook.save(path);
            }

            String workbookXml = ZipPackageHelper.readEntryText(path, "xl/workbook.xml");
            assertTrue(workbookXml.contains("Alpha"), "workbook.xml must reference sheet 'Alpha'");
            assertTrue(workbookXml.contains("Beta"), "workbook.xml must reference sheet 'Beta'");
        }
    }
}
