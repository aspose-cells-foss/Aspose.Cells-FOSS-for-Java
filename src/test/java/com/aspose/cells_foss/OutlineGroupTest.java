package com.aspose.cells_foss;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 tests for row/column outline (grouping) â€?OG-* test cases.
 *
 * <p>Outline levels map to the OOXML {@code outlineLevel} attribute on {@code <row>} and
 * {@code <col>} elements. Level 0 means ungrouped; levels 1â€? represent nesting depth.
 * The {@code collapsed} flag indicates whether the outline group is currently collapsed.
 */
class OutlineGroupTest {

    // =========================================================================
    // OG-01 â€?OG-06  Row grouping â€?in-memory API
    // =========================================================================

    /**
     * Verifies that row group level defaults to zero.
     */
    @Test
    void OG_01_rowGroupLevelDefaultsToZero() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Row row = wb.getWorksheets().get(0).getCells().getRows().get(5);
            assertEquals(0, row.getGroupLevel());
        }
    }

    /**
     * Verifies that row group level roundtrips in memory.
     */
    @Test
    void OG_02_rowGroupLevelRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Row row = wb.getWorksheets().get(0).getCells().getRows().get(3);
            row.setGroupLevel(2);
            assertEquals(2, row.getGroupLevel());
        }
    }

    /**
     * Verifies that row collapsed defaults false.
     */
    @Test
    void OG_03_rowCollapsedDefaultsFalse() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            assertFalse(wb.getWorksheets().get(0).getCells().getRows().get(0).isCollapsed());
        }
    }

    /**
     * Verifies that row collapsed roundtrips in memory.
     */
    @Test
    void OG_04_rowCollapsedRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Row row = wb.getWorksheets().get(0).getCells().getRows().get(1);
            row.setCollapsed(true);
            assertTrue(row.isCollapsed());
        }
    }

    /**
     * Verifies that row group level boundary accepted.
     */
    @Test
    void OG_05_rowGroupLevelBoundaryAccepted() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getCells().getRows().get(0).setGroupLevel(1);
            ws.getCells().getRows().get(1).setGroupLevel(7);
            assertEquals(1, ws.getCells().getRows().get(0).getGroupLevel());
            assertEquals(7, ws.getCells().getRows().get(1).getGroupLevel());
        }
    }

    /**
     * Verifies that row group level out of range throws.
     */
    @Test
    void OG_06_rowGroupLevelOutOfRangeThrows() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Row row = wb.getWorksheets().get(0).getCells().getRows().get(0);
            assertThrows(CellsException.class, () -> row.setGroupLevel(-1));
            assertThrows(CellsException.class, () -> row.setGroupLevel(8));
        }
    }

    // =========================================================================
    // OG-10 â€?OG-15  Column grouping â€?in-memory API
    // =========================================================================

    /**
     * Verifies that column group level defaults to zero.
     */
    @Test
    void OG_10_columnGroupLevelDefaultsToZero() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Column col = wb.getWorksheets().get(0).getCells().getColumns().get(2);
            assertEquals(0, col.getGroupLevel());
        }
    }

    /**
     * Verifies that column group level roundtrips in memory.
     */
    @Test
    void OG_11_columnGroupLevelRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Column col = wb.getWorksheets().get(0).getCells().getColumns().get(0);
            col.setGroupLevel(3);
            assertEquals(3, col.getGroupLevel());
        }
    }

    /**
     * Verifies that column collapsed defaults false.
     */
    @Test
    void OG_12_columnCollapsedDefaultsFalse() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            assertFalse(wb.getWorksheets().get(0).getCells().getColumns().get(0).isCollapsed());
        }
    }

    /**
     * Verifies that column collapsed roundtrips in memory.
     */
    @Test
    void OG_13_columnCollapsedRoundtripsInMemory() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Column col = wb.getWorksheets().get(0).getCells().getColumns().get(1);
            col.setCollapsed(true);
            assertTrue(col.isCollapsed());
        }
    }

    /**
     * Verifies that column group level boundary accepted.
     */
    @Test
    void OG_14_columnGroupLevelBoundaryAccepted() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getCells().getColumns().get(0).setGroupLevel(1);
            ws.getCells().getColumns().get(1).setGroupLevel(7);
            assertEquals(1, ws.getCells().getColumns().get(0).getGroupLevel());
            assertEquals(7, ws.getCells().getColumns().get(1).getGroupLevel());
        }
    }

    /**
     * Verifies that column group level out of range throws.
     */
    @Test
    void OG_15_columnGroupLevelOutOfRangeThrows() {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Column col = wb.getWorksheets().get(0).getCells().getColumns().get(0);
            assertThrows(CellsException.class, () -> col.setGroupLevel(-1));
            assertThrows(CellsException.class, () -> col.setGroupLevel(8));
        }
    }

    // =========================================================================
    // OG-20 â€?OG-27  XLSX round-trips
    // =========================================================================

    /**
     * Verifies that row group level roundtrips through xlsx.
     */
    @Test
    void OG_20_rowGroupLevelRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("OutlineGroupTest")) {
            String path = dir.getPath("outline.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().getRows().get(1).setGroupLevel(1);
                ws.getCells().getRows().get(2).setGroupLevel(1);
                ws.getCells().getRows().get(3).setGroupLevel(2);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                Worksheet ws = loaded.getWorksheets().get(0);
                assertEquals(1, ws.getCells().getRows().get(1).getGroupLevel());
                assertEquals(1, ws.getCells().getRows().get(2).getGroupLevel());
                assertEquals(2, ws.getCells().getRows().get(3).getGroupLevel());
                assertEquals(0, ws.getCells().getRows().get(0).getGroupLevel());
            }
        }
    }

    /**
     * Verifies that row collapsed roundtrips through xlsx.
     */
    @Test
    void OG_21_rowCollapsedRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("OutlineGroupTest")) {
            String path = dir.getPath("outline.xlsx");
            try (Workbook wb = new Workbook()) {
                Row row = wb.getWorksheets().get(0).getCells().getRows().get(2);
                row.setGroupLevel(1);
                row.setCollapsed(true);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                Row row = loaded.getWorksheets().get(0).getCells().getRows().get(2);
                assertEquals(1, row.getGroupLevel());
                assertTrue(row.isCollapsed());
            }
        }
    }

    /**
     * Verifies that column group level roundtrips through xlsx.
     */
    @Test
    void OG_22_columnGroupLevelRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("OutlineGroupTest")) {
            String path = dir.getPath("outline.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().getColumns().get(1).setGroupLevel(1);
                ws.getCells().getColumns().get(2).setGroupLevel(1);
                ws.getCells().getColumns().get(3).setGroupLevel(2);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                Worksheet ws = loaded.getWorksheets().get(0);
                assertEquals(1, ws.getCells().getColumns().get(1).getGroupLevel());
                assertEquals(1, ws.getCells().getColumns().get(2).getGroupLevel());
                assertEquals(2, ws.getCells().getColumns().get(3).getGroupLevel());
                assertEquals(0, ws.getCells().getColumns().get(0).getGroupLevel());
            }
        }
    }

    /**
     * Verifies that column collapsed roundtrips through xlsx.
     */
    @Test
    void OG_23_columnCollapsedRoundtripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("OutlineGroupTest")) {
            String path = dir.getPath("outline.xlsx");
            try (Workbook wb = new Workbook()) {
                Column col = wb.getWorksheets().get(0).getCells().getColumns().get(2);
                col.setGroupLevel(1);
                col.setCollapsed(true);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                Column col = loaded.getWorksheets().get(0).getCells().getColumns().get(2);
                assertEquals(1, col.getGroupLevel());
                assertTrue(col.isCollapsed());
            }
        }
    }

    /**
     * Verifies that row without group level has no outline level in xml.
     */
    @Test
    void OG_24_rowWithoutGroupLevelHasNoOutlineLevelInXml() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("OutlineGroupTest")) {
            String path = dir.getPath("nolevel.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().get(0, 0).putValue("A");
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertFalse(xml.contains("outlineLevel"), "No outlineLevel attr when all rows are at level 0");
        }
    }

    /**
     * Verifies that row group level appears in worksheet xml.
     */
    @Test
    void OG_25_rowGroupLevelAppearsInWorksheetXml() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("OutlineGroupTest")) {
            String path = dir.getPath("outline.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().getRows().get(1).setGroupLevel(2);
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(xml.contains("outlineLevel=\"2\""), "worksheet XML must contain outlineLevel=\"2\"");
        }
    }

    /**
     * Verifies that column group level appears in worksheet xml.
     */
    @Test
    void OG_26_columnGroupLevelAppearsInWorksheetXml() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("OutlineGroupTest")) {
            String path = dir.getPath("outline.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().getColumns().get(3).setGroupLevel(1);
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(xml.contains("outlineLevel=\"1\""), "worksheet XML must contain outlineLevel=\"1\" for column");
        }
    }

    /**
     * Verifies that multi level nested group roundtrips.
     */
    @Test
    void OG_27_multiLevelNestedGroupRoundtrips() throws Exception {
        try (TemporaryDirectory dir = new TemporaryDirectory("OutlineGroupTest")) {
            String path = dir.getPath("nested.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                // outer group (level 1) rows 1â€?, inner group (level 2) rows 2â€?
                for (int r = 1; r <= 5; r++) ws.getCells().getRows().get(r).setGroupLevel(1);
                for (int r = 2; r <= 3; r++) ws.getCells().getRows().get(r).setGroupLevel(2);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                Worksheet ws = loaded.getWorksheets().get(0);
                assertEquals(0, ws.getCells().getRows().get(0).getGroupLevel());
                assertEquals(1, ws.getCells().getRows().get(1).getGroupLevel());
                assertEquals(2, ws.getCells().getRows().get(2).getGroupLevel());
                assertEquals(2, ws.getCells().getRows().get(3).getGroupLevel());
                assertEquals(1, ws.getCells().getRows().get(4).getGroupLevel());
                assertEquals(1, ws.getCells().getRows().get(5).getGroupLevel());
            }
        }
    }

    // =========================================================================
    // OG-30  Row and column grouping coexist on same sheet
    // =========================================================================

    /**
     * Verifies that row and column grouping coexist.
     */
    @Test
    void OG_30_rowAndColumnGroupingCoexist() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory dir = new TemporaryDirectory("OutlineGroupTest")) {
            String path = dir.getPath("mixed.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().getRows().get(2).setGroupLevel(1);
                ws.getCells().getColumns().get(2).setGroupLevel(1);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                Worksheet ws = loaded.getWorksheets().get(0);
                assertEquals(1, ws.getCells().getRows().get(2).getGroupLevel());
                assertEquals(1, ws.getCells().getColumns().get(2).getGroupLevel());
                assertEquals(0, ws.getCells().getRows().get(0).getGroupLevel());
                assertEquals(0, ws.getCells().getColumns().get(0).getGroupLevel());
            }
        }
    }
}
