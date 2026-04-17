package com.aspose.cells_foss;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.common.usermodel.HyperlinkType;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 tests for the Hyperlink / HyperlinkCollection API — HL-* test cases.
 */
class HyperlinkTest {

    // =========================================================================
    // HL-01 to HL-16: HyperlinkCollection operations
    // =========================================================================

    /** HL-01: A fresh worksheet has no hyperlinks. */
    @Test
    void HL_01_newWorksheetHasNoHyperlinks() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertEquals(0, ws.getHyperlinks().getCount());
        }
    }

    /** HL-02: add(cellName, rows, cols, address) returns index 0 and count becomes 1. */
    @Test
    void HL_02_addByCellNameReturnsZeroIndex() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            int idx = ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            assertEquals(0, idx);
            assertEquals(1, ws.getHyperlinks().getCount());
        }
    }

    /** HL-03: Second add returns index 1. */
    @Test
    void HL_03_secondAddReturnsIndex1() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            int idx = ws.getHyperlinks().add("B2", 1, 1, "https://other.com");
            assertEquals(1, idx);
            assertEquals(2, ws.getHyperlinks().getCount());
        }
    }

    /** HL-04: add by row/col index returns correct index. */
    @Test
    void HL_04_addByRowColIndex() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            int idx = ws.getHyperlinks().add(0, 0, 1, 1, "https://example.com");
            assertEquals(0, idx);
            assertEquals(1, ws.getHyperlinks().getCount());
        }
    }

    /** HL-05: add(startCell, endCell, address, text, tip) adds a range hyperlink. */
    @Test
    void HL_05_addByStartEndCellName() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            int idx = ws.getHyperlinks().add("A1", "C3", "https://example.com", "Click me", "My tip");
            assertEquals(0, idx);
            assertEquals(1, ws.getHyperlinks().getCount());
        }
    }

    /** HL-06: get(0) returns the hyperlink added first. */
    @Test
    void HL_06_getByIndexReturnsCorrectHyperlink() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            Hyperlink hl = ws.getHyperlinks().get(0);
            assertNotNull(hl);
            assertEquals("https://example.com", hl.getAddress());
        }
    }

    /** HL-07: removeAt(0) removes the first hyperlink and count drops. */
    @Test
    void HL_07_removeAtDecrementsCount() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            ws.getHyperlinks().add("B2", 1, 1, "https://other.com");
            ws.getHyperlinks().removeAt(0);
            assertEquals(1, ws.getHyperlinks().getCount());
            assertEquals("https://other.com", ws.getHyperlinks().get(0).getAddress());
        }
    }

    /** HL-08: Hyperlink.delete() removes it from the collection. */
    @Test
    void HL_08_hyperlinkDeleteRemovesFromCollection() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            ws.getHyperlinks().get(0).delete();
            assertEquals(0, ws.getHyperlinks().getCount());
        }
    }

    /** HL-09: get() with out-of-range index throws CellsException. */
    @Test
    void HL_09_getOutOfRangeThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class, () -> ws.getHyperlinks().get(0));
        }
    }

    /** HL-10: get() with negative index throws CellsException. */
    @Test
    void HL_10_getNegativeIndexThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class, () -> ws.getHyperlinks().get(-1));
        }
    }

    /** HL-11: removeAt() with out-of-range index throws CellsException. */
    @Test
    void HL_11_removeAtOutOfRangeThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class, () -> ws.getHyperlinks().removeAt(0));
        }
    }

    /** HL-12: Adding overlapping hyperlinks throws CellsException. */
    @Test
    void HL_12_overlappingHyperlinksThrow() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://first.com");
            assertThrows(CellsException.class,
                    () -> ws.getHyperlinks().add("A1", 1, 1, "https://second.com"));
        }
    }

    /** HL-13: Zero or negative dimensions throw CellsException. */
    @Test
    void HL_13_zeroDimensionThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class,
                    () -> ws.getHyperlinks().add("A1", 0, 1, "https://example.com"));
            assertThrows(CellsException.class,
                    () -> ws.getHyperlinks().add("A1", 1, 0, "https://example.com"));
            assertThrows(CellsException.class,
                    () -> ws.getHyperlinks().add("A1", -1, 1, "https://example.com"));
        }
    }

    /** HL-14: Blank address throws CellsException on add. */
    @Test
    void HL_14_blankAddressOnAddThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class,
                    () -> ws.getHyperlinks().add("A1", 1, 1, ""));
            assertThrows(CellsException.class,
                    () -> ws.getHyperlinks().add("A1", 1, 1, "   "));
        }
    }

    /** HL-15: Negative row index throws CellsException on add by row/col. */
    @Test
    void HL_15_negativeRowIndexThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class,
                    () -> ws.getHyperlinks().add(-1, 0, 1, 1, "https://example.com"));
        }
    }

    /** HL-16: Negative column index throws CellsException on add by row/col. */
    @Test
    void HL_16_negativeColumnIndexThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class,
                    () -> ws.getHyperlinks().add(0, -1, 1, 1, "https://example.com"));
        }
    }

    // =========================================================================
    // HL-20 to HL-26: Address and link type detection
    // =========================================================================

    /** HL-20: An http:// address yields EXTERNAL link type. */
    @Test
    void HL_20_httpAddressIsExternalType() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            assertEquals(TargetModeType.EXTERNAL, ws.getHyperlinks().get(0).getLinkType());
        }
    }

    /** HL-21: A mailto: address yields EMAIL link type. */
    @Test
    void HL_21_mailtoAddressIsEmailType() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "mailto:user@example.com");
            assertEquals(TargetModeType.EMAIL, ws.getHyperlinks().get(0).getLinkType());
        }
    }

    /** HL-22: An address beginning with # is stored as CELL_REFERENCE and # is stripped. */
    @Test
    void HL_22_hashPrefixBecomCellReference() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "#Sheet1!A1");
            Hyperlink hl = ws.getHyperlinks().get(0);
            assertEquals(TargetModeType.CELL_REFERENCE, hl.getLinkType());
            assertEquals("Sheet1!A1", hl.getAddress());
        }
    }

    /** HL-23: An address containing ! (no # prefix) is stored as CELL_REFERENCE. */
    @Test
    void HL_23_bangNotationIsCellReference() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "Sheet1!B2");
            Hyperlink hl = ws.getHyperlinks().get(0);
            assertEquals(TargetModeType.CELL_REFERENCE, hl.getLinkType());
            assertEquals("Sheet1!B2", hl.getAddress());
        }
    }

    /** HL-24: setAddress() with a new http URL updates the address and keeps EXTERNAL type. */
    @Test
    void HL_24_setAddressUpdatesExternalUrl() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            Hyperlink hl = ws.getHyperlinks().get(0);
            hl.setAddress("https://updated.com");
            assertEquals("https://updated.com", hl.getAddress());
            assertEquals(TargetModeType.EXTERNAL, hl.getLinkType());
        }
    }

    /** HL-25: setAddress() with mailto: changes link type to EMAIL. */
    @Test
    void HL_25_setAddressChangesToEmail() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            Hyperlink hl = ws.getHyperlinks().get(0);
            hl.setAddress("mailto:test@example.com");
            assertEquals(TargetModeType.EMAIL, hl.getLinkType());
            assertEquals("mailto:test@example.com", hl.getAddress());
        }
    }

    /** HL-26: setAddress() with a # prefix changes type to CELL_REFERENCE and strips #. */
    @Test
    void HL_26_setAddressWithHashChangesToCellRef() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            Hyperlink hl = ws.getHyperlinks().get(0);
            hl.setAddress("#Sheet2!C5");
            assertEquals(TargetModeType.CELL_REFERENCE, hl.getLinkType());
            assertEquals("Sheet2!C5", hl.getAddress());
        }
    }

    // =========================================================================
    // HL-30 to HL-34: Display text and screen tip
    // =========================================================================

    /** HL-30: Default textToDisplay is an empty string. */
    @Test
    void HL_30_defaultTextToDisplayIsEmpty() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            assertEquals("", ws.getHyperlinks().get(0).getTextToDisplay());
        }
    }

    /** HL-31: Default screenTip is an empty string. */
    @Test
    void HL_31_defaultScreenTipIsEmpty() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            assertEquals("", ws.getHyperlinks().get(0).getScreenTip());
        }
    }

    /** HL-32: setTextToDisplay persists via getTextToDisplay. */
    @Test
    void HL_32_textToDisplayRoundtrips() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            Hyperlink hl = ws.getHyperlinks().get(0);
            hl.setTextToDisplay("Click here");
            assertEquals("Click here", hl.getTextToDisplay());
        }
    }

    /** HL-33: setScreenTip persists via getScreenTip. */
    @Test
    void HL_33_screenTipRoundtrips() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            Hyperlink hl = ws.getHyperlinks().get(0);
            hl.setScreenTip("Hover text");
            assertEquals("Hover text", hl.getScreenTip());
        }
    }

    /** HL-34: add(start, end, address, text, tip) captures textToDisplay and screenTip. */
    @Test
    void HL_34_addWithTextAndTipPreservesValues() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", "A1", "https://example.com", "Display text", "My tooltip");
            Hyperlink hl = ws.getHyperlinks().get(0);
            assertEquals("Display text", hl.getTextToDisplay());
            assertEquals("My tooltip", hl.getScreenTip());
        }
    }

    // =========================================================================
    // HL-40 to HL-45: getArea() for 1x1 and multi-cell; edge-case collection tests
    // =========================================================================

    /** HL-40: getArea() for a 1x1 hyperlink returns just the cell name. */
    @Test
    void HL_40_getAreaFor1x1ReturnsCellName() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("B3", 1, 1, "https://example.com");
            assertEquals("B3", ws.getHyperlinks().get(0).getArea());
        }
    }

    /** HL-41: getArea() for a multi-cell hyperlink returns a range like "A1:C3". */
    @Test
    void HL_41_getAreaForMultiCellReturnsRange() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 3, 3, "https://example.com");
            assertEquals("A1:C3", ws.getHyperlinks().get(0).getArea());
        }
    }

    /** HL-42: add(firstRow, firstColumn, totalRows, totalColumns, address) places hyperlink at correct area. */
    @Test
    void HL_42_addByRowColIndexSetsCorrectArea() throws Exception {
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            // Row 0, col 0 = A1; 2 rows, 2 cols => A1:B2
            ws.getHyperlinks().add(0, 0, 2, 2, "https://example.com");
            assertEquals("A1:B2", ws.getHyperlinks().get(0).getArea());
        }
    }

    /** HL-43: Three non-overlapping hyperlinks can coexist. */
    @Test
    void HL_43_threeNonOverlappingHyperlinks() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://one.com");
            ws.getHyperlinks().add("C1", 1, 1, "https://two.com");
            ws.getHyperlinks().add("E1", 1, 1, "https://three.com");
            assertEquals(3, ws.getHyperlinks().getCount());
        }
    }

    /** HL-44: Overlapping by range (not same cell) also throws. */
    @Test
    void HL_44_partialOverlapThrows() throws Exception {
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 2, 2, "https://first.com");
            // A1:B2 overlaps with B2:C3
            assertThrows(CellsException.class,
                    () -> ws.getHyperlinks().add("B2", 2, 2, "https://second.com"));
        }
    }

    /** HL-45: After delete, the same cell can be re-used. */
    @Test
    void HL_45_afterDeleteSameCellCanBeAdded() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
            ws.getHyperlinks().get(0).delete();
            int idx = ws.getHyperlinks().add("A1", 1, 1, "https://new.com");
            assertEquals(0, idx);
            assertEquals(1, ws.getHyperlinks().getCount());
        }
    }

    // =========================================================================
    // HL-50 to HL-55: XLSX round-trip (disabled — serializer not implemented)
    // =========================================================================

    /** HL-50: External URL survives save/reload round-trip. */
    @Test
    void HL_50_externalUrlRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl50.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().get("A1").putValue("Visit example.com");
                ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                Worksheet ws = loaded.getWorksheets().get(0);
                assertEquals(1, ws.getHyperlinks().getCount());
                assertEquals("https://example.com", ws.getHyperlinks().get(0).getAddress());
                assertEquals(TargetModeType.EXTERNAL, ws.getHyperlinks().get(0).getLinkType());
            }
        }
    }

    /** HL-51: mailto: address survives save/reload round-trip. */
    @Test
    void HL_51_mailtoAddressRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl51.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().get("B2").putValue("Email user@example.com");
                ws.getHyperlinks().add("B2", 1, 1, "mailto:user@example.com");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                Worksheet ws = loaded.getWorksheets().get(0);
                assertEquals(1, ws.getHyperlinks().getCount());
                assertEquals(TargetModeType.EMAIL, ws.getHyperlinks().get(0).getLinkType());
            }
        }
    }

    /** HL-52: Cell-reference hyperlink (# prefix) survives save/reload round-trip. */
    @Test
    void HL_52_cellRefHyperlinkRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl52.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().get("A1").putValue("Target cell");
                ws.getCells().get("C3").putValue("Go to A1");
                ws.getHyperlinks().add("C3", 1, 1, "#Sheet1!A1");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                Worksheet ws = loaded.getWorksheets().get(0);
                assertEquals(1, ws.getHyperlinks().getCount());
                assertEquals(TargetModeType.CELL_REFERENCE, ws.getHyperlinks().get(0).getLinkType());
                assertEquals("Sheet1!A1", ws.getHyperlinks().get(0).getAddress());
            }
        }
    }

    /** HL-53: screenTip text survives save/reload round-trip. */
    @Test
    void HL_53_screenTipRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl53.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().get("A1").putValue("Display");
                ws.getHyperlinks().add("A1", "A1", "https://example.com", "Display", "My Tip");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("My Tip", loaded.getWorksheets().get(0).getHyperlinks().get(0).getScreenTip());
            }
        }
    }

    /** HL-54: textToDisplay survives save/reload round-trip. */
    @Test
    void HL_54_textToDisplayRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl54.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().get("A1").putValue("Click here");
                ws.getHyperlinks().add("A1", "A1", "https://example.com", "Click here", "");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("Click here", loaded.getWorksheets().get(0).getHyperlinks().get(0).getTextToDisplay());
            }
        }
    }

    /** HL-55: Multi-cell range hyperlink survives save/reload round-trip. */
    @Test
    void HL_55_multiCellHyperlinkRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl55.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().get("A1").putValue("Multi-cell link (A1:C2)");
                ws.getHyperlinks().add("A1", 2, 3, "https://example.com");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals(1, loaded.getWorksheets().get(0).getHyperlinks().getCount());
                assertEquals("A1:C2", loaded.getWorksheets().get(0).getHyperlinks().get(0).getArea());
            }
        }
    }

    // =========================================================================
    // HL-60 to HL-65: Integration tests via Apache POI (disabled)
    // =========================================================================

    /** HL-60: POI reads back an external URL hyperlink written by Aspose. */
    @Test
    void HL_60_poiReadsExternalUrlFromAsposeXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl60.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().get("A1").putValue("Visit example.com");
                ws.getHyperlinks().add("A1", 1, 1, "https://example.com");
                wb.save(path);
            }
            try (XSSFWorkbook poiWb = (XSSFWorkbook) WorkbookFactory.create(new java.io.File(path))) {
                XSSFSheet sheet = poiWb.getSheetAt(0);
                var hyperlinks = sheet.getHyperlinkList();
                assertEquals(1, hyperlinks.size());
                assertEquals("https://example.com", hyperlinks.get(0).getAddress());
                assertEquals(HyperlinkType.URL, hyperlinks.get(0).getType());
            }
        }
    }

    /** HL-61: POI reads back a mailto hyperlink written by Aspose. */
    @Test
    void HL_61_poiReadsMailtoHyperlinkFromAsposeXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl61.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().get("B2").putValue("Contact user@example.com");
                ws.getHyperlinks().add("B2", 1, 1, "mailto:user@example.com");
                wb.save(path);
            }
            try (XSSFWorkbook poiWb = (XSSFWorkbook) WorkbookFactory.create(new java.io.File(path))) {
                XSSFSheet sheet = poiWb.getSheetAt(0);
                var hyperlinks = sheet.getHyperlinkList();
                assertEquals(1, hyperlinks.size());
                assertEquals(HyperlinkType.EMAIL, hyperlinks.get(0).getType());
            }
        }
    }

    /** HL-62: Aspose reads back an external URL hyperlink written by POI. */
    @Test
    void HL_62_asposeReadsExternalUrlFromPoiXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl62.xlsx");
            try (XSSFWorkbook poiWb = new XSSFWorkbook()) {
                XSSFSheet sheet = poiWb.createSheet("Sheet1");
                org.apache.poi.xssf.usermodel.XSSFHyperlink link =
                        poiWb.getCreationHelper().createHyperlink(HyperlinkType.URL);
                link.setAddress("https://example.com");
                var cell = sheet.createRow(0).createCell(0);
                cell.setCellValue("Visit example.com");
                cell.setHyperlink(link);
                try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path)) {
                    poiWb.write(fos);
                }
            }
            try (Workbook wb = new Workbook(path)) {
                Worksheet ws = wb.getWorksheets().get(0);
                assertEquals(1, ws.getHyperlinks().getCount());
                assertEquals("https://example.com", ws.getHyperlinks().get(0).getAddress());
            }
        }
    }

    /** HL-63: Aspose reads back a mailto hyperlink written by POI. */
    @Test
    void HL_63_asposeReadsMailtoHyperlinkFromPoiXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl63.xlsx");
            try (XSSFWorkbook poiWb = new XSSFWorkbook()) {
                XSSFSheet sheet = poiWb.createSheet("Sheet1");
                org.apache.poi.xssf.usermodel.XSSFHyperlink link =
                        poiWb.getCreationHelper().createHyperlink(HyperlinkType.EMAIL);
                link.setAddress("mailto:user@example.com");
                var cell = sheet.createRow(0).createCell(0);
                cell.setCellValue("Email user@example.com");
                cell.setHyperlink(link);
                try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path)) {
                    poiWb.write(fos);
                }
            }
            try (Workbook wb = new Workbook(path)) {
                Worksheet ws = wb.getWorksheets().get(0);
                assertEquals(1, ws.getHyperlinks().getCount());
                assertEquals(TargetModeType.EMAIL, ws.getHyperlinks().get(0).getLinkType());
            }
        }
    }

    /** HL-64: Two non-overlapping hyperlinks survive round-trip through XLSX, preserving order. */
    @Test
    void HL_64_twoHyperlinksRoundtripViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl64.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getCells().get("A1").putValue("First link");
                ws.getCells().get("C1").putValue("Second link");
                ws.getHyperlinks().add("A1", 1, 1, "https://first.com");
                ws.getHyperlinks().add("C1", 1, 1, "https://second.com");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals(2, loaded.getWorksheets().get(0).getHyperlinks().getCount());
            }
        }
    }

    /** HL-65: POI-created cell-reference hyperlink is readable by Aspose. */
    @Test
    void HL_65_asposeReadsCellRefHyperlinkFromPoiXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("HyperlinkTest")) {
            String path = tempDir.getPath("hl65.xlsx");
            try (XSSFWorkbook poiWb = new XSSFWorkbook()) {
                XSSFSheet sheet = poiWb.createSheet("Sheet1");
                org.apache.poi.xssf.usermodel.XSSFHyperlink link =
                        poiWb.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
                link.setAddress("Sheet1!A1");
                var cell = sheet.createRow(0).createCell(0);
                cell.setCellValue("Go to Sheet1!A1");
                cell.setHyperlink(link);
                try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path)) {
                    poiWb.write(fos);
                }
            }
            try (Workbook wb = new Workbook(path)) {
                Worksheet ws = wb.getWorksheets().get(0);
                assertEquals(1, ws.getHyperlinks().getCount());
                assertEquals(TargetModeType.CELL_REFERENCE, ws.getHyperlinks().get(0).getLinkType());
            }
        }
    }
}
