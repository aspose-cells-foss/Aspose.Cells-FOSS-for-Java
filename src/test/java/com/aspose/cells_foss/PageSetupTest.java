package com.aspose.cells_foss;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.util.Arrays;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 tests for PageSetup API — PS-* test cases.
 */
class PageSetupTest {

    private static final double MARGIN_DELTA = 1e-5;

    // =========================================================================
    // 3.1 Paper Size and Orientation
    // =========================================================================

    /**
     * Verifies that paper size roundtrips.
     */
    @Test
    void PS_01_paperSizeRoundtrips() {
        PaperSizeType[] sizes = {
            PaperSizeType.PAPER_A4, PaperSizeType.PAPER_A3,
            PaperSizeType.PAPER_LETTER, PaperSizeType.PAPER_LEGAL, PaperSizeType.PAPER_B5
        };
        // Walk the current collection so every entry is processed consistently.
        for (PaperSizeType size : sizes) {
            Workbook wb = new Workbook();
            PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
            ps.setPaperSize(size);
            assertEquals(size, ps.getPaperSize(), "Expected paper size " + size);
        }
    }

    /**
     * Verifies that orientation roundtrips.
     */
    @Test
    void PS_02_orientationRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (PageOrientationType ori : PageOrientationType.values()) {
            Workbook wb = new Workbook();
            PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
            ps.setOrientation(ori);
            assertEquals(ori, ps.getOrientation(), "Expected orientation " + ori);
        }
    }

    /**
     * Verifies that first page number roundtrips.
     */
    @Test
    void PS_03_firstPageNumberRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();

        ps.setFirstPageNumber(1);
        assertEquals(1, ps.getFirstPageNumber());

        ps.setFirstPageNumber(3);
        assertEquals(3, ps.getFirstPageNumber());

        ps.setFirstPageNumber(null);
        assertNull(ps.getFirstPageNumber());
    }

    /**
     * Verifies that scale roundtrips.
     */
    @Test
    void PS_04_scaleRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();

        ps.setScale(50);
        assertEquals(50, ps.getScale());

        ps.setScale(100);
        assertEquals(100, ps.getScale());

        ps.setScale(200);
        assertEquals(200, ps.getScale());

        ps.setScale(null);
        assertNull(ps.getScale());
    }

    /**
     * Verifies that fit to pages wide roundtrips.
     */
    @Test
    void PS_05_fitToPagesWideRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();

        ps.setFitToPagesWide(1);
        assertEquals(1, ps.getFitToPagesWide());

        ps.setFitToPagesWide(2);
        assertEquals(2, ps.getFitToPagesWide());

        ps.setFitToPagesWide(null);
        assertNull(ps.getFitToPagesWide());
    }

    /**
     * Verifies that fit to pages tall roundtrips.
     */
    @Test
    void PS_06_fitToPagesTallRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();

        ps.setFitToPagesTall(1);
        assertEquals(1, ps.getFitToPagesTall());

        ps.setFitToPagesTall(3);
        assertEquals(3, ps.getFitToPagesTall());

        ps.setFitToPagesTall(null);
        assertNull(ps.getFitToPagesTall());
    }

    // =========================================================================
    // 3.2 Margins (in centimeters)
    // =========================================================================

    /**
     * Verifies that left margin roundtrips.
     */
    @Test
    void PS_10_leftMarginRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (double v : new double[]{0.5, 0.7, 1.0}) {
            Workbook wb = new Workbook();
            PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
            ps.setLeftMargin(v);
            assertEquals(v, ps.getLeftMargin(), MARGIN_DELTA, "Left margin " + v);
        }
    }

    /**
     * Verifies that right margin roundtrips.
     */
    @Test
    void PS_11_rightMarginRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (double v : new double[]{0.5, 0.7, 1.0}) {
            Workbook wb = new Workbook();
            PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
            ps.setRightMargin(v);
            assertEquals(v, ps.getRightMargin(), MARGIN_DELTA, "Right margin " + v);
        }
    }

    /**
     * Verifies that top margin roundtrips.
     */
    @Test
    void PS_12_topMarginRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (double v : new double[]{0.75, 1.0}) {
            Workbook wb = new Workbook();
            PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
            ps.setTopMargin(v);
            assertEquals(v, ps.getTopMargin(), MARGIN_DELTA, "Top margin " + v);
        }
    }

    /**
     * Verifies that bottom margin roundtrips.
     */
    @Test
    void PS_13_bottomMarginRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (double v : new double[]{0.75, 1.0}) {
            Workbook wb = new Workbook();
            PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
            ps.setBottomMargin(v);
            assertEquals(v, ps.getBottomMargin(), MARGIN_DELTA, "Bottom margin " + v);
        }
    }

    /**
     * Verifies that header margin roundtrips.
     */
    @Test
    void PS_14_headerMarginRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (double v : new double[]{0.3, 0.5}) {
            Workbook wb = new Workbook();
            PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
            ps.setHeaderMargin(v);
            assertEquals(v, ps.getHeaderMargin(), MARGIN_DELTA, "Header margin " + v);
        }
    }

    /**
     * Verifies that footer margin roundtrips.
     */
    @Test
    void PS_15_footerMarginRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (double v : new double[]{0.3, 0.5}) {
            Workbook wb = new Workbook();
            PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
            ps.setFooterMargin(v);
            assertEquals(v, ps.getFooterMargin(), MARGIN_DELTA, "Footer margin " + v);
        }
    }

    /**
     * Verifies that all margins set together.
     */
    @Test
    void PS_16_allMarginsSetTogether() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setLeftMargin(0.5);
        ps.setRightMargin(0.6);
        ps.setTopMargin(0.7);
        ps.setBottomMargin(0.8);
        ps.setHeaderMargin(0.3);
        ps.setFooterMargin(0.4);

        assertEquals(0.5, ps.getLeftMargin(), MARGIN_DELTA);
        assertEquals(0.6, ps.getRightMargin(), MARGIN_DELTA);
        assertEquals(0.7, ps.getTopMargin(), MARGIN_DELTA);
        assertEquals(0.8, ps.getBottomMargin(), MARGIN_DELTA);
        assertEquals(0.3, ps.getHeaderMargin(), MARGIN_DELTA);
        assertEquals(0.4, ps.getFooterMargin(), MARGIN_DELTA);
    }

    // =========================================================================
    // 3.3 Header and Footer
    // =========================================================================

    /**
     * Verifies that left header roundtrips.
     */
    @Test
    void PS_20_leftHeaderRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setLeftHeader("Left Header");
        assertEquals("Left Header", ps.getLeftHeader());
        ps.setLeftHeader("");
        assertEquals("", ps.getLeftHeader());
    }

    /**
     * Verifies that center header roundtrips.
     */
    @Test
    void PS_21_centerHeaderRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setCenterHeader("Center Header");
        assertEquals("Center Header", ps.getCenterHeader());
    }

    /**
     * Verifies that right header roundtrips.
     */
    @Test
    void PS_22_rightHeaderRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setRightHeader("Right Header");
        assertEquals("Right Header", ps.getRightHeader());
    }

    /**
     * Verifies that left footer roundtrips.
     */
    @Test
    void PS_23_leftFooterRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setLeftFooter("Left Footer");
        assertEquals("Left Footer", ps.getLeftFooter());
    }

    /**
     * Verifies that center footer roundtrips.
     */
    @Test
    void PS_24_centerFooterRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setCenterFooter("Center Footer");
        assertEquals("Center Footer", ps.getCenterFooter());
    }

    /**
     * Verifies that right footer roundtrips.
     */
    @Test
    void PS_25_rightFooterRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setRightFooter("Right Footer");
        assertEquals("Right Footer", ps.getRightFooter());
    }

    /**
     * Verifies that header footer empty strings preserved.
     */
    @Test
    void PS_26_headerFooterEmptyStringsPreserved() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setLeftHeader("");
        ps.setCenterHeader("");
        ps.setRightHeader("");
        ps.setLeftFooter("");
        ps.setCenterFooter("");
        ps.setRightFooter("");

        assertEquals("", ps.getLeftHeader());
        assertEquals("", ps.getCenterHeader());
        assertEquals("", ps.getRightHeader());
        assertEquals("", ps.getLeftFooter());
        assertEquals("", ps.getCenterFooter());
        assertEquals("", ps.getRightFooter());
    }

    // =========================================================================
    // 3.4 Print Options
    // =========================================================================

    /**
     * Verifies that print gridlines roundtrips.
     */
    @Test
    void PS_30_printGridlinesRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setPrintGridlines(true);
        assertTrue(ps.getPrintGridlines());
        ps.setPrintGridlines(false);
        assertFalse(ps.getPrintGridlines());
    }

    /**
     * Verifies that print headings roundtrips.
     */
    @Test
    void PS_31_printHeadingsRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setPrintHeadings(true);
        assertTrue(ps.getPrintHeadings());
        ps.setPrintHeadings(false);
        assertFalse(ps.getPrintHeadings());
    }

    /**
     * Verifies that center horizontally roundtrips.
     */
    @Test
    void PS_32_centerHorizontallyRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setCenterHorizontally(true);
        assertTrue(ps.getCenterHorizontally());
        ps.setCenterHorizontally(false);
        assertFalse(ps.getCenterHorizontally());
    }

    /**
     * Verifies that center vertically roundtrips.
     */
    @Test
    void PS_33_centerVerticallyRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setCenterVertically(true);
        assertTrue(ps.getCenterVertically());
        ps.setCenterVertically(false);
        assertFalse(ps.getCenterVertically());
    }

    /**
     * Verifies that print area roundtrips.
     */
    @Test
    void PS_34_printAreaRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setPrintArea("A1:D10");
        assertEquals("A1:D10", ps.getPrintArea());
        ps.setPrintArea("");
        assertEquals("", ps.getPrintArea());
    }

    /**
     * Verifies that print title rows roundtrips.
     */
    @Test
    void PS_35_printTitleRowsRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setPrintTitleRows("$1:$2");
        assertEquals("$1:$2", ps.getPrintTitleRows());
    }

    /**
     * Verifies that print title columns roundtrips.
     */
    @Test
    void PS_36_printTitleColumnsRoundtrips() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.setPrintTitleColumns("$A:$B");
        assertEquals("$A:$B", ps.getPrintTitleColumns());
    }

    // =========================================================================
    // 3.5 Page Breaks
    // =========================================================================

    /**
     * Verifies that add horizontal page break.
     */
    @Test
    void PS_40_addHorizontalPageBreak() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addHorizontalPageBreak(4);
        assertTrue(ps.getHorizontalPageBreaks().contains(4));
    }

    /**
     * Verifies that add multiple horizontal breaks.
     */
    @Test
    void PS_41_addMultipleHorizontalBreaks() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addHorizontalPageBreak(4);
        ps.addHorizontalPageBreak(7);
        ps.addHorizontalPageBreak(12);
        assertEquals(3, ps.getHorizontalPageBreaks().size());
        assertTrue(ps.getHorizontalPageBreaks().contains(4));
        assertTrue(ps.getHorizontalPageBreaks().contains(7));
        assertTrue(ps.getHorizontalPageBreaks().contains(12));
    }

    /**
     * Verifies that clear horizontal breaks.
     */
    @Test
    void PS_42_clearHorizontalBreaks() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addHorizontalPageBreak(4);
        ps.addHorizontalPageBreak(7);
        ps.clearHorizontalPageBreaks();
        assertTrue(ps.getHorizontalPageBreaks().isEmpty());
    }

    /**
     * Verifies that add vertical page break.
     */
    @Test
    void PS_43_addVerticalPageBreak() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addVerticalPageBreak(2);
        assertTrue(ps.getVerticalPageBreaks().contains(2));
    }

    /**
     * Verifies that add multiple vertical breaks.
     */
    @Test
    void PS_44_addMultipleVerticalBreaks() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addVerticalPageBreak(2);
        ps.addVerticalPageBreak(5);
        assertEquals(2, ps.getVerticalPageBreaks().size());
        assertTrue(ps.getVerticalPageBreaks().contains(2));
        assertTrue(ps.getVerticalPageBreaks().contains(5));
    }

    /**
     * Verifies that clear vertical breaks.
     */
    @Test
    void PS_45_clearVerticalBreaks() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addVerticalPageBreak(2);
        ps.addVerticalPageBreak(5);
        ps.clearVerticalPageBreaks();
        assertTrue(ps.getVerticalPageBreaks().isEmpty());
    }

    /**
     * Verifies that duplicate horizontal break ignored.
     */
    @Test
    void PS_46_duplicateHorizontalBreakIgnored() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addHorizontalPageBreak(4);
        ps.addHorizontalPageBreak(4);
        assertEquals(1, ps.getHorizontalPageBreaks().size());
    }

    /**
     * Verifies that duplicate vertical break ignored.
     */
    @Test
    void PS_47_duplicateVerticalBreakIgnored() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addVerticalPageBreak(3);
        ps.addVerticalPageBreak(3);
        assertEquals(1, ps.getVerticalPageBreaks().size());
    }

    /**
     * Verifies that horizontal breaks returned in ascending order.
     */
    @Test
    void PS_48_horizontalBreaksReturnedInAscendingOrder() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addHorizontalPageBreak(12);
        ps.addHorizontalPageBreak(4);
        ps.addHorizontalPageBreak(7);
        java.util.List<Integer> breaks = ps.getHorizontalPageBreaks();
        assertEquals(java.util.Arrays.asList(4, 7, 12), breaks);
    }

    /**
     * Verifies that vertical breaks returned in ascending order.
     */
    @Test
    void PS_49_verticalBreaksReturnedInAscendingOrder() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addVerticalPageBreak(5);
        ps.addVerticalPageBreak(1);
        ps.addVerticalPageBreak(3);
        java.util.List<Integer> breaks = ps.getVerticalPageBreaks();
        assertEquals(java.util.Arrays.asList(1, 3, 5), breaks);
    }

    /**
     * Verifies that negative horizontal break throws.
     */
    @Test
    void PS_4A_negativeHorizontalBreakThrows() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        assertThrows(CellsException.class, () -> ps.addHorizontalPageBreak(-1));
    }

    /**
     * Verifies that negative vertical break throws.
     */
    @Test
    void PS_4B_negativeVerticalBreakThrows() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        assertThrows(CellsException.class, () -> ps.addVerticalPageBreak(-1));
    }

    /**
     * Verifies that zero row index accepted.
     */
    @Test
    void PS_4C_zeroRowIndexAccepted() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addHorizontalPageBreak(0);
        assertTrue(ps.getHorizontalPageBreaks().contains(0));
    }

    /**
     * Verifies that large row index accepted.
     */
    @Test
    void PS_4D_largeRowIndexAccepted() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
        ps.addHorizontalPageBreak(999999);
        assertTrue(ps.getHorizontalPageBreaks().contains(999999));
    }

    // =========================================================================
    // 3.6 Page Setup — XLSX Round-Trip
    // =========================================================================

    /**
     * Verifies that paper size orientation roundtrip.
     */
    @Test
    void PS_50_paperSizeOrientationRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("PageSetupTest")) {
            String path = tempDir.getPath("page-setup.xlsx");
            try (Workbook wb = new Workbook()) {
                PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
                ps.setPaperSize(PaperSizeType.PAPER_A4);
                ps.setOrientation(PageOrientationType.LANDSCAPE);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            PageSetup ps = loaded.getWorksheets().get(0).getPageSetup();
            assertEquals(PaperSizeType.PAPER_A4, ps.getPaperSize());
            assertEquals(PageOrientationType.LANDSCAPE, ps.getOrientation());
        }
    }

    /**
     * Verifies that all six margins roundtrip.
     */
    @Test
    void PS_51_allSixMarginsRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("PageSetupTest")) {
            String path = tempDir.getPath("page-setup-margins.xlsx");
            try (Workbook wb = new Workbook()) {
                PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
                ps.setLeftMargin(0.5);
                ps.setRightMargin(0.6);
                ps.setTopMargin(0.7);
                ps.setBottomMargin(0.8);
                ps.setHeaderMargin(0.3);
                ps.setFooterMargin(0.4);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            PageSetup ps = loaded.getWorksheets().get(0).getPageSetup();
            assertEquals(0.5, ps.getLeftMargin(), 1e-6);
            assertEquals(0.6, ps.getRightMargin(), 1e-6);
            assertEquals(0.7, ps.getTopMargin(), 1e-6);
            assertEquals(0.8, ps.getBottomMargin(), 1e-6);
            assertEquals(0.3, ps.getHeaderMargin(), 1e-6);
            assertEquals(0.4, ps.getFooterMargin(), 1e-6);
        }
    }

    /**
     * Verifies that header footer text roundtrip.
     */
    @Test
    void PS_52_headerFooterTextRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("PageSetupTest")) {
            String path = tempDir.getPath("page-setup-hf.xlsx");
            try (Workbook wb = new Workbook()) {
                PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
                ps.setLeftHeader("LH");
                ps.setCenterHeader("CH");
                ps.setRightHeader("RH");
                ps.setLeftFooter("LF");
                ps.setCenterFooter("CF");
                ps.setRightFooter("RF");
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            PageSetup ps = loaded.getWorksheets().get(0).getPageSetup();
            assertEquals("LH", ps.getLeftHeader());
            assertEquals("CH", ps.getCenterHeader());
            assertEquals("RH", ps.getRightHeader());
            assertEquals("LF", ps.getLeftFooter());
            assertEquals("CF", ps.getCenterFooter());
            assertEquals("RF", ps.getRightFooter());
        }
    }

    /**
     * Verifies that print options roundtrip.
     */
    @Test
    void PS_53_printOptionsRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("PageSetupTest")) {
            String path = tempDir.getPath("page-setup-print.xlsx");
            try (Workbook wb = new Workbook()) {
                PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
                ps.setPrintGridlines(true);
                ps.setPrintHeadings(true);
                ps.setCenterHorizontally(true);
                ps.setCenterVertically(true);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            PageSetup ps = loaded.getWorksheets().get(0).getPageSetup();
            assertTrue(ps.getPrintGridlines());
            assertTrue(ps.getPrintHeadings());
            assertTrue(ps.getCenterHorizontally());
            assertTrue(ps.getCenterVertically());
        }
    }

    /**
     * Verifies that horizontal page breaks roundtrip.
     */
    @Test
    void PS_54_horizontalPageBreaksRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("PageSetupTest")) {
            String path = tempDir.getPath("page-setup-hbreaks.xlsx");
            try (Workbook wb = new Workbook()) {
                PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
                ps.addHorizontalPageBreak(3);
                ps.addHorizontalPageBreak(7);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            PageSetup ps = loaded.getWorksheets().get(0).getPageSetup();
            assertTrue(ps.getHorizontalPageBreaks().contains(3));
            assertTrue(ps.getHorizontalPageBreaks().contains(7));
        }
    }

    /**
     * Verifies that vertical page breaks roundtrip.
     */
    @Test
    void PS_55_verticalPageBreaksRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("PageSetupTest")) {
            String path = tempDir.getPath("page-setup-vbreaks.xlsx");
            try (Workbook wb = new Workbook()) {
                PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
                ps.addVerticalPageBreak(2);
                ps.addVerticalPageBreak(5);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            PageSetup ps = loaded.getWorksheets().get(0).getPageSetup();
            assertTrue(ps.getVerticalPageBreaks().contains(2));
            assertTrue(ps.getVerticalPageBreaks().contains(5));
        }
    }

    /**
     * Verifies that scale roundtrip.
     */
    @Test
    void PS_56_scaleRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("PageSetupTest")) {
            String path = tempDir.getPath("page-setup-scale.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getPageSetup().setScale(75);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertEquals(75, loaded.getWorksheets().get(0).getPageSetup().getScale());
        }
    }

    /**
     * Verifies that fit to pages roundtrip.
     */
    @Test
    void PS_57_fitToPagesRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("PageSetupTest")) {
            String path = tempDir.getPath("fit-pages.xlsx");
            try (Workbook wb = new Workbook()) {
                PageSetup ps = wb.getWorksheets().get(0).getPageSetup();
                ps.setFitToPagesWide(2);
                ps.setFitToPagesTall(3);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            PageSetup ps = loaded.getWorksheets().get(0).getPageSetup();
            assertEquals(2, ps.getFitToPagesWide());
            assertEquals(3, ps.getFitToPagesTall());
        }
    }

    // =========================================================================
    // 3.7 Integration: Write Full Page Setup → Verify via API and POI
    // =========================================================================

    /**
     * Verifies that to ps 73 integration full page setup.
     */
    @Test
    void PS_60_to_PS_73_integrationFullPageSetup() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("PageSetupTest")) {
            String path = tempDir.getPath("page-setup-full.xlsx");

            try (Workbook wb = new Workbook()) {
                PageSetup ps = wb.getWorksheets().get(0).getPageSetup();

                // PS-60: Paper A4
                ps.setPaperSize(PaperSizeType.PAPER_A4);

                // PS-61: Landscape
                ps.setOrientation(PageOrientationType.LANDSCAPE);

                // PS-62: Scale 75%
                ps.setScale(75);

                // PS-63: All 6 margins (in centimeters)
                ps.setLeftMargin(0.5);
                ps.setRightMargin(0.6);
                ps.setTopMargin(0.7);
                ps.setBottomMargin(0.8);
                ps.setHeaderMargin(0.3);
                ps.setFooterMargin(0.4);

                // PS-64: Center header
                ps.setCenterHeader("Report \u0026P of \u0026N");

                // PS-65: Right footer
                ps.setRightFooter("\u0026D");

                // PS-66: Print gridlines
                ps.setPrintGridlines(true);

                // PS-67: Center horizontally
                ps.setCenterHorizontally(true);

                // PS-68: Center vertically
                ps.setCenterVertically(true);

                // PS-69: Horizontal page break at row 5
                ps.addHorizontalPageBreak(5);

                // PS-70: Vertical page break at col 3
                ps.addVerticalPageBreak(3);

                // PS-71: Fit to 1x2 pages
                ps.setFitToPagesWide(1);
                ps.setFitToPagesTall(2);

                // PS-72: First page number
                ps.setFirstPageNumber(3);

                // PS-73: Print headings
                ps.setPrintHeadings(true);

                wb.save(path);
            }

            // --- API verification ---
            Workbook loaded = new Workbook(path);
            PageSetup ps = loaded.getWorksheets().get(0).getPageSetup();

            // PS-60
            assertEquals(PaperSizeType.PAPER_A4, ps.getPaperSize());

            // PS-61
            assertEquals(PageOrientationType.LANDSCAPE, ps.getOrientation());

            // PS-62
            assertEquals(75, ps.getScale());

            // PS-63
            assertEquals(0.5, ps.getLeftMargin(), 1e-6);
            assertEquals(0.6, ps.getRightMargin(), 1e-6);
            assertEquals(0.7, ps.getTopMargin(), 1e-6);
            assertEquals(0.8, ps.getBottomMargin(), 1e-6);
            assertEquals(0.3, ps.getHeaderMargin(), 1e-6);
            assertEquals(0.4, ps.getFooterMargin(), 1e-6);

            // PS-64
            assertEquals("Report \u0026P of \u0026N", ps.getCenterHeader());

            // PS-65
            assertEquals("\u0026D", ps.getRightFooter());

            // PS-66
            assertTrue(ps.getPrintGridlines());

            // PS-67
            assertTrue(ps.getCenterHorizontally());

            // PS-68
            assertTrue(ps.getCenterVertically());

            // PS-69
            assertTrue(ps.getHorizontalPageBreaks().contains(5));

            // PS-70
            assertTrue(ps.getVerticalPageBreaks().contains(3));

            // PS-71
            assertEquals(1, ps.getFitToPagesWide());
            assertEquals(2, ps.getFitToPagesTall());

            // PS-72
            assertEquals(3, ps.getFirstPageNumber());

            // PS-73
            assertTrue(ps.getPrintHeadings());

            // --- POI verification ---
            try (org.apache.poi.ss.usermodel.Workbook poiWb = WorkbookFactory.create(new File(path))) {
                org.apache.poi.ss.usermodel.Sheet poiSheet = poiWb.getSheetAt(0);
                org.apache.poi.ss.usermodel.PrintSetup printSetup = poiSheet.getPrintSetup();

                // PS-60: A4 paper = code 9
                assertEquals(9, printSetup.getPaperSize());

                // PS-61: Landscape
                assertTrue(printSetup.getLandscape());

                // PS-62: Scale 75
                assertEquals(75, printSetup.getScale());

                // PS-63: Margins (POI uses inches; our API stores centimeters, converts to inches internally)
                assertEquals(0.5 / 2.54, poiSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.LeftMargin), 1e-6);
                assertEquals(0.6 / 2.54, poiSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.RightMargin), 1e-6);
                assertEquals(0.7 / 2.54, poiSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.TopMargin), 1e-6);
                assertEquals(0.8 / 2.54, poiSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.BottomMargin), 1e-6);
                assertEquals(0.3 / 2.54, poiSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.HeaderMargin), 1e-6);
                assertEquals(0.4 / 2.54, poiSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.FooterMargin), 1e-6);

                // PS-64: Center header
                assertEquals("Report \u0026P of \u0026N", poiSheet.getHeader().getCenter());

                // PS-65: Right footer
                assertEquals("&D", poiSheet.getFooter().getRight());

                // PS-66: Print gridlines
                assertTrue(poiSheet.isPrintGridlines());

                // PS-67: Horizontally center
                assertTrue(poiSheet.getHorizontallyCenter());

                // PS-68: Vertically center
                assertTrue(poiSheet.getVerticallyCenter());

                // PS-69: Horizontal page break at row 5
                assertTrue(Arrays.stream(poiSheet.getRowBreaks()).anyMatch(r -> r == 5),
                    "Should have row break at 5");

                // PS-70: Vertical page break at col 3
                assertTrue(Arrays.stream(poiSheet.getColumnBreaks()).anyMatch(c -> c == 3),
                    "Should have column break at 3");

                // PS-71: Fit width=1, fit height=2
                assertEquals(1, printSetup.getFitWidth());
                assertEquals(2, printSetup.getFitHeight());

                // PS-72: First page number
                assertEquals(3, printSetup.getPageStart());
                assertTrue(printSetup.getUsePage());

                // PS-73: Print headings — verify via worksheet XML
                String wsXml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
                assertTrue(wsXml.contains("headings=\"1\"") || wsXml.contains("<printOptions") && wsXml.contains("headings"),
                    "Worksheet XML should contain print headings option");
            }
        }
    }
}
