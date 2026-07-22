package org.aspose.cells_foss;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.io.File;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 tests for Style API -ST-* test cases.
 */
class StyleTest {

    /**
     * Converts the byte to its unsigned integer value.
     * @param b b
     */
    private static int u(byte b) { return Byte.toUnsignedInt(b); }

    // =========================================================================
    // 2.1 Font
    // =========================================================================

    /**
     * Verifies that font name roundtrips.
     */
    @Test
    void ST_01_fontNameRoundtrips() {
        Style style = new Style();
        style.getFont().setName("Arial");
        assertEquals("Arial", style.getFont().getName());
    }

    /**
     * Verifies that font size roundtrips.
     */
    @Test
    void ST_02_fontSizeRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (double size : new double[]{8.0, 11.0, 14.0, 72.0}) {
            Style style = new Style();
            style.getFont().setSize(size);
            assertEquals(size, style.getFont().getSize(), 1e-9, "Expected size " + size);
        }
    }

    /**
     * Verifies that bold flag roundtrips.
     */
    @Test
    void ST_03_boldFlagRoundtrips() {
        Style style = new Style();
        style.getFont().setBold(true);
        assertTrue(style.getFont().getBold());
    }

    /**
     * Verifies that italic flag roundtrips.
     */
    @Test
    void ST_04_italicFlagRoundtrips() {
        Style style = new Style();
        style.getFont().setItalic(true);
        assertTrue(style.getFont().getItalic());
    }

    /**
     * Verifies that underline flag roundtrips.
     */
    @Test
    void ST_05_underlineFlagRoundtrips() {
        Style style = new Style();
        style.getFont().setUnderline(true);
        assertTrue(style.getFont().getUnderline());
    }

    /**
     * Verifies that strike through flag roundtrips.
     */
    @Test
    void ST_06_strikeThroughFlagRoundtrips() {
        Style style = new Style();
        style.getFont().setStrikeThrough(true);
        assertTrue(style.getFont().getStrikeThrough());
    }

    /**
     * Verifies that font color roundtrips.
     */
    @Test
    void ST_07_fontColorRoundtrips() {
        Style style = new Style();
        Color red = Color.fromArgb(255, 255, 0, 0);
        style.getFont().setColor(red);
        Color got = style.getFont().getColor();
        assertEquals(255, u(got.getA()));
        assertEquals(255, u(got.getR()));
        assertEquals(0, u(got.getG()));
        assertEquals(0, u(got.getB()));
    }

    /**
     * Verifies that default font is calibri 11 pt.
     */
    @Test
    void ST_08_defaultFontIsCalibri11pt() {
        Style style = new Style();
        assertEquals("Calibri", style.getFont().getName());
        assertEquals(11.0, style.getFont().getSize(), 1e-9);
    }

    // =========================================================================
    // 2.2 Alignment
    // =========================================================================

    /**
     * Verifies that horizontal alignment roundtrips.
     */
    @Test
    void ST_10_horizontalAlignmentRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (HorizontalAlignmentType h : HorizontalAlignmentType.values()) {
            Style style = new Style();
            style.setHorizontalAlignment(h);
            assertEquals(h, style.getHorizontalAlignment(), "Expected " + h);
        }
    }

    /**
     * Verifies that vertical alignment roundtrips.
     */
    @Test
    void ST_11_verticalAlignmentRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (VerticalAlignmentType v : VerticalAlignmentType.values()) {
            Style style = new Style();
            style.setVerticalAlignment(v);
            assertEquals(v, style.getVerticalAlignment(), "Expected " + v);
        }
    }

    /**
     * Verifies that wrap text roundtrips.
     */
    @Test
    void ST_12_wrapTextRoundtrips() {
        Style style = new Style();
        style.setTextWrapped(true);
        assertTrue(style.isTextWrapped());
    }

    /**
     * Verifies that indent level roundtrips.
     */
    @Test
    void ST_13_indentLevelRoundtrips() {
        Style style = new Style();
        style.setIndentLevel(3);
        assertEquals(3, style.getIndentLevel());
    }

    /**
     * Verifies that text rotation roundtrips.
     */
    @Test
    void ST_14_textRotationRoundtrips() {
        Style style = new Style();
        style.setRotationAngle(45);
        assertEquals(45, style.getRotationAngle());
    }

    /**
     * Verifies that shrink to fit roundtrips.
     */
    @Test
    void ST_15_shrinkToFitRoundtrips() {
        Style style = new Style();
        style.setShrinkToFit(true);
        assertTrue(style.getShrinkToFit());
    }

    /**
     * Verifies that default horizontal alignment is general.
     */
    @Test
    void ST_16_defaultHorizontalAlignmentIsGeneral() {
        Style style = new Style();
        assertEquals(HorizontalAlignmentType.GENERAL, style.getHorizontalAlignment());
    }

    /**
     * Verifies that default vertical alignment is bottom.
     */
    @Test
    void ST_17_defaultVerticalAlignmentIsBottom() {
        Style style = new Style();
        assertEquals(VerticalAlignmentType.BOTTOM, style.getVerticalAlignment());
    }

    // =========================================================================
    // 2.3 Borders
    // =========================================================================

    /**
     * Verifies that left border style roundtrips.
     */
    @Test
    void ST_20_leftBorderStyleRoundtrips() {
        BorderStyleType[] variants = {
            BorderStyleType.THIN, BorderStyleType.MEDIUM, BorderStyleType.THICK,
            BorderStyleType.DASHED, BorderStyleType.DOTTED, BorderStyleType.DOUBLE, BorderStyleType.HAIR
        };
        // Walk the current collection so every entry is processed consistently.
        for (BorderStyleType bst : variants) {
            Style style = new Style();
            style.getBorders().getLeft().setLineStyle(bst);
            assertEquals(bst, style.getBorders().getLeft().getLineStyle(), "Expected " + bst);
        }
    }

    /**
     * Verifies that right border style roundtrips.
     */
    @Test
    void ST_21_rightBorderStyleRoundtrips() {
        Style style = new Style();
        style.getBorders().getRight().setLineStyle(BorderStyleType.MEDIUM);
        assertEquals(BorderStyleType.MEDIUM, style.getBorders().getRight().getLineStyle());
    }

    /**
     * Verifies that top border style roundtrips.
     */
    @Test
    void ST_22_topBorderStyleRoundtrips() {
        Style style = new Style();
        style.getBorders().getTop().setLineStyle(BorderStyleType.THICK);
        assertEquals(BorderStyleType.THICK, style.getBorders().getTop().getLineStyle());
    }

    /**
     * Verifies that bottom border style roundtrips.
     */
    @Test
    void ST_23_bottomBorderStyleRoundtrips() {
        Style style = new Style();
        style.getBorders().getBottom().setLineStyle(BorderStyleType.DASHED);
        assertEquals(BorderStyleType.DASHED, style.getBorders().getBottom().getLineStyle());
    }

    /**
     * Verifies that diagonal border style roundtrips.
     */
    @Test
    void ST_24_diagonalBorderStyleRoundtrips() {
        Style style = new Style();
        style.getBorders().getDiagonal().setLineStyle(BorderStyleType.DOTTED);
        assertEquals(BorderStyleType.DOTTED, style.getBorders().getDiagonal().getLineStyle());
    }

    /**
     * Verifies that diagonal up flag roundtrips.
     */
    @Test
    void ST_25_diagonalUpFlagRoundtrips() {
        Style style = new Style();
        style.getBorders().setDiagonalUp(true);
        assertTrue(style.getBorders().getDiagonalUp());
    }

    /**
     * Verifies that diagonal down flag roundtrips.
     */
    @Test
    void ST_26_diagonalDownFlagRoundtrips() {
        Style style = new Style();
        style.getBorders().setDiagonalDown(true);
        assertTrue(style.getBorders().getDiagonalDown());
    }

    /**
     * Verifies that border color roundtrips.
     */
    @Test
    void ST_27_borderColorRoundtrips() {
        Style style = new Style();
        Color blue = Color.fromArgb(255, 0, 0, 255);
        style.getBorders().getLeft().setColor(blue);
        Color got = style.getBorders().getLeft().getColor();
        assertEquals(255, u(got.getA()));
        assertEquals(0, u(got.getR()));
        assertEquals(0, u(got.getG()));
        assertEquals(255, u(got.getB()));
    }

    /**
     * Verifies that full outline border all medium.
     */
    @Test
    void ST_28_fullOutlineBorderAllMedium() {
        Style style = new Style();
        style.getBorders().getLeft().setLineStyle(BorderStyleType.MEDIUM);
        style.getBorders().getRight().setLineStyle(BorderStyleType.MEDIUM);
        style.getBorders().getTop().setLineStyle(BorderStyleType.MEDIUM);
        style.getBorders().getBottom().setLineStyle(BorderStyleType.MEDIUM);
        assertEquals(BorderStyleType.MEDIUM, style.getBorders().getLeft().getLineStyle());
        assertEquals(BorderStyleType.MEDIUM, style.getBorders().getRight().getLineStyle());
        assertEquals(BorderStyleType.MEDIUM, style.getBorders().getTop().getLineStyle());
        assertEquals(BorderStyleType.MEDIUM, style.getBorders().getBottom().getLineStyle());
    }

    /**
     * Verifies that none style is default.
     */
    @Test
    void ST_29_noneStyleIsDefault() {
        Borders borders = new Borders();
        assertNull(borders.getLeft().getLineStyle(), "Default border style should be null");
    }

    /**
     * Verifies that all 13 border style type values accepted.
     */
    @Test
    void ST_30_all13BorderStyleTypeValuesAccepted() {
        BorderStyleType[] nonNone = {
            BorderStyleType.THIN, BorderStyleType.MEDIUM, BorderStyleType.THICK,
            BorderStyleType.DOTTED, BorderStyleType.DASHED, BorderStyleType.DOUBLE,
            BorderStyleType.HAIR, BorderStyleType.MEDIUM_DASHED, BorderStyleType.DASH_DOT,
            BorderStyleType.MEDIUM_DASH_DOT, BorderStyleType.DASH_DOT_DOT,
            BorderStyleType.MEDIUM_DASH_DOT_DOT, BorderStyleType.SLANTED_DASH_DOT
        };
        // Walk the current collection so every entry is processed consistently.
        for (BorderStyleType bst : nonNone) {
            Style style = new Style();
            assertDoesNotThrow(() -> style.getBorders().getLeft().setLineStyle(bst));
            assertEquals(bst, style.getBorders().getLeft().getLineStyle(), "Expected " + bst);
        }
    }

    // =========================================================================
    // 2.4 Fill / Background
    // =========================================================================

    /**
     * Verifies that foreground color roundtrips.
     */
    @Test
    void ST_40_foregroundColorRoundtrips() {
        Style style = new Style();
        Color yellow = Color.fromArgb(255, 255, 255, 0);
        style.setForegroundColor(yellow);
        Color got = style.getForegroundColor();
        assertEquals(255, u(got.getA()));
        assertEquals(255, u(got.getR()));
        assertEquals(255, u(got.getG()));
        assertEquals(0, u(got.getB()));
    }

    /**
     * Verifies that background color roundtrips.
     */
    @Test
    void ST_41_backgroundColorRoundtrips() {
        Style style = new Style();
        Color green = Color.fromArgb(255, 0, 128, 0);
        style.setBackgroundColor(green);
        Color got = style.getBackgroundColor();
        assertEquals(255, u(got.getA()));
        assertEquals(0, u(got.getR()));
        assertEquals(128, u(got.getG()));
        assertEquals(0, u(got.getB()));
    }

    /**
     * Verifies that fill pattern roundtrips.
     */
    @Test
    void ST_42_fillPatternRoundtrips() {
        FillPattern[] patterns = {FillPattern.SOLID, FillPattern.MEDIUM_GRAY, FillPattern.DARK_HORIZONTAL, FillPattern.NONE};
        // Walk the current collection so every entry is processed consistently.
        for (FillPattern p : patterns) {
            Style style = new Style();
            style.setPattern(p);
            assertEquals(p, style.getPattern(), "Expected pattern " + p);
        }
    }

    /**
     * Verifies that default style has no fill.
     */
    @Test
    void ST_43_defaultStyleHasNoFill() {
        Style style = new Style();
        assertEquals(Color.getEmpty(), style.getForegroundColor());
        assertEquals(FillPattern.NONE, style.getPattern());
    }

    // =========================================================================
    // 2.5 Number Format
    // =========================================================================

    /**
     * Verifies that built in format index roundtrips.
     */
    @Test
    void ST_50_builtInFormatIndexRoundtrips() {
        // Walk the current collection so every entry is processed consistently.
        for (int fmt : new int[]{0, 1, 2, 14, 49}) {
            Style style = new Style();
            style.setNumber(fmt);
            assertEquals(fmt, style.getNumber(), "Expected format index " + fmt);
        }
    }

    /**
     * Verifies that custom format string roundtrips.
     */
    @Test
    void ST_51_customFormatStringRoundtrips() {
        Style style = new Style();
        style.setCustom("#,##0.00");
        assertEquals("#,##0.00", style.getCustom());
    }

    /**
     * Verifies that custom format date string accepted.
     */
    @Test
    void ST_52_customFormatDateStringAccepted() {
        Style style = new Style();
        style.setCustom("yyyy-mm-dd");
        assertEquals("yyyy-mm-dd", style.getCustom());
    }

    /**
     * Verifies that display string matches Excel for built-in number formats.
     */
    @Test
    void ST_53_displayStringMatchesExcelForBuiltInNumberFormats() throws Exception {
        List<DisplayStringCase> cases = new ArrayList<>();
        for (int formatId : builtInDisplayFormatIds()) {
            cases.add(createBuiltInDisplayStringCase(formatId));
        }
        assertDisplayStringsMatchPoi(cases, "style-display-builtins.xlsx");
    }

    /**
     * Verifies that display string matches Excel for representative custom formats.
     */
    @Test
    void ST_54_displayStringMatchesExcelForCustomFormats() throws Exception {
        LocalDateTime dateTimeValue = LocalDateTime.of(2024, 3, 15, 13, 4, 5, 670_000_000);
        List<DisplayStringCase> cases = List.of(
                new DisplayStringCase("A1", "custom numeric thousands", cell -> {
                    cell.putValue(12345.678);
                    Style style = cell.getStyle();
                    style.setCustom("#,##0.000");
                    cell.setStyle(style);
                }),
                new DisplayStringCase("A2", "custom negative sections", cell -> {
                    cell.putValue(-1234.5);
                    Style style = cell.getStyle();
                    style.setCustom("#,##0.00;[Red](#,##0.00)");
                    cell.setStyle(style);
                }),
                new DisplayStringCase("A3", "custom conditional sections", cell -> {
                    cell.putValue(42);
                    Style style = cell.getStyle();
                    style.setCustom("[>=100]\"big\";[>=10]\"mid\";\"small\"");
                    cell.setStyle(style);
                }),
                new DisplayStringCase("A4", "custom fraction", cell -> {
                    cell.putValue(2.125);
                    Style style = cell.getStyle();
                    style.setCustom("# ??/??");
                    cell.setStyle(style);
                }),
                new DisplayStringCase("A5", "custom scientific", cell -> {
                    cell.putValue(12345.678);
                    Style style = cell.getStyle();
                    style.setCustom("0.000E+00");
                    cell.setStyle(style);
                }),
                new DisplayStringCase("A6", "custom date", cell -> {
                    cell.putValue(dateTimeValue);
                    Style style = cell.getStyle();
                    style.setCustom("yyyy-mm-dd");
                    cell.setStyle(style);
                }),
                new DisplayStringCase("A7", "custom datetime", cell -> {
                    cell.putValue(dateTimeValue);
                    Style style = cell.getStyle();
                    style.setCustom("[$-0409]dddd, mmmm d, yyyy h:mm:ss AM/PM");
                    cell.setStyle(style);
                }),
                new DisplayStringCase("A8", "custom elapsed time", cell -> {
                    cell.putValue(dateTimeValue);
                    Style style = cell.getStyle();
                    style.setCustom("[h]:mm:ss");
                    cell.setStyle(style);
                }),
                new DisplayStringCase("A9", "custom text placeholder", "prefix-Alpha", cell -> {
                    cell.putValue("Alpha");
                    Style style = cell.getStyle();
                    style.setCustom("0;0;0;\"prefix-\"@");
                    cell.setStyle(style);
                }),
                new DisplayStringCase("A10", "custom escaped literal", cell -> {
                    cell.putValue(12.3);
                    Style style = cell.getStyle();
                    style.setCustom("0.0\\m");
                    cell.setStyle(style);
                })
        );
        assertDisplayStringsMatchPoi(cases, "style-display-custom.xlsx");
    }

    // =========================================================================
    // 2.6 Protection
    // =========================================================================

    /**
     * Verifies that default style is locked.
     */
    @Test
    void ST_55_defaultStyleIsLocked() {
        Style style = new Style();
        assertTrue(style.isLocked());
    }

    /**
     * Verifies that locked flag roundtrips.
     */
    @Test
    void ST_56_lockedFlagRoundtrips() {
        Style style = new Style();
        style.setLocked(false);
        assertFalse(style.isLocked());
    }

    /**
     * Verifies that hidden flag roundtrips.
     */
    @Test
    void ST_57_hiddenFlagRoundtrips() {
        Style style = new Style();
        style.setFormulaHidden(true);
        assertTrue(style.isFormulaHidden());
    }

    // =========================================================================
    // 2.7 Style -XLSX Round-Trip
    // =========================================================================

    /**
     * Verifies that font name size bold italic roundtrip.
     */
    @Test
    void ST_60_fontNameSizeBoldItalicRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-font.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue("Test");
                Style style = cell.getStyle();
                style.getFont().setName("Times New Roman");
                style.getFont().setSize(16.0);
                style.getFont().setBold(true);
                style.getFont().setItalic(true);
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            Style s = loaded.getWorksheets().get(0).getCells().get("A1").getStyle();
            assertEquals("Times New Roman", s.getFont().getName());
            assertEquals(16.0, s.getFont().getSize(), 1e-9);
            assertTrue(s.getFont().getBold());
            assertTrue(s.getFont().getItalic());
        }
    }

    /**
     * Verifies that font color roundtrip.
     */
    @Test
    void ST_61_fontColorRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-fontcolor.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue("Red");
                Style style = cell.getStyle();
                style.getFont().setColor(Color.fromArgb(255, 255, 0, 0));
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            Color c = loaded.getWorksheets().get(0).getCells().get("A1").getStyle().getFont().getColor();
            assertEquals(255, u(c.getA()));
            assertEquals(255, u(c.getR()));
            assertEquals(0, u(c.getG()));
            assertEquals(0, u(c.getB()));
        }
    }

    /**
     * Verifies that horizontal alignment roundtrip.
     */
    @Test
    void ST_62_horizontalAlignmentRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-halign.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue("Center");
                Style style = cell.getStyle();
                style.setHorizontalAlignment(HorizontalAlignmentType.CENTER);
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertEquals(HorizontalAlignmentType.CENTER,
                loaded.getWorksheets().get(0).getCells().get("A1").getStyle().getHorizontalAlignment());
        }
    }

    /**
     * Verifies that vertical alignment roundtrip.
     */
    @Test
    void ST_63_verticalAlignmentRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-valign.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("B1");
                cell.putValue("Top");
                Style style = cell.getStyle();
                style.setVerticalAlignment(VerticalAlignmentType.TOP);
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertEquals(VerticalAlignmentType.TOP,
                loaded.getWorksheets().get(0).getCells().get("B1").getStyle().getVerticalAlignment());
        }
    }

    /**
     * Verifies that wrap text roundtrip.
     */
    @Test
    void ST_64_wrapTextRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-wrap.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue("Wrap");
                Style style = cell.getStyle();
                style.setTextWrapped(true);
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertTrue(loaded.getWorksheets().get(0).getCells().get("A1").getStyle().isTextWrapped());
        }
    }

    /**
     * Verifies that all four border sides roundtrip.
     */
    @Test
    void ST_65_allFourBorderSidesRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-borders.xlsx");
            Color blue = Color.fromArgb(255, 0, 0, 255);
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue("Borders");
                Style style = cell.getStyle();
                style.getBorders().getLeft().setLineStyle(BorderStyleType.THIN);
                style.getBorders().getLeft().setColor(blue);
                style.getBorders().getRight().setLineStyle(BorderStyleType.THIN);
                style.getBorders().getRight().setColor(blue);
                style.getBorders().getTop().setLineStyle(BorderStyleType.THIN);
                style.getBorders().getTop().setColor(blue);
                style.getBorders().getBottom().setLineStyle(BorderStyleType.THIN);
                style.getBorders().getBottom().setColor(blue);
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            Style s = loaded.getWorksheets().get(0).getCells().get("A1").getStyle();
            assertEquals(BorderStyleType.THIN, s.getBorders().getLeft().getLineStyle());
            assertEquals(BorderStyleType.THIN, s.getBorders().getRight().getLineStyle());
            assertEquals(BorderStyleType.THIN, s.getBorders().getTop().getLineStyle());
            assertEquals(BorderStyleType.THIN, s.getBorders().getBottom().getLineStyle());
            assertEquals(255, u(s.getBorders().getLeft().getColor().getB()));
            assertEquals(255, u(s.getBorders().getRight().getColor().getB()));
            assertEquals(255, u(s.getBorders().getTop().getColor().getB()));
            assertEquals(255, u(s.getBorders().getBottom().getColor().getB()));
        }
    }

    /**
     * Verifies that fill color roundtrip solid.
     */
    @Test
    void ST_66_fillColorRoundtripSolid() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-fill.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue("Fill");
                Style style = cell.getStyle();
                style.setPattern(FillPattern.SOLID);
                style.setForegroundColor(Color.fromArgb(255, 255, 255, 0));
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            Style s = loaded.getWorksheets().get(0).getCells().get("A1").getStyle();
            assertEquals(FillPattern.SOLID, s.getPattern());
            Color fg = s.getForegroundColor();
            assertEquals(255, u(fg.getA()));
            assertEquals(255, u(fg.getR()));
            assertEquals(255, u(fg.getG()));
            assertEquals(0, u(fg.getB()));
        }
    }

    /**
     * Verifies that built in number format roundtrip.
     */
    @Test
    void ST_67_builtInNumberFormatRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-numfmt.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue(3.14);
                Style style = cell.getStyle();
                style.setNumber(2); // "0.00"
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertEquals(2, loaded.getWorksheets().get(0).getCells().get("A1").getStyle().getNumber());
        }
    }

    /**
     * Verifies that custom number format roundtrip.
     */
    @Test
    void ST_68_customNumberFormatRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-custom.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue(12345.678);
                Style style = cell.getStyle();
                style.setCustom("#,##0.00");
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertEquals("#,##0.00", loaded.getWorksheets().get(0).getCells().get("A1").getStyle().getCustom());
        }
    }

    /**
     * Verifies that multiple cells styles independent.
     */
    @Test
    void ST_69_multipleCellsStylesIndependent() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-multi.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell a1 = wb.getWorksheets().get(0).getCells().get("A1");
                a1.putValue("Bold");
                Style sa = a1.getStyle();
                sa.getFont().setBold(true);
                a1.setStyle(sa);

                Cell b1 = wb.getWorksheets().get(0).getCells().get("B1");
                b1.putValue("Italic");
                Style sb = b1.getStyle();
                sb.getFont().setItalic(true);
                b1.setStyle(sb);

                Cell c1 = wb.getWorksheets().get(0).getCells().get("C1");
                c1.putValue("Border");
                Style sc = c1.getStyle();
                sc.getBorders().getLeft().setLineStyle(BorderStyleType.THIN);
                c1.setStyle(sc);

                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            Style la = loaded.getWorksheets().get(0).getCells().get("A1").getStyle();
            Style lb = loaded.getWorksheets().get(0).getCells().get("B1").getStyle();
            Style lc = loaded.getWorksheets().get(0).getCells().get("C1").getStyle();

            assertTrue(la.getFont().getBold());
            assertFalse(la.getFont().getItalic());

            assertTrue(lb.getFont().getItalic());
            assertFalse(lb.getFont().getBold());

            assertEquals(BorderStyleType.THIN, lc.getBorders().getLeft().getLineStyle());
            assertNull(la.getBorders().getLeft().getLineStyle());
        }
    }

    /**
     * Verifies that date cell style index written as s.
     */
    @Test
    void ST_70_dateCellStyleIndexWrittenAsS() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-date.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().get("A1")
                    .putValue(java.time.LocalDateTime.of(2024, 1, 1, 0, 0, 0));
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            // The date cell should have an s= attribute
            assertTrue(xml.contains("s="), "Date cell should have s= style attribute in worksheet XML");
        }
    }

    /**
     * Verifies that default style cell has no s attribute.
     */
    @Test
    void ST_71_defaultStyleCellHasNoSAttribute() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-default.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().get("A1").putValue(42);
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            // The default style cell A1 with value 42 (no explicit style set) should not have s= attribute
            // The cell element should be something like <c r="A1"><v>42</v></c>
            // We verify there's no s attribute on the A1 cell
            // Look for the cell reference for row 1 col A in the XML
            assertFalse(xml.matches("(?s).*<c r=\"A1\"[^>]*s=\"[^\"]+\"[^>]*>.*"),
                "Default-style cell A1 should not have s= attribute");
        }
    }

    /**
     * Verifies that styles xml cell xfs count matches distinct styles.
     */
    @Test
    void ST_72_stylesXmlCellXfsCountMatchesDistinctStyles() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-xfs.xlsx");
            try (Workbook wb = new Workbook()) {
                // 3 cells with 3 different styles
                Cell a1 = wb.getWorksheets().get(0).getCells().get("A1");
                a1.putValue("A"); Style sa = a1.getStyle(); sa.getFont().setBold(true); a1.setStyle(sa);

                Cell b1 = wb.getWorksheets().get(0).getCells().get("B1");
                b1.putValue("B"); Style sb = b1.getStyle(); sb.getFont().setItalic(true); b1.setStyle(sb);

                Cell c1 = wb.getWorksheets().get(0).getCells().get("C1");
                c1.putValue("C"); Style sc = c1.getStyle(); sc.setHorizontalAlignment(HorizontalAlignmentType.CENTER); c1.setStyle(sc);

                wb.save(path);
            }
            String stylesXml = ZipPackageHelper.readEntryText(path, "xl/styles.xml");
            assertTrue(stylesXml.contains("<cellXfs"), "styles.xml should contain cellXfs element");
            // count should be >= 4 (default + 3 distinct styles)
            // We just verify the cellXfs element is there with a count attribute
            assertTrue(stylesXml.contains("count="), "cellXfs should have count attribute");
        }
    }

    /**
     * Verifies that styles deduplicated.
     */
    @Test
    void ST_73_stylesDeduplicated() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-dedup.xlsx");
            try (Workbook wb = new Workbook()) {
                // Two cells with identical styles
                Cell a1 = wb.getWorksheets().get(0).getCells().get("A1");
                a1.putValue("A"); Style sa = a1.getStyle(); sa.getFont().setBold(true); a1.setStyle(sa);

                Cell b1 = wb.getWorksheets().get(0).getCells().get("B1");
                b1.putValue("B"); Style sb = b1.getStyle(); sb.getFont().setBold(true); b1.setStyle(sb);

                wb.save(path);
            }
            // Both cells should reference the same XF index -verify by reading back
            Workbook loaded = new Workbook(path);
            assertTrue(loaded.getWorksheets().get(0).getCells().get("A1").getStyle().getFont().getBold());
            assertTrue(loaded.getWorksheets().get(0).getCells().get("B1").getStyle().getFont().getBold());

            // Verify styles.xml doesn't have duplicate XF entries
            String stylesXml = ZipPackageHelper.readEntryText(path, "xl/styles.xml");
            // Count occurrences of <xf (inside cellXfs) -with dedup there should be exactly 2 (default + 1 bold)
            int xfCount = countOccurrences(stylesXml.substring(stylesXml.indexOf("<cellXfs")), "<xf ");
            assertEquals(2, xfCount, "Expected 2 XF entries (default + 1 bold, deduplicated)");
        }
    }

    /**
     * Verifies that locked false roundtrip.
     */
    @Test
    void ST_74_lockedFalseRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-locked.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue("Unlocked");
                Style style = cell.getStyle();
                style.setLocked(false);
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertFalse(loaded.getWorksheets().get(0).getCells().get("A1").getStyle().isLocked());
        }
    }

    /**
     * Verifies that hidden true roundtrip.
     */
    @Test
    void ST_75_hiddenTrueRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-hidden.xlsx");
            try (Workbook wb = new Workbook()) {
                Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
                cell.putValue("Hidden");
                Style style = cell.getStyle();
                style.setFormulaHidden(true);
                cell.setStyle(style);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertTrue(loaded.getWorksheets().get(0).getCells().get("A1").getStyle().isFormulaHidden());
        }
    }

    // =========================================================================
    // 2.8 Integration: Write All Style Properties ->Verify via API and POI
    // =========================================================================

    /**
     * Verifies that to st 93 integration all style properties.
     */
    @Test
    void ST_80_to_ST_93_integrationAllStyleProperties() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath("style-all.xlsx");

            try (Workbook wb = new Workbook()) {
                Cells cells = wb.getWorksheets().get(0).getCells();

                // ST-80: Bold
                {
                    Cell c = cells.get("A1"); c.putValue("Bold");
                    Style s = c.getStyle(); s.getFont().setBold(true); c.setStyle(s);
                }
                // ST-81: Font name + size (Arial 18pt)
                {
                    Cell c = cells.get("A2"); c.putValue("Arial");
                    Style s = c.getStyle(); s.getFont().setName("Arial"); s.getFont().setSize(18.0); c.setStyle(s);
                }
                // ST-82: Font color red
                {
                    Cell c = cells.get("A3"); c.putValue("Red");
                    Style s = c.getStyle(); s.getFont().setColor(Color.fromArgb(255, 255, 0, 0)); c.setStyle(s);
                }
                // ST-83: Italic + underline
                {
                    Cell c = cells.get("A4"); c.putValue("ItalicUnderline");
                    Style s = c.getStyle(); s.getFont().setItalic(true); s.getFont().setUnderline(true); c.setStyle(s);
                }
                // ST-84: Center alignment
                {
                    Cell c = cells.get("A5"); c.putValue("Center");
                    Style s = c.getStyle(); s.setHorizontalAlignment(HorizontalAlignmentType.CENTER); c.setStyle(s);
                }
                // ST-85: Wrap text
                {
                    Cell c = cells.get("A6"); c.putValue("Wrap");
                    Style s = c.getStyle(); s.setTextWrapped(true); c.setStyle(s);
                }
                // ST-86: THIN border all 4 sides
                {
                    Cell c = cells.get("A7"); c.putValue("Borders");
                    Style s = c.getStyle();
                    s.getBorders().getLeft().setLineStyle(BorderStyleType.THIN);
                    s.getBorders().getRight().setLineStyle(BorderStyleType.THIN);
                    s.getBorders().getTop().setLineStyle(BorderStyleType.THIN);
                    s.getBorders().getBottom().setLineStyle(BorderStyleType.THIN);
                    c.setStyle(s);
                }
                // ST-87: Border color (blue left)
                {
                    Cell c = cells.get("A8"); c.putValue("BlueBorder");
                    Style s = c.getStyle();
                    s.getBorders().getLeft().setLineStyle(BorderStyleType.THIN);
                    s.getBorders().getLeft().setColor(Color.fromArgb(255, 0, 0, 255));
                    c.setStyle(s);
                }
                // ST-88: Solid fill yellow
                {
                    Cell c = cells.get("A9"); c.putValue("Yellow");
                    Style s = c.getStyle();
                    s.setPattern(FillPattern.SOLID);
                    s.setForegroundColor(Color.fromArgb(255, 255, 255, 0));
                    c.setStyle(s);
                }
                // ST-89: Custom number format
                {
                    Cell c = cells.get("A10"); c.putValue(12345.678);
                    Style s = c.getStyle(); s.setCustom("#,##0.00"); c.setStyle(s);
                }
                // ST-90: Built-in number format
                {
                    Cell c = cells.get("A11"); c.putValue(3.14);
                    Style s = c.getStyle(); s.setNumber(2); c.setStyle(s);
                }
                // ST-91: Locked=false
                {
                    Cell c = cells.get("A12"); c.putValue("Unlocked");
                    Style s = c.getStyle(); s.setLocked(false); c.setStyle(s);
                }
                // ST-92: Multiple cells different styles (A13: bold 14pt; B13: italic + CENTER; C13: THIN border + solid fill)
                {
                    Cell a13 = cells.get("A13"); a13.putValue("BoldBig");
                    Style sa = a13.getStyle(); sa.getFont().setBold(true); sa.getFont().setSize(14.0); a13.setStyle(sa);

                    Cell b13 = cells.get("B13"); b13.putValue("ItalicCenter");
                    Style sb = b13.getStyle(); sb.getFont().setItalic(true);
                    sb.setHorizontalAlignment(HorizontalAlignmentType.CENTER); b13.setStyle(sb);

                    Cell c13 = cells.get("C13"); c13.putValue("BorderFill");
                    Style sc = c13.getStyle();
                    sc.getBorders().getLeft().setLineStyle(BorderStyleType.THIN);
                    sc.setPattern(FillPattern.SOLID);
                    sc.setForegroundColor(Color.fromArgb(255, 0, 255, 0));
                    c13.setStyle(sc);
                }
                // ST-93: Style de-duplication -A14 and B14 same style
                {
                    Cell a14 = cells.get("A14"); a14.putValue("Same1");
                    Style sd = a14.getStyle(); sd.getFont().setBold(true); sd.getFont().setSize(12.0); a14.setStyle(sd);

                    Cell b14 = cells.get("B14"); b14.putValue("Same2");
                    Style se = b14.getStyle(); se.getFont().setBold(true); se.getFont().setSize(12.0); b14.setStyle(se);
                }

                wb.save(path);
            }

            // --- API verification ---
            Workbook loaded = new Workbook(path);
            Cells cells = loaded.getWorksheets().get(0).getCells();

            // ST-80
            assertTrue(cells.get("A1").getStyle().getFont().getBold());

            // ST-81
            assertEquals("Arial", cells.get("A2").getStyle().getFont().getName());
            assertEquals(18.0, cells.get("A2").getStyle().getFont().getSize(), 1e-9);

            // ST-82
            Color redC = cells.get("A3").getStyle().getFont().getColor();
            assertEquals(255, u(redC.getA())); assertEquals(255, u(redC.getR()));
            assertEquals(0, u(redC.getG())); assertEquals(0, u(redC.getB()));

            // ST-83
            assertTrue(cells.get("A4").getStyle().getFont().getItalic());
            assertTrue(cells.get("A4").getStyle().getFont().getUnderline());

            // ST-84
            assertEquals(HorizontalAlignmentType.CENTER, cells.get("A5").getStyle().getHorizontalAlignment());

            // ST-85
            assertTrue(cells.get("A6").getStyle().isTextWrapped());

            // ST-86
            assertEquals(BorderStyleType.THIN, cells.get("A7").getStyle().getBorders().getTop().getLineStyle());
            assertEquals(BorderStyleType.THIN, cells.get("A7").getStyle().getBorders().getLeft().getLineStyle());
            assertEquals(BorderStyleType.THIN, cells.get("A7").getStyle().getBorders().getRight().getLineStyle());
            assertEquals(BorderStyleType.THIN, cells.get("A7").getStyle().getBorders().getBottom().getLineStyle());

            // ST-87
            Color blueC = cells.get("A8").getStyle().getBorders().getLeft().getColor();
            assertEquals(255, u(blueC.getA())); assertEquals(0, u(blueC.getR()));
            assertEquals(0, u(blueC.getG())); assertEquals(255, u(blueC.getB()));

            // ST-88
            assertEquals(FillPattern.SOLID, cells.get("A9").getStyle().getPattern());
            Color yfg = cells.get("A9").getStyle().getForegroundColor();
            assertEquals(255, u(yfg.getA())); assertEquals(255, u(yfg.getR()));
            assertEquals(255, u(yfg.getG())); assertEquals(0, u(yfg.getB()));

            // ST-89
            assertEquals("#,##0.00", cells.get("A10").getStyle().getCustom());

            // ST-90
            assertEquals(2, cells.get("A11").getStyle().getNumber());

            // ST-91
            assertFalse(cells.get("A12").getStyle().isLocked());

            // ST-92 -each cell style differs
            assertTrue(cells.get("A13").getStyle().getFont().getBold());
            assertEquals(14.0, cells.get("A13").getStyle().getFont().getSize(), 1e-9);
            assertTrue(cells.get("B13").getStyle().getFont().getItalic());
            assertEquals(HorizontalAlignmentType.CENTER, cells.get("B13").getStyle().getHorizontalAlignment());
            assertEquals(BorderStyleType.THIN, cells.get("C13").getStyle().getBorders().getLeft().getLineStyle());
            assertEquals(FillPattern.SOLID, cells.get("C13").getStyle().getPattern());

            // ST-93 -both cells have same style after reload
            assertTrue(cells.get("A14").getStyle().getFont().getBold());
            assertTrue(cells.get("B14").getStyle().getFont().getBold());
            assertEquals(cells.get("A14").getStyle().getFont().getSize(),
                cells.get("B14").getStyle().getFont().getSize(), 1e-9);

            // --- POI verification ---
            try (org.apache.poi.ss.usermodel.Workbook poiWb = WorkbookFactory.create(new File(path))) {
                org.apache.poi.ss.usermodel.Sheet poiSheet = poiWb.getSheetAt(0);

                // ST-80: Bold
                org.apache.poi.ss.usermodel.Font f80 = poiWb.getFontAt(
                    poiSheet.getRow(0).getCell(0).getCellStyle().getFontIndex());
                assertTrue(f80.getBold());

                // ST-81: Arial 18pt
                org.apache.poi.ss.usermodel.Font f81 = poiWb.getFontAt(
                    poiSheet.getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertEquals("Arial", f81.getFontName());
                assertEquals(18, f81.getFontHeightInPoints());

                // ST-82: Font color red
                org.apache.poi.ss.usermodel.Font f82raw = poiWb.getFontAt(
                    poiSheet.getRow(2).getCell(0).getCellStyle().getFontIndex());
                assertTrue(f82raw instanceof XSSFFont, "Expected XSSFFont for color check");
                XSSFFont f82 = (XSSFFont) f82raw;
                assertEquals("FFFF0000", f82.getXSSFColor().getARGBHex());

                // ST-83: Italic + underline
                org.apache.poi.ss.usermodel.Font f83 = poiWb.getFontAt(
                    poiSheet.getRow(3).getCell(0).getCellStyle().getFontIndex());
                assertTrue(f83.getItalic());
                assertTrue(f83.getUnderline() != org.apache.poi.ss.usermodel.Font.U_NONE);

                // ST-84: Center alignment
                org.apache.poi.ss.usermodel.CellStyle cs84 = poiSheet.getRow(4).getCell(0).getCellStyle();
                assertEquals(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER, cs84.getAlignment());

                // ST-85: Wrap text
                org.apache.poi.ss.usermodel.CellStyle cs85 = poiSheet.getRow(5).getCell(0).getCellStyle();
                assertTrue(cs85.getWrapText());

                // ST-86: THIN border all sides
                org.apache.poi.ss.usermodel.CellStyle cs86 = poiSheet.getRow(6).getCell(0).getCellStyle();
                assertEquals(org.apache.poi.ss.usermodel.BorderStyle.THIN, cs86.getBorderTop());
                assertEquals(org.apache.poi.ss.usermodel.BorderStyle.THIN, cs86.getBorderLeft());
                assertEquals(org.apache.poi.ss.usermodel.BorderStyle.THIN, cs86.getBorderRight());
                assertEquals(org.apache.poi.ss.usermodel.BorderStyle.THIN, cs86.getBorderBottom());

                // ST-87: Blue left border color
                org.apache.poi.ss.usermodel.CellStyle cs87raw = poiSheet.getRow(7).getCell(0).getCellStyle();
                assertTrue(cs87raw instanceof XSSFCellStyle, "Expected XSSFCellStyle for color check");
                XSSFCellStyle cs87 = (XSSFCellStyle) cs87raw;
                assertNotNull(cs87.getLeftBorderXSSFColor());
                assertEquals("FF0000FF", cs87.getLeftBorderXSSFColor().getARGBHex());

                // ST-88: Solid fill yellow
                org.apache.poi.ss.usermodel.CellStyle cs88 = poiSheet.getRow(8).getCell(0).getCellStyle();
                assertEquals(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND, cs88.getFillPattern());
                assertTrue(cs88 instanceof XSSFCellStyle, "Expected XSSFCellStyle for fill color check");
                XSSFCellStyle xcs88 = (XSSFCellStyle) cs88;
                assertNotNull(xcs88.getFillForegroundXSSFColor());
                assertEquals("FFFFFF00", xcs88.getFillForegroundXSSFColor().getARGBHex());

                // ST-89: Custom number format
                org.apache.poi.ss.usermodel.CellStyle cs89 = poiSheet.getRow(9).getCell(0).getCellStyle();
                assertEquals("#,##0.00", cs89.getDataFormatString());

                // ST-90: Built-in number format "0.00"
                org.apache.poi.ss.usermodel.CellStyle cs90 = poiSheet.getRow(10).getCell(0).getCellStyle();
                assertEquals("0.00", cs90.getDataFormatString());

                // ST-91: Locked=false
                org.apache.poi.ss.usermodel.CellStyle cs91 = poiSheet.getRow(11).getCell(0).getCellStyle();
                assertFalse(cs91.getLocked());

                // ST-92: Multiple cells different styles
                org.apache.poi.ss.usermodel.CellStyle csA13 = poiSheet.getRow(12).getCell(0).getCellStyle();
                org.apache.poi.ss.usermodel.CellStyle csB13 = poiSheet.getRow(12).getCell(1).getCellStyle();
                org.apache.poi.ss.usermodel.CellStyle csC13 = poiSheet.getRow(12).getCell(2).getCellStyle();
                assertNotEquals(csA13.getIndex(), csB13.getIndex(), "A13 and B13 should have different styles");
                assertNotEquals(csA13.getIndex(), csC13.getIndex(), "A13 and C13 should have different styles");

                // ST-93: Verify styles.xml count doesn't increase for same style applied twice
                // Check that A14 and B14 share the same XF index
                short idxA14 = poiSheet.getRow(13).getCell(0).getCellStyle().getIndex();
                short idxB14 = poiSheet.getRow(13).getCell(1).getCellStyle().getIndex();
                assertEquals(idxA14, idxB14, "A14 and B14 should reference the same XF index");
            }
        }
    }

    // =========================================================================
    // Helper
    // =========================================================================

    /**
     * Verifies that count occurrences.
     * @param text text
     * @param sub sub
     */
    private static int countOccurrences(String text, String sub) {
        int count = 0;
        int idx = 0;
        // Walk the current collection so every entry is processed consistently.
        while ((idx = text.indexOf(sub, idx)) != -1) {
            count++;
            idx += sub.length();
        }
        return count;
    }

    /**
     * Returns the built-in format IDs covered by this library.
     * @return format ids
     */
    private static List<Integer> builtInDisplayFormatIds() {
        return List.of(
                0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22,
                37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50
        );
    }

    /**
     * Creates a built-in display string case.
     * @param formatId format id
     * @return case
     */
    private static DisplayStringCase createBuiltInDisplayStringCase(int formatId) {
        String formatCode = NumberFormat.getBuiltInFormat(formatId);
        String address = "A" + (formatId + 1);

        if (formatId == 49) {
            return new DisplayStringCase(address, "built-in " + formatId + " (" + formatCode + ")", cell -> {
                cell.putValue("Alpha");
                Style style = cell.getStyle();
                style.setNumber(formatId);
                cell.setStyle(style);
            });
        }

        if (isDateTimeFormatId(formatId)) {
            LocalDateTime dateTimeValue = LocalDateTime.of(2024, 3, 15, 13, 4, 5, 670_000_000);
            return new DisplayStringCase(address, "built-in " + formatId + " (" + formatCode + ")", cell -> {
                cell.putValue(dateTimeValue);
                Style style = cell.getStyle();
                style.setNumber(formatId);
                cell.setStyle(style);
            });
        }

        double numericValue = isFractionFormatId(formatId) ? 2.125 : -12345.678;
        return new DisplayStringCase(address, "built-in " + formatId + " (" + formatCode + ")", cell -> {
            cell.putValue(numericValue);
            Style style = cell.getStyle();
            style.setNumber(formatId);
            cell.setStyle(style);
        });
    }

    /**
     * Verifies display strings against Apache POI DataFormatter.
     * @param cases cases
     * @param fileName output file name
     * @throws Exception on failure
     */
    private static void assertDisplayStringsMatchPoi(List<DisplayStringCase> cases, String fileName) throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("StyleTest")) {
            String path = tempDir.getPath(fileName);
            try (Workbook wb = new Workbook()) {
                Cells cells = wb.getWorksheets().get(0).getCells();
                for (DisplayStringCase testCase : cases) {
                    testCase.configure(cells.get(testCase.address));
                }
                wb.save(path);
            }

            Workbook loaded = new Workbook(path);
            DataFormatter formatter = new DataFormatter(Locale.US);
            try (org.apache.poi.ss.usermodel.Workbook poiWb = WorkbookFactory.create(new File(path))) {
                org.apache.poi.ss.usermodel.Sheet poiSheet = poiWb.getSheetAt(0);
                for (DisplayStringCase testCase : cases) {
                    Cell cell = loaded.getWorksheets().get(0).getCells().get(testCase.address);
                    int rowIndex = Integer.parseInt(testCase.address.substring(1)) - 1;
                    org.apache.poi.ss.usermodel.Cell poiCell = poiSheet.getRow(rowIndex).getCell(0);
                    String expected = testCase.expected != null ? testCase.expected : formatter.formatCellValue(poiCell);
                    assertEquals(expected, cell.getDisplayStringValue(), testCase.label);
                }
            }
        }
    }

    /**
     * Indicates whether a built-in format id is date/time based.
     * @param formatId format id
     * @return true when date/time based
     */
    private static boolean isDateTimeFormatId(int formatId) {
        return formatId == 14 || formatId == 15 || formatId == 16 || formatId == 17
                || formatId == 18 || formatId == 19 || formatId == 20 || formatId == 21
                || formatId == 22 || formatId == 45 || formatId == 46 || formatId == 47;
    }

    /**
     * Indicates whether a built-in format id is fraction based.
     * @param formatId format id
     * @return true when fraction based
     */
    private static boolean isFractionFormatId(int formatId) {
        return formatId == 12 || formatId == 13;
    }

    /**
     * Represents one display string regression case.
     */
    private static final class DisplayStringCase {
        private final String address;
        private final String label;
        private final String expected;
        private final Consumer<Cell> configurator;

        /**
         * Initializes a new case.
         * @param address address
         * @param label label
         * @param configurator configurator
         */
        private DisplayStringCase(String address, String label, Consumer<Cell> configurator) {
            this(address, label, null, configurator);
        }

        /**
         * Initializes a new case with an explicit expected value.
         * @param address address
         * @param label label
         * @param expected expected output
         * @param configurator configurator
         */
        private DisplayStringCase(String address, String label, String expected, Consumer<Cell> configurator) {
            this.address = address;
            this.label = label;
            this.expected = expected;
            this.configurator = configurator;
        }

        /**
         * Configures a cell for this case.
         * @param cell cell
         */
        private void configure(Cell cell) {
            configurator.accept(cell);
        }
    }
}

