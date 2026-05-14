package com.aspose.cells_foss;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Factory for style test workbooks. Mirrors C# StyleScenarioFactory.
 */
public final class StyleScenarioFactory {

    private StyleScenarioFactory() {}

    /** Applies the "primary" style covering all public style settings. */
    public static void applyPrimaryStyle(Style style) {
        style.getFont().setName("Arial");
        style.getFont().setSize(12.0);
        style.getFont().setBold(true);
        style.getFont().setItalic(true);
        style.getFont().setUnderline(true);
        style.getFont().setStrikeThrough(true);
        style.getFont().setColor(Color.fromArgb(255, 255, 0, 0));
        style.getBorders().getRight().setLineStyle(BorderStyleType.MEDIUM_DASH_DOT);
        style.getBorders().getDiagonal().setLineStyle(BorderStyleType.SLANTED_DASH_DOT);
        style.getBorders().setDiagonalUp(true);
        style.getBorders().setDiagonalDown(true);
        style.setPattern(FillPattern.LIGHT_GRID);
        style.setForegroundColor(Color.fromArgb(255, 0, 128, 0));
        style.setBackgroundColor(Color.fromArgb(255, 255, 255, 0));
        style.setNumber(4);
        style.setHorizontalAlignment(HorizontalAlignmentType.DISTRIBUTED);
        style.setVerticalAlignment(VerticalAlignmentType.DISTRIBUTED);
        style.setTextWrapped(true);
        style.setIndentLevel(2);
        style.setRotationAngle(45);
        style.setShrinkToFit(true);
        style.setReadingOrder(2);
        style.setRelativeIndent(1);
        style.setLocked(false);
        style.setFormulaHidden(true);
    }

    /** Asserts that the style matches the primary style set by applyPrimaryStyle. */
    public static void assertPrimaryStyle(Style style) {
        assertEquals("Arial", style.getFont().getName());
        assertTrue(style.getFont().getBold());
        assertTrue(style.getFont().getItalic());
        assertTrue(style.getFont().getUnderline());
        assertTrue(style.getFont().getStrikeThrough());
        assertEquals(BorderStyleType.MEDIUM_DASH_DOT, style.getBorders().getRight().getLineStyle());
        assertEquals(BorderStyleType.SLANTED_DASH_DOT, style.getBorders().getDiagonal().getLineStyle());
        assertTrue(style.getBorders().getDiagonalUp());
        assertTrue(style.getBorders().getDiagonalDown());
        assertEquals(FillPattern.LIGHT_GRID, style.getPattern());
        assertEquals(4, style.getNumber());
        assertEquals("#,##0.00", style.getNumberFormat());
        assertEquals(HorizontalAlignmentType.DISTRIBUTED, style.getHorizontalAlignment());
        assertEquals(VerticalAlignmentType.DISTRIBUTED, style.getVerticalAlignment());
        assertTrue(style.isTextWrapped());
        assertEquals(2, style.getIndentLevel());
        assertEquals(45, style.getRotationAngle());
        assertTrue(style.getShrinkToFit());
        assertEquals(2, style.getReadingOrder());
        assertEquals(1, style.getRelativeIndent());
        assertFalse(style.isLocked());
        assertTrue(style.isFormulaHidden());
    }

    /** Applies a custom number format style. */
    public static void applyCustomNumberStyle(Style style) {
        style.setCustom("[Blue]0.000");
    }

    /** Asserts the custom number format style. */
    public static void assertCustomNumberStyle(Style style) {
        assertEquals("[Blue]0.000", style.getCustom());
        assertEquals(0, style.getNumber());
    }

    /** Creates a workbook with primary and custom styles on specific cells. */
    public static Workbook createStyledWorkbook() {
        Workbook wb = WorkbookScenarioFactory.createMixedCellWorkbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        Cell primary = sheet.getCells().get("A1");
        Style ps = primary.getStyle();
        applyPrimaryStyle(ps);
        primary.setStyle(ps);

        Cell custom = sheet.getCells().get("B2");
        Style cs = custom.getStyle();
        applyCustomNumberStyle(cs);
        custom.setStyle(cs);

        return wb;
    }
}
