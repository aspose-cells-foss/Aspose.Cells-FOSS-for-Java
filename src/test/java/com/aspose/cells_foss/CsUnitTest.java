package com.aspose.cells_foss;

import com.aspose.cells_foss.core.CellAddress;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Java port of C# UnitTests/Program.cs.
 * Covers cell data, styles, worksheet APIs, auto-filter, defined names, hyperlinks,
 * validations, conditional formatting, page setup, list objects, pictures, comments, and charts.
 */
class CsUnitTest {

    @TempDir Path tempDir;

    // =========================================================================
    // Cell data
    // =========================================================================

    @Test
    void a1_indexers_roundtrip() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.getCells().get(2, 27).putValue("AB3");
        sheet.getCells().get(0, 0).putValue(42);

        assertEquals("AB3", sheet.getCells().get("AB3").getStringValue());
        assertEquals("42", sheet.getCells().get(0, 0).getStringValue());
        assertEquals(CellValueType.STRING, sheet.getCells().get("AB3").getType());
        assertEquals(CellValueType.NUMBER, sheet.getCells().get("A1").getType());
    }

    @Test
    void put_value_overloads_assign_expected_types() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        LocalDateTime ts = LocalDateTime.of(2024, 5, 6, 7, 8, 9);

        sheet.getCells().get("A1").putValue("alpha");
        sheet.getCells().get("B1").putValue(123);
        sheet.getCells().get("C1").putValue(12.5);
        sheet.getCells().get("D1").putValue(6.02214076E+23);
        sheet.getCells().get("E1").putValue(true);
        sheet.getCells().get("F1").putValue(ts);

        assertEquals(CellValueType.STRING,   sheet.getCells().get("A1").getType());
        assertEquals(CellValueType.NUMBER,   sheet.getCells().get("B1").getType());
        assertEquals(CellValueType.NUMBER,   sheet.getCells().get("C1").getType());
        assertEquals(CellValueType.NUMBER,   sheet.getCells().get("D1").getType());
        assertEquals(CellValueType.BOOLEAN,  sheet.getCells().get("E1").getType());
        assertEquals(CellValueType.DATE_TIME, sheet.getCells().get("F1").getType());

        assertEquals("alpha", sheet.getCells().get("A1").getValue());
        assertEquals(123, ((Number)sheet.getCells().get("B1").getValue()).intValue());
        assertTrue(Math.abs((Double)sheet.getCells().get("D1").getValue() - 6.02214076E+23) < 1E+10);
        assertEquals(true, sheet.getCells().get("E1").getValue());
        assertEquals(ts, sheet.getCells().get("F1").getValue());
    }

    @Test
    void stringvalue_formats_supported_scalar_types() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        LocalDateTime ts = LocalDateTime.of(2024, 5, 6, 7, 8, 9);

        sheet.getCells().get("A1").putValue(true);
        sheet.getCells().get("B1").putValue(123);
        sheet.getCells().get("C1").putValue(12.5);
        sheet.getCells().get("D1").putValue(ts);

        assertEquals("TRUE",  sheet.getCells().get("A1").getStringValue());
        assertEquals("123",   sheet.getCells().get("B1").getStringValue());
        assertEquals("12.5",  sheet.getCells().get("C1").getStringValue());
        assertFalse(sheet.getCells().get("D1").getStringValue().isBlank());
    }

    @Test
    void formula_property_normalizes_and_preserves_cached_value() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("C3");
        cell.putValue(20);
        cell.setFormula("A1+B1");

        assertEquals(CellValueType.FORMULA, cell.getType());
        assertEquals("=A1+B1", cell.getFormula());
        assertEquals("20", cell.getStringValue());
    }

    @Test
    void blank_cells_are_blank_by_default() {
        Cell cell = new Workbook().getWorksheets().get(0).getCells().get("Z99");
        assertEquals(CellValueType.BLANK, cell.getType());
        assertNull(cell.getValue());
        assertEquals("", cell.getStringValue());
    }

    // =========================================================================
    // Worksheet name & collection guards
    // =========================================================================

    @Test
    void worksheet_name_and_collection_guards() {
        Workbook wb = new Workbook();
        wb.getWorksheets().add("Data");

        assertThrows(CellsException.class, () -> wb.getWorksheets().add("data"));
        assertThrows(CellsException.class, () -> wb.getWorksheets().get(0).getCells().get("1A"));
        assertThrows(CellsException.class, () -> wb.getWorksheets().get(0).getCells().get(-1, 0));

        CellAddress parsed = CellAddress.parse("AB3");
        assertEquals(2,  parsed.getRowIndex());
        assertEquals(27, parsed.getColumnIndex());
        assertEquals("AB3", parsed.toString());
    }

    // =========================================================================
    // Style
    // =========================================================================

    @Test
    void style_mutation_requires_setstyle_and_returns_clones() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");

        Style style = cell.getStyle();
        style.getFont().setBold(true);
        style.setHorizontalAlignment(HorizontalAlignmentType.RIGHT);

        Style untouched = cell.getStyle();
        assertFalse(untouched.getFont().getBold());
        assertEquals(HorizontalAlignmentType.GENERAL, untouched.getHorizontalAlignment());

        cell.setStyle(style);
        Style applied = cell.getStyle();
        assertTrue(applied.getFont().getBold());
        assertEquals(HorizontalAlignmentType.RIGHT, applied.getHorizontalAlignment());

        applied.getFont().setItalic(true);
        assertFalse(cell.getStyle().getFont().getItalic());
    }

    @Test
    void style_api_covers_all_public_settings() {
        Workbook wb = new Workbook();
        Cell primaryCell = wb.getWorksheets().get(0).getCells().get("A1");
        primaryCell.putValue(1);

        Style ps = primaryCell.getStyle();
        StyleScenarioFactory.applyPrimaryStyle(ps);

        // Not yet applied — cell still has default style
        Style untouched = primaryCell.getStyle();
        assertEquals("Calibri", untouched.getFont().getName());

        primaryCell.setStyle(ps);
        StyleScenarioFactory.assertPrimaryStyle(primaryCell.getStyle());

        // NumberFormat combined setter / getter
        Style nf = new Style();
        nf.setNumberFormat("0.00%");
        assertEquals(10, nf.getNumber());
        assertNull(nf.getCustom());
        nf.setNumberFormat("[Blue]0.000");
        assertEquals(0, nf.getNumber());
        assertEquals("[Blue]0.000", nf.getCustom());

        assertThrows(CellsException.class, () -> nf.setIndentLevel(-1));
        assertThrows(CellsException.class, () -> nf.setRotationAngle(181));
        assertThrows(CellsException.class, () -> nf.setReadingOrder(3));
    }

    // =========================================================================
    // Worksheet row/column/merge
    // =========================================================================

    @Test
    void worksheet_row_column_and_merge_apis() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        sheet.setVisibilityType(VisibilityType.HIDDEN);
        sheet.getCells().getRows().get(2).setHeight(19.75);
        sheet.getCells().getRows().get(4).setHidden(true);
        sheet.getCells().getColumns().get(1).setWidth(25.5d);
        sheet.getCells().getColumns().get(3).setHidden(true);
        sheet.getCells().merge(1, 1, 2, 3);

        assertEquals(VisibilityType.HIDDEN, sheet.getVisibilityType());
        assertEquals(19.75, sheet.getCells().getRows().get(2).getHeight(), 1e-9);
        assertTrue(sheet.getCells().getRows().get(4).isHidden());
        assertEquals(25.5, sheet.getCells().getColumns().get(1).getWidth(), 1e-9);
        assertTrue(sheet.getCells().getColumns().get(3).isHidden());
        assertEquals(1, sheet.getCells().getMergedCells().size());
        assertThrows(CellsException.class, () -> sheet.getCells().merge(2, 2, 2, 2));
    }

    @Test
    void worksheet_view_apis() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

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

        sheet.setTabColor(Color.getEmpty());
        assertEquals(Color.getEmpty(), sheet.getTabColor());
        assertThrows(CellsException.class, () -> sheet.setZoom(9));
        assertThrows(CellsException.class, () -> sheet.setZoom(401));
    }

    @Test
    void worksheet_protection_apis() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        sheet.protect();
        sheet.getProtection().setAllowEditingObject(false);
        sheet.getProtection().setAllowFormattingCell(false);
        sheet.getProtection().setAllowInsertingRow(false);
        sheet.getProtection().setAllowSelectingUnlockedCell(false);

        assertTrue(sheet.getProtection().isProtected());
        assertFalse(sheet.getProtection().getAllowEditingObject());
        assertFalse(sheet.getProtection().getAllowFormattingCell());
        assertFalse(sheet.getProtection().getAllowInsertingRow());
        assertFalse(sheet.getProtection().getAllowSelectingUnlockedCell());

        sheet.unprotect();
        assertFalse(sheet.getProtection().isProtected());
        assertTrue(sheet.getProtection().getAllowEditingObject());

        sheet.getProtection().setAllowFiltering(false);
        assertTrue(sheet.getProtection().isProtected());
        assertFalse(sheet.getProtection().getAllowFiltering());
    }

    // =========================================================================
    // AutoFilter
    // =========================================================================

    @Test
    void autofilter_apis_mutate_expected_settings() {
        Workbook wb = AutoFilterScenarioFactory.createAutoFilterWorkbook();
        AutoFilterScenarioFactory.assertAutoFilter(wb);

        Worksheet sheet = wb.getWorksheets().get(0);
        assertThrows(CellsException.class, () -> sheet.getAutoFilter().getFilterColumns().add(-1));
        assertThrows(CellsException.class, () -> sheet.getAutoFilter().getFilterColumns().add(0));
        assertThrows(CellsException.class, () -> sheet.getAutoFilter().getSortState().getSortConditions().add("1A"));

        sheet.getAutoFilter().getFilterColumns().removeAt(4);
        assertEquals(4, sheet.getAutoFilter().getFilterColumns().getCount());
        sheet.getAutoFilter().clear();
        assertEquals("", sheet.getAutoFilter().getRange());
        assertEquals(0, sheet.getAutoFilter().getFilterColumns().getCount());
        assertEquals(0, sheet.getAutoFilter().getSortState().getSortConditions().getCount());
    }

    // =========================================================================
    // Defined names
    // =========================================================================

    @Test
    void defined_name_apis_mutate_expected_settings() {
        Workbook wb = new Workbook();
        wb.getWorksheets().get(0).setName("Data");
        wb.getWorksheets().add("Scoped");

        int i1 = wb.getDefinedNames().add("Total", "=SUM(Data!$A$1:$A$2)");
        DefinedName total = wb.getDefinedNames().get(i1);
        total.setHidden(true);
        total.setComment("Workbook scope");

        int i2 = wb.getDefinedNames().add("Input", "'Scoped'!$B$2", 1);
        DefinedName scoped = wb.getDefinedNames().get(i2);
        scoped.setComment("Local scope");

        assertEquals(2, wb.getDefinedNames().getCount());
        assertEquals("Total", total.getName());
        assertTrue(total.getFormula().contains("SUM"));
        assertNull(total.getLocalSheetIndex());
        assertTrue(total.isHidden());
        assertEquals("Workbook scope", total.getComment());

        assertEquals("Input", scoped.getName());
        assertEquals(Integer.valueOf(1), scoped.getLocalSheetIndex());
        assertEquals("Local scope", scoped.getComment());

        assertThrows(CellsException.class, () -> wb.getDefinedNames().add("Total", "1"));
        assertThrows(CellsException.class, () -> wb.getDefinedNames().add("_xlnm.Print_Area", "A1"));
        assertThrows(CellsException.class, () -> wb.getDefinedNames().add("Broken", "1", 5));
    }

    // =========================================================================
    // Hyperlinks
    // =========================================================================

    @Test
    void hyperlink_apis_mutate_expected_settings() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        sheet.getCells().get("A1").putValue("Docs");
        int ei = sheet.getHyperlinks().add("A1", 1, 1, "https://example.com/docs");
        Hyperlink ext = sheet.getHyperlinks().get(ei);
        ext.setTextToDisplay("Docs");
        ext.setScreenTip("External docs");

        int ii = sheet.getHyperlinks().add("B2", 1, 1, "Sheet1!C3");
        sheet.getHyperlinks().get(ii).setTextToDisplay("Jump");

        int ri = sheet.getHyperlinks().add("C4", 2, 2, "mailto:test@example.com");
        sheet.getHyperlinks().get(ri).setScreenTip("Send mail");

        assertEquals(3, sheet.getHyperlinks().getCount());
        assertEquals("https://example.com/docs", ext.getAddress());
        assertEquals("External docs", ext.getScreenTip());
        assertEquals("Docs", ext.getTextToDisplay());

        assertThrows(CellsException.class, () -> sheet.getHyperlinks().add("A1", 1, 1, "https://overlap.example.com"));
        assertThrows(CellsException.class, () -> sheet.getHyperlinks().add("Z1", 0, 1, "https://invalid.example.com"));
        assertThrows(CellsException.class, () -> sheet.getHyperlinks().add("A2", 1, 1, ""));
        assertThrows(CellsException.class, () -> sheet.getHyperlinks().get(-1));

        sheet.getHyperlinks().removeAt(1);
        assertEquals(2, sheet.getHyperlinks().getCount());
    }

    // =========================================================================
    // Data validation
    // =========================================================================

    @Test
    void validation_apis_mutate_expected_settings() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        int pi = sheet.getValidations().add(CellArea.createCellArea("A1", "B2"));
        Validation primary = sheet.getValidations().get(pi);
        primary.setType(ValidationType.WHOLE_NUMBER);
        primary.setOperator(OperatorType.BETWEEN);
        primary.setFormula1("=1");
        primary.setFormula2("=10");
        primary.setShowError(true);
        primary.setErrorTitle("Whole Number");
        primary.setErrorMessage("Enter 1-10");
        primary.addArea(CellArea.createCellArea("D4", "D5"));

        assertEquals(1, sheet.getValidations().getCount());
        assertEquals(2, primary.getAreas().size());
        assertEquals("1", primary.getFormula1());
        assertEquals("10", primary.getFormula2());
        assertEquals(ValidationType.WHOLE_NUMBER, sheet.getValidations().getValidationInCell(0, 0).getType());
        assertEquals(ValidationType.WHOLE_NUMBER, sheet.getValidations().getValidationInCell(4, 3).getType());
        assertThrows(CellsException.class, () -> sheet.getValidations().add(CellArea.createCellArea("B2", "C3")));

        int si = sheet.getValidations().add(CellArea.createCellArea("F1", "F1"));
        Validation second = sheet.getValidations().get(si);
        second.setType(ValidationType.LIST);
        second.setFormula1("\"Y,N\"");
        assertEquals(2, sheet.getValidations().getCount());

        sheet.getValidations().removeACell(0, 0);
        assertNull(sheet.getValidations().getValidationInCell(0, 0));
        assertNotNull(sheet.getValidations().getValidationInCell(0, 1));

        sheet.getValidations().removeArea(CellArea.createCellArea("F1", "F1"));
        assertEquals(1, sheet.getValidations().getCount());
        assertThrows(CellsException.class, () -> sheet.getValidations().removeACell(-1, 0));
        assertThrows(CellsException.class, () -> sheet.getValidations().get(-1));
    }

    // =========================================================================
    // Conditional formatting
    // =========================================================================

    @Test
    void conditional_formatting_apis_mutate_expected_settings() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        int idx = sheet.getConditionalFormattings().add();
        FormatConditionCollection col = sheet.getConditionalFormattings().get(idx);
        col.addArea(CellArea.createCellArea("A1", "A3"));
        int ci = col.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "=1", "=9");
        FormatCondition cond = col.get(ci);
        cond.setStopIfTrue(true);
        cond.setPriority(1);
        Style style = cond.getStyle();
        style.setPattern(FillPattern.SOLID);
        style.setForegroundColor(Color.fromArgb(255, 255, 0, 0));
        cond.setStyle(style);

        assertEquals(1, sheet.getConditionalFormattings().getCount());
        assertEquals(1, col.getRangeCount());
        assertEquals(1, col.getCount());
        assertEquals("1", cond.getFormula1());
        assertEquals("9", cond.getFormula2());
        assertTrue(cond.getStopIfTrue());
        assertEquals(FillPattern.SOLID, cond.getStyle().getPattern());

        col.addCondition(FormatConditionType.EXPRESSION);
        assertEquals(2, col.getCount());
        col.removeCondition(1);
        assertEquals(1, col.getCount());

        col.addArea(CellArea.createCellArea("C1", "C2"));
        assertEquals(2, col.getRangeCount());
        col.removeArea(0, 0, 1, 1);
        assertEquals(2, col.getRangeCount());
        assertEquals(0, col.getCellArea(0).getFirstRow());
        assertEquals(2, col.getCellArea(0).getFirstColumn());

        sheet.getConditionalFormattings().removeArea(0, 2, 2, 1);
        assertEquals(1, sheet.getConditionalFormattings().getCount());
        sheet.getConditionalFormattings().removeAt(0);
        assertEquals(0, sheet.getConditionalFormattings().getCount());
        assertThrows(CellsException.class, () -> sheet.getConditionalFormattings().get(-1));
    }

    @Test
    void conditional_formatting_advanced_apis_mutate_expected_settings() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        // ContainsText
        FormatConditionCollection cfc = sheet.getConditionalFormattings().get(sheet.getConditionalFormattings().add());
        cfc.addArea(CellArea.createCellArea("A1", "A10"));
        FormatCondition ctRule = cfc.get(cfc.addCondition(FormatConditionType.CONTAINS_TEXT));
        ctRule.setFormula1("error");
        ctRule.setPriority(2);

        // TimePeriod
        FormatConditionCollection cfc2 = sheet.getConditionalFormattings().get(sheet.getConditionalFormattings().add());
        cfc2.addArea(CellArea.createCellArea("B1", "B10"));
        FormatCondition tpRule = cfc2.get(cfc2.addCondition(FormatConditionType.TIME_PERIOD));
        tpRule.setTimePeriod("today");

        // Top10
        FormatConditionCollection cfc3 = sheet.getConditionalFormattings().get(sheet.getConditionalFormattings().add());
        cfc3.addArea(CellArea.createCellArea("C1", "C10"));
        FormatCondition topRule = cfc3.get(cfc3.addCondition(FormatConditionType.TOP_10));
        topRule.setPercent(true);
        topRule.setRank(10);

        // ColorScale
        FormatConditionCollection cfc4 = sheet.getConditionalFormattings().get(sheet.getConditionalFormattings().add());
        cfc4.addArea(CellArea.createCellArea("D1", "D10"));
        FormatCondition colorRule = cfc4.get(cfc4.addCondition(FormatConditionType.COLOR_SCALE));
        colorRule.setColorScaleCount(3);
        colorRule.setMinColor(Color.fromArgb(255, 248, 105, 107));
        colorRule.setMidColor(Color.fromArgb(255, 255, 235, 132));
        colorRule.setMaxColor(Color.fromArgb(255, 99, 190, 123));

        // DataBar
        FormatConditionCollection cfc5 = sheet.getConditionalFormattings().get(sheet.getConditionalFormattings().add());
        cfc5.addArea(CellArea.createCellArea("E1", "E10"));
        FormatCondition dbRule = cfc5.get(cfc5.addCondition(FormatConditionType.DATA_BAR));
        dbRule.setBarColor(Color.fromArgb(255, 99, 142, 198));
        dbRule.setShowBorder(true);
        dbRule.setDirection("left-to-right");

        // IconSet
        FormatConditionCollection cfc6 = sheet.getConditionalFormattings().get(sheet.getConditionalFormattings().add());
        cfc6.addArea(CellArea.createCellArea("F1", "F10"));
        FormatCondition isRule = cfc6.get(cfc6.addCondition(FormatConditionType.ICON_SET));
        isRule.setIconSetType("4Arrows");
        isRule.setReverseIcons(true);
        isRule.setShowIconOnly(true);

        assertEquals(6, sheet.getConditionalFormattings().getCount());
        assertEquals(FormatConditionType.CONTAINS_TEXT, ctRule.getType());
        assertEquals("error", ctRule.getFormula1());
        assertEquals(2, ctRule.getPriority());
        assertEquals("today", tpRule.getTimePeriod());
        assertTrue(topRule.getPercent());
        assertEquals(10, topRule.getRank());
        assertEquals(3, colorRule.getColorScaleCount());
        assertEquals(Color.fromArgb(255, 248, 105, 107), colorRule.getMinColor());
        assertEquals(Color.fromArgb(255, 99, 142, 198), dbRule.getBarColor());
        assertTrue(dbRule.getShowBorder());
        assertEquals("4Arrows", isRule.getIconSetType());
        assertTrue(isRule.getReverseIcons());
        assertTrue(isRule.getShowIconOnly());
    }

    // =========================================================================
    // Page setup
    // =========================================================================

    @Test
    void page_setup_apis_mutate_expected_settings() {
        Workbook wb = new Workbook();
        PageSetup ps = wb.getWorksheets().get(0).getPageSetup();

        ps.setLeftMargin(0.508);
        ps.setRightMargin(0.635);
        ps.setOrientation(PageOrientationType.LANDSCAPE);
        ps.setPaperSize(PaperSizeType.PAPER_A4);
        ps.setFirstPageNumber(2);
        ps.setScale(90);
        ps.setFitToPagesWide(1);
        ps.setFitToPagesTall(3);
        ps.setPrintArea("$A$1:$D$20");
        ps.setPrintTitleRows("$1:$2");
        ps.setPrintTitleColumns("$A:$B");
        ps.setLeftHeader("LH");
        ps.setCenterFooter("CF");
        ps.setPrintGridlines(true);
        ps.setCenterHorizontally(true);
        ps.addHorizontalPageBreak(5);
        ps.addVerticalPageBreak(2);

        assertEquals(0.508, ps.getLeftMargin(), 1e-9);
        assertEquals(0.635, ps.getRightMargin(), 1e-9);
        assertEquals(PageOrientationType.LANDSCAPE, ps.getOrientation());
        assertEquals(PaperSizeType.PAPER_A4, ps.getPaperSize());
        assertEquals(Integer.valueOf(2), ps.getFirstPageNumber());
        assertEquals(Integer.valueOf(90), ps.getScale());
        assertEquals(Integer.valueOf(1), ps.getFitToPagesWide());
        assertEquals(Integer.valueOf(3), ps.getFitToPagesTall());
        assertEquals("$A$1:$D$20", ps.getPrintArea());
        assertEquals("$1:$2", ps.getPrintTitleRows());
        assertEquals("$A:$B", ps.getPrintTitleColumns());
        assertEquals("LH", ps.getLeftHeader());
        assertEquals("CF", ps.getCenterFooter());
        assertTrue(ps.getPrintGridlines());
        assertTrue(ps.getCenterHorizontally());
        assertEquals(1, ps.getHorizontalPageBreaks().size());
        assertEquals(Integer.valueOf(5), ps.getHorizontalPageBreaks().get(0));
        assertEquals(1, ps.getVerticalPageBreaks().size());
        assertEquals(Integer.valueOf(2), ps.getVerticalPageBreaks().get(0));
        assertThrows(CellsException.class, () -> ps.setScale(5));
        assertThrows(CellsException.class, () -> ps.setLeftMargin(-0.1));
    }

    // =========================================================================
    // List objects (tables)
    // =========================================================================

    @Test
    void list_object_apis_mutate_expected_settings() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        sheet.getCells().get("A1").putValue("Name");
        sheet.getCells().get("B1").putValue("Score");
        sheet.getCells().get("A2").putValue("Alice");
        sheet.getCells().get("B2").putValue(95);
        sheet.getCells().get("A3").putValue("Bob");
        sheet.getCells().get("B3").putValue(87);

        int idx = sheet.getListObjects().add("A1", "B3", true);
        assertEquals(0, idx);
        assertEquals(1, sheet.getListObjects().getCount());

        ListObject table = sheet.getListObjects().get(idx);
        assertEquals("Table1", table.getDisplayName());
        assertEquals(0, table.getStartRow());
        assertEquals(0, table.getStartColumn());
        assertEquals(2, table.getEndRow());
        assertEquals(1, table.getEndColumn());
        assertTrue(table.isShowHeaderRow());
        assertFalse(table.isShowTotals());
        assertEquals(2, table.getListColumns().getCount());
        assertEquals("Name",  table.getListColumns().get(0).getName());
        assertEquals("Score", table.getListColumns().get(1).getName());

        table.setTableStyleType(TableStyleType.TABLE_STYLE_MEDIUM_9);
        assertEquals(TableStyleType.TABLE_STYLE_MEDIUM_9, table.getTableStyleType());
        assertEquals("TableStyleMedium9", table.getTableStyleName());

        table.setTableStyleName("MyCustom");
        assertEquals("MyCustom", table.getTableStyleName());
        assertEquals(TableStyleType.CUSTOM, table.getTableStyleType());

        table.setTableStyleType(TableStyleType.NONE);
        assertEquals("", table.getTableStyleName());

        table.setShowTotals(true);
        table.getListColumns().get(1).setTotalsCalculation(TotalsCalculation.SUM);
        assertEquals(TotalsCalculation.SUM, table.getListColumns().get(1).getTotalsCalculation());

        table.getListColumns().get(0).setTotalsRowLabel("Total");
        assertEquals("Total", table.getListColumns().get(0).getTotalsRowLabel());

        table.setDisplayName("SalesTable");
        assertEquals("SalesTable", table.getDisplayName());
        assertEquals("SalesTable", sheet.getListObjects().get("SalesTable").getDisplayName());

        table.setComment("Test comment");
        assertEquals("Test comment", table.getComment());

        table.showAutoFilter();
        table.removeAutoFilter();

        table.resize(0, 0, 3, 1, true);
        assertEquals(3, table.getEndRow());

        assertThrows(CellsException.class, () -> table.setDisplayName(""));
        assertThrows(CellsException.class, () -> sheet.getListObjects().add(0, 0, 2, 1, true));
        assertThrows(CellsException.class, () -> sheet.getListObjects().get(5));
        assertThrows(CellsException.class, () -> sheet.getListObjects().removeAt(5));
        assertThrows(CellsException.class, () -> sheet.getListObjects().get("NoSuchTable"));

        sheet.getCells().get("D1").putValue("Category");
        sheet.getCells().get("E1").putValue("Units");
        int idx2 = sheet.getListObjects().add(0, 3, 2, 4, true);
        ListObject table2 = sheet.getListObjects().get(idx2);
        assertEquals(2, sheet.getListObjects().getCount());
        assertEquals("Table2", table2.getDisplayName());

        table2.convertToRange();
        assertEquals(1, sheet.getListObjects().getCount());

        sheet.getListObjects().removeAt(0);
        assertEquals(0, sheet.getListObjects().getCount());
    }

    // =========================================================================
    // Pictures
    // =========================================================================

    @Test
    void picture_apis_mutate_expected_settings() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        byte[] jpegBytes = PictureScenarioFactory.loadJpeg();
        byte[] pngBytes  = PictureScenarioFactory.loadPng();

        assertEquals(0, sheet.getPictures().getCount());

        int idx = sheet.getPictures().add(1, 2, 5, 6, jpegBytes);
        assertEquals(0, idx);
        assertEquals(1, sheet.getPictures().getCount());

        Picture pic = sheet.getPictures().get(0);
        assertEquals("Picture 1", pic.getName());
        assertEquals(1, pic.getUpperLeftRow());
        assertEquals(2, pic.getUpperLeftColumn());
        assertEquals(5, pic.getLowerRightRow());
        assertEquals(6, pic.getLowerRightColumn());
        assertEquals(ImageType.JPEG, pic.getImageType());
        assertEquals(jpegBytes.length, pic.getData().length);

        pic.setName("Logo");
        assertEquals("Logo", sheet.getPictures().get(0).getName());

        pic.setUpperLeftRow(0);
        pic.setUpperLeftColumn(0);
        pic.setLowerRightRow(3);
        pic.setLowerRightColumn(3);
        assertEquals(0, sheet.getPictures().get(0).getUpperLeftRow());
        assertEquals(3, sheet.getPictures().get(0).getLowerRightRow());

        int idx2 = sheet.getPictures().add(5, 0, 8, 4, pngBytes);
        assertEquals(1, idx2);
        assertEquals(2, sheet.getPictures().getCount());
        assertEquals(ImageType.PNG, sheet.getPictures().get(1).getImageType());

        sheet.getPictures().removeAt(0);
        assertEquals(1, sheet.getPictures().getCount());
        assertEquals(ImageType.PNG, sheet.getPictures().get(0).getImageType());

        assertThrows(CellsException.class, () -> sheet.getPictures().add(-1, 0, 3, 3, jpegBytes));
        assertThrows(CellsException.class, () -> sheet.getPictures().add(0, 0, 3, 3, (byte[])null));
        assertThrows(CellsException.class, () -> sheet.getPictures().add(0, 0, 3, 3, new byte[0]));
        assertThrows(CellsException.class, () -> sheet.getPictures().removeAt(99));
        assertThrows(CellsException.class, () -> sheet.getPictures().get(99));
    }

    // =========================================================================
    // Charts
    // =========================================================================

    @Test
    void chart_collection_is_empty_on_new_worksheet() {
        assertEquals(0, new Workbook().getWorksheets().get(0).getCharts().getCount());
    }

    @Test
    void chart_add_column_chart_returns_index_zero() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Sheet1");
        int idx = sheet.getCharts().add(ChartType.COLUMN, "Sheet1!$B$1:$B$5", 5, 0, 15, 8);
        assertEquals(0, idx);
        assertEquals(1, sheet.getCharts().getCount());
    }

    @Test
    void chart_add_multiple_charts_increments_count() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Sheet1");
        sheet.getCharts().add(ChartType.COLUMN, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
        sheet.getCharts().add(ChartType.LINE,   "Sheet1!$C$1:$C$5", 12, 0, 22, 8);
        assertEquals(2, sheet.getCharts().getCount());
        assertEquals(ChartType.COLUMN, sheet.getCharts().get(0).getChartType());
        assertEquals(ChartType.LINE,   sheet.getCharts().get(1).getChartType());
    }

    @Test
    void chart_add_sets_correct_position() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Sheet1");
        sheet.getCharts().add(ChartType.BAR, "Sheet1!$B$1:$B$5", 7, 2, 20, 12);
        Chart chart = sheet.getCharts().get(0);
        assertEquals(7,  chart.getUpperLeftRow());
        assertEquals(2,  chart.getUpperLeftColumn());
        assertEquals(20, chart.getLowerRightRow());
        assertEquals(12, chart.getLowerRightColumn());
    }

    @Test
    void chart_add_sets_correct_chart_type() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Sheet1");
        sheet.getCharts().add(ChartType.PIE, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
        assertEquals(ChartType.PIE, sheet.getCharts().get(0).getChartType());
    }

    @Test
    void chart_add_unsupported_chartex_type_throws() {
        Worksheet sheet = new Workbook().getWorksheets().get(0);
        assertThrows(UnsupportedFeatureException.class,
            () -> sheet.getCharts().add(ChartType.WATERFALL, "Sheet1!$B$1:$B$5", 0, 0, 10, 8));
        assertThrows(UnsupportedFeatureException.class,
            () -> sheet.getCharts().add(ChartType.TREEMAP, "Sheet1!$B$1:$B$5", 0, 0, 10, 8));
        assertThrows(UnsupportedFeatureException.class,
            () -> sheet.getCharts().add(ChartType.FUNNEL, "Sheet1!$B$1:$B$5", 0, 0, 10, 8));
    }

    @Test
    void chart_index_out_of_range_throws() {
        Worksheet sheet = new Workbook().getWorksheets().get(0);
        assertThrows(CellsException.class, () -> sheet.getCharts().get(0));
        sheet.setName("Sheet1");
        sheet.getCharts().add(ChartType.COLUMN, "Sheet1!$B$1:$B$5", 0, 0, 10, 8);
        assertThrows(CellsException.class, () -> sheet.getCharts().get(1));
        assertThrows(CellsException.class, () -> sheet.getCharts().get(-1));
    }

    @Test void chart_roundtrip_column(@TempDir Path tmp) { roundtripChart(ChartType.COLUMN, tmp); }
    @Test void chart_roundtrip_bar(@TempDir Path tmp)    { roundtripChart(ChartType.BAR, tmp); }
    @Test void chart_roundtrip_line(@TempDir Path tmp)   { roundtripChart(ChartType.LINE, tmp); }
    @Test void chart_roundtrip_area(@TempDir Path tmp)   { roundtripChart(ChartType.AREA, tmp); }
    @Test void chart_roundtrip_pie(@TempDir Path tmp)    { roundtripChart(ChartType.PIE, tmp); }
    @Test void chart_roundtrip_doughnut(@TempDir Path tmp)   { roundtripChart(ChartType.DOUGHNUT, tmp); }
    @Test void chart_roundtrip_scatter(@TempDir Path tmp)    { roundtripChart(ChartType.SCATTER, tmp); }
    @Test void chart_roundtrip_radar(@TempDir Path tmp)  { roundtripChart(ChartType.RADAR, tmp); }
    @Test void chart_roundtrip_bubble(@TempDir Path tmp) { roundtripChart(ChartType.BUBBLE, tmp); }
    @Test void chart_roundtrip_stock(@TempDir Path tmp)  { roundtripChart(ChartType.STOCK, tmp); }
    @Test void chart_roundtrip_column3d(@TempDir Path tmp) { roundtripChart(ChartType.COLUMN_3D, tmp); }
    @Test void chart_roundtrip_bar3d(@TempDir Path tmp)   { roundtripChart(ChartType.BAR_3D, tmp); }
    @Test void chart_roundtrip_line3d(@TempDir Path tmp)  { roundtripChart(ChartType.LINE_3D, tmp); }
    @Test void chart_roundtrip_area3d(@TempDir Path tmp)  { roundtripChart(ChartType.AREA_3D, tmp); }
    @Test void chart_roundtrip_pie3d(@TempDir Path tmp)   { roundtripChart(ChartType.PIE_3D, tmp); }
    @Test void chart_roundtrip_surface3d(@TempDir Path tmp) { roundtripChart(ChartType.SURFACE_3D, tmp); }
    @Test void chart_roundtrip_contour(@TempDir Path tmp)   { roundtripChart(ChartType.CONTOUR, tmp); }

    @Test
    void chart_worksheet_with_picture_and_chart_saves_both(@TempDir Path tmp) throws Exception {
        byte[] jpegBytes = PictureScenarioFactory.loadJpeg();
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Sheet1");
        sheet.getPictures().add(0, 0, 5, 5, jpegBytes);
        sheet.getCharts().add(ChartType.COLUMN, "Sheet1!$B$1:$B$5", 6, 0, 20, 10);
        assertEquals(1, sheet.getPictures().getCount());
        assertEquals(1, sheet.getCharts().getCount());

        Path outPath = tmp.resolve("pic_chart.xlsx");
        wb.save(outPath.toString());

        Workbook reloaded = new Workbook(outPath.toString());
        assertEquals(1, reloaded.getWorksheets().get(0).getPictures().getCount());
        assertEquals(1, reloaded.getWorksheets().get(0).getCharts().getCount());
        assertEquals(ChartType.COLUMN, reloaded.getWorksheets().get(0).getCharts().get(0).getChartType());
    }

    // =========================================================================
    // Comments
    // =========================================================================

    @Test
    void comment_collection_empty_on_new_worksheet() {
        assertEquals(0, new Workbook().getWorksheets().get(0).getComments().getCount());
    }

    @Test
    void comment_add_by_row_col_sets_defaults() {
        Workbook wb = new Workbook();
        Comment c = wb.getWorksheets().get(0).getComments().add(2, 3);
        assertEquals(1, wb.getWorksheets().get(0).getComments().getCount());
        assertEquals(2, c.getRow());
        assertEquals(3, c.getColumn());
        assertEquals("", c.getNote());
        assertEquals("", c.getAuthor());
        assertFalse(c.isVisible());
        assertEquals(129, c.getWidth());
        assertEquals(75, c.getHeight());
    }

    @Test
    void comment_add_by_cell_name_parses_reference() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        Comment c = sheet.getComments().add("B3");
        assertEquals(2, c.getRow());
        assertEquals(1, c.getColumn());
        assertEquals(1, sheet.getComments().getCount());
        Comment byName = sheet.getComments().get("B3");
        assertNotNull(byName);
        assertEquals(2, byName.getRow());
    }

    @Test
    void comment_add_duplicate_cell_throws() {
        Worksheet sheet = new Workbook().getWorksheets().get(0);
        sheet.getComments().add(0, 0);
        assertThrows(CellsException.class, () -> sheet.getComments().add(0, 0));
    }

    @Test
    void comment_add_negative_row_throws() {
        assertThrows(CellsException.class, () -> new Workbook().getWorksheets().get(0).getComments().add(-1, 0));
    }

    @Test
    void comment_add_negative_column_throws() {
        assertThrows(CellsException.class, () -> new Workbook().getWorksheets().get(0).getComments().add(0, -1));
    }

    @Test
    void comment_count_increments_after_add() {
        Worksheet sheet = new Workbook().getWorksheets().get(0);
        sheet.getComments().add(0, 0);
        assertEquals(1, sheet.getComments().getCount());
        sheet.getComments().add(1, 0);
        assertEquals(2, sheet.getComments().getCount());
        sheet.getComments().add(0, 1);
        assertEquals(3, sheet.getComments().getCount());
    }

    @Test
    void comment_remove_at_index_decrements_count() {
        Worksheet sheet = new Workbook().getWorksheets().get(0);
        sheet.getComments().add(0, 0);
        sheet.getComments().add(1, 0);
        assertEquals(2, sheet.getComments().getCount());
        sheet.getComments().removeAt(0);
        assertEquals(1, sheet.getComments().getCount());
        assertEquals(1, sheet.getComments().get(0).getRow());
    }

    @Test
    void comment_remove_at_index_out_of_range_throws() {
        Worksheet sheet = new Workbook().getWorksheets().get(0);
        assertThrows(CellsException.class, () -> sheet.getComments().removeAt(0));
        sheet.getComments().add(0, 0);
        assertThrows(CellsException.class, () -> sheet.getComments().removeAt(1));
        assertThrows(CellsException.class, () -> sheet.getComments().removeAt(-1));
    }

    @Test
    void comment_remove_at_cell_name_noop_when_not_found() {
        Worksheet sheet = new Workbook().getWorksheets().get(0);
        sheet.getComments().add(0, 0);
        sheet.getComments().removeAt("Z99");
        assertEquals(1, sheet.getComments().getCount());
    }

    @Test
    void comment_indexer_by_cell_name_returns_null_when_not_found() {
        Worksheet sheet = new Workbook().getWorksheets().get(0);
        sheet.getComments().add(0, 0);
        assertNotNull(sheet.getComments().get("A1"));
        assertNull(sheet.getComments().get("B2"));
    }

    @Test
    void comment_note_and_author_mutate_via_property() {
        Worksheet sheet = new Workbook().getWorksheets().get(0);
        Comment c = sheet.getComments().add(0, 0);
        c.setNote("Review this");
        c.setAuthor("Alice");
        c.setVisible(true);
        assertEquals("Review this", sheet.getComments().get(0).getNote());
        assertEquals("Alice",       sheet.getComments().get(0).getAuthor());
        assertTrue(sheet.getComments().get(0).isVisible());
    }

    @Test
    void comment_is_visible_defaults_false() {
        assertFalse(new Workbook().getWorksheets().get(0).getComments().add(0, 0).isVisible());
    }

    @Test
    void comment_width_height_defaults_are_129_and_75() {
        Comment c = new Workbook().getWorksheets().get(0).getComments().add(0, 0);
        assertEquals(129, c.getWidth());
        assertEquals(75,  c.getHeight());
    }

    @Test
    void comment_width_zero_throws() {
        Comment c = new Workbook().getWorksheets().get(0).getComments().add(0, 0);
        assertThrows(CellsException.class, () -> c.setWidth(0));
        assertThrows(CellsException.class, () -> c.setWidth(-5));
    }

    @Test
    void comment_height_zero_throws() {
        Comment c = new Workbook().getWorksheets().get(0).getComments().add(0, 0);
        assertThrows(CellsException.class, () -> c.setHeight(0));
        assertThrows(CellsException.class, () -> c.setHeight(-1));
    }

    // =========================================================================
    // Private helpers
    // =========================================================================

    private static void roundtripChart(ChartType type, Path tmp) {
        String name = type.name() + ".xlsx";
        Path outPath = tmp.resolve(name);
        Workbook wb = ChartScenarioFactory.createChartWorkbook(type);
        wb.save(outPath.toString());
        Workbook reloaded = new Workbook(outPath.toString());
        ChartScenarioFactory.assertChartWorkbook(reloaded, type);
    }
}
