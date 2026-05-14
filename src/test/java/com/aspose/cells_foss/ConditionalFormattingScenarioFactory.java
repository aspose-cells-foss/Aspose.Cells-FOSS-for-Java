package com.aspose.cells_foss;

import static org.junit.jupiter.api.Assertions.*;

/** Factory for conditional-formatting test workbooks. Mirrors C# ConditionalFormattingScenarioFactory. */
public final class ConditionalFormattingScenarioFactory {

    private ConditionalFormattingScenarioFactory() {}

    public static Workbook createConditionalFormattingWorkbook() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Formatted");

        // Cell value rule on A1:A10
        int idx = sheet.getConditionalFormattings().add();
        FormatConditionCollection cfc = sheet.getConditionalFormattings().get(idx);
        cfc.addArea(CellArea.createCellArea("A1", "A10"));
        int ci = cfc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "=10", "=90");
        FormatCondition c = cfc.get(ci);
        c.setPriority(1);
        Style style = c.getStyle();
        style.getFont().setBold(true);
        style.setPattern(FillPattern.SOLID);
        style.setForegroundColor(Color.fromArgb(255, 255, 255, 0));
        c.setStyle(style);

        // Expression rule on B1:B10
        int idx2 = sheet.getConditionalFormattings().add();
        FormatConditionCollection cfc2 = sheet.getConditionalFormattings().get(idx2);
        cfc2.addArea(CellArea.createCellArea("B1", "B10"));
        int ci2 = cfc2.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "=B1>0", "");
        FormatCondition c2 = cfc2.get(ci2);
        c2.setPriority(2);

        return wb;
    }

    public static void assertConditionalFormattings(Workbook wb) {
        Worksheet sheet = wb.getWorksheets().get(0);
        assertEquals(2, sheet.getConditionalFormattings().getCount());

        FormatConditionCollection cfc = sheet.getConditionalFormattings().get(0);
        assertEquals(1, cfc.getCount());
        FormatCondition c = cfc.get(0);
        assertEquals(FormatConditionType.CELL_VALUE, c.getType());
        assertEquals(OperatorType.BETWEEN, c.getOperator());
        assertEquals("10", c.getFormula1());
        assertEquals("90", c.getFormula2());
        assertTrue(c.getStyle().getFont().getBold());
    }

    public static Workbook createAdvancedConditionalFormattingWorkbook() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Advanced");

        // ContainsText rule
        int idx = sheet.getConditionalFormattings().add();
        FormatConditionCollection cfc = sheet.getConditionalFormattings().get(idx);
        cfc.addArea(CellArea.createCellArea("A1", "A10"));
        int ci = cfc.addCondition(FormatConditionType.CONTAINS_TEXT);
        FormatCondition c = cfc.get(ci);
        c.setFormula1("error");
        c.setPriority(1);

        // ColorScale rule
        int idx2 = sheet.getConditionalFormattings().add();
        FormatConditionCollection cfc2 = sheet.getConditionalFormattings().get(idx2);
        cfc2.addArea(CellArea.createCellArea("B1", "B10"));
        int ci2 = cfc2.addCondition(FormatConditionType.COLOR_SCALE);
        FormatCondition c2 = cfc2.get(ci2);
        c2.setColorScaleCount(3);
        c2.setMinColor(Color.fromArgb(255, 248, 105, 107));
        c2.setMidColor(Color.fromArgb(255, 255, 235, 132));
        c2.setMaxColor(Color.fromArgb(255, 99, 190, 123));

        // DataBar rule
        int idx3 = sheet.getConditionalFormattings().add();
        FormatConditionCollection cfc3 = sheet.getConditionalFormattings().get(idx3);
        cfc3.addArea(CellArea.createCellArea("C1", "C10"));
        int ci3 = cfc3.addCondition(FormatConditionType.DATA_BAR);
        FormatCondition c3 = cfc3.get(ci3);
        c3.setBarColor(Color.fromArgb(255, 99, 142, 198));
        c3.setShowBorder(true);
        c3.setDirection("left-to-right");

        return wb;
    }

    public static void assertAdvancedConditionalFormattings(Workbook wb) {
        Worksheet sheet = wb.getWorksheets().get(0);
        assertEquals(3, sheet.getConditionalFormattings().getCount());

        FormatCondition ct = sheet.getConditionalFormattings().get(0).get(0);
        assertEquals(FormatConditionType.CONTAINS_TEXT, ct.getType());
        assertEquals("error", ct.getFormula1());

        FormatCondition cs = sheet.getConditionalFormattings().get(1).get(0);
        assertEquals(FormatConditionType.COLOR_SCALE, cs.getType());
        assertEquals(3, cs.getColorScaleCount());

        FormatCondition db = sheet.getConditionalFormattings().get(2).get(0);
        assertEquals(FormatConditionType.DATA_BAR, db.getType());
        assertTrue(db.getShowBorder());
    }
}
