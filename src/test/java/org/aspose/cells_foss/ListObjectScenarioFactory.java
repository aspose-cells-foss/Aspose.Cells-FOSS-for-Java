package org.aspose.cells_foss;

import static org.junit.jupiter.api.Assertions.*;

/** Factory for ListObject (table) test workbooks. Mirrors C# ListObjectScenarioFactory. */
public final class ListObjectScenarioFactory {

    private ListObjectScenarioFactory() {}

    public static Workbook createTableWorkbook() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("SalesData");

        sheet.getCells().get("A1").putValue("Product");
        sheet.getCells().get("B1").putValue("Region");
        sheet.getCells().get("C1").putValue("Revenue");
        sheet.getCells().get("A2").putValue("Widget");
        sheet.getCells().get("B2").putValue("North");
        sheet.getCells().get("C2").putValue(1200);
        sheet.getCells().get("A3").putValue("Gadget");
        sheet.getCells().get("B3").putValue("South");
        sheet.getCells().get("C3").putValue(850);
        sheet.getCells().get("A4").putValue("Doohickey");
        sheet.getCells().get("B4").putValue("East");
        sheet.getCells().get("C4").putValue(2100);

        int idx = sheet.getListObjects().add("A1", "C4", true);
        ListObject table = sheet.getListObjects().get(idx);
        table.setTableStyleType(TableStyleType.TABLE_STYLE_MEDIUM_2);
        table.setShowTableStyleRowStripes(true);
        table.setShowTableStyleFirstColumn(false);
        table.setShowTableStyleLastColumn(false);
        table.setShowTableStyleColumnStripes(false);
        table.setShowTotals(true);
        table.getListColumns().get(2).setTotalsCalculation(TotalsCalculation.SUM);
        table.getListColumns().get(0).setTotalsRowLabel("Total");

        return wb;
    }

    public static void assertTableWorkbook(Workbook wb) {
        Worksheet sheet = wb.getWorksheets().get(0);
        assertEquals(1, sheet.getListObjects().getCount());

        ListObject table = sheet.getListObjects().get(0);
        assertEquals("Table1", table.getDisplayName());
        assertEquals(0, table.getStartRow());
        assertEquals(0, table.getStartColumn());
        assertEquals(4, table.getEndRow());
        assertEquals(2, table.getEndColumn());
        assertTrue(table.isShowHeaderRow());
        assertTrue(table.isShowTotals());
        assertEquals(TableStyleType.TABLE_STYLE_MEDIUM_2, table.getTableStyleType());
        assertEquals("TableStyleMedium2", table.getTableStyleName());
        assertTrue(table.isShowTableStyleRowStripes());
        assertFalse(table.isShowTableStyleFirstColumn());
        assertFalse(table.isShowTableStyleLastColumn());
        assertFalse(table.isShowTableStyleColumnStripes());
        assertEquals(3, table.getListColumns().getCount());
        assertEquals("Product", table.getListColumns().get(0).getName());
        assertEquals("Region",  table.getListColumns().get(1).getName());
        assertEquals("Revenue", table.getListColumns().get(2).getName());
        assertEquals(TotalsCalculation.SUM,  table.getListColumns().get(2).getTotalsCalculation());
        assertEquals(TotalsCalculation.NONE, table.getListColumns().get(0).getTotalsCalculation());
        assertEquals("Total",   table.getListColumns().get(0).getTotalsRowLabel());
    }
}

