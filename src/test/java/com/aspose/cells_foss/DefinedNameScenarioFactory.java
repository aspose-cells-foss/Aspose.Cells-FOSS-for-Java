package com.aspose.cells_foss;

import static org.junit.jupiter.api.Assertions.*;

/** Factory for defined-name test workbooks. Mirrors C# DefinedNameScenarioFactory. */
public final class DefinedNameScenarioFactory {

    private DefinedNameScenarioFactory() {}

    public static Workbook createDefinedNamesWorkbook() {
        Workbook wb = new Workbook();
        wb.getWorksheets().get(0).setName("Data");
        wb.getWorksheets().add("Scoped");

        // Workbook-scoped name
        int i1 = wb.getDefinedNames().add("Total", "=SUM(Data!$A$1:$A$10)");
        DefinedName total = wb.getDefinedNames().get(i1);
        total.setHidden(true);
        total.setComment("Workbook total");

        // Sheet-scoped name (scope = sheet index 1)
        int i2 = wb.getDefinedNames().add("Input", "'Scoped'!$B$2", 1);
        DefinedName input = wb.getDefinedNames().get(i2);
        input.setComment("Local input");

        return wb;
    }

    public static void assertDefinedNames(Workbook wb) {
        assertEquals(2, wb.getDefinedNames().getCount());

        DefinedName total = wb.getDefinedNames().get(0);
        assertEquals("Total", total.getName());
        assertTrue(total.getFormula().contains("SUM"));
        assertNull(total.getLocalSheetIndex());
        assertTrue(total.isHidden());
        assertEquals("Workbook total", total.getComment());

        DefinedName input = wb.getDefinedNames().get(1);
        assertEquals("Input", input.getName());
        assertEquals(Integer.valueOf(1), input.getLocalSheetIndex());
        assertEquals("Local input", input.getComment());
    }
}
