package org.aspose.cells_foss;

import java.time.LocalDateTime;

/**
 * Factory for creating mixed-cell test workbooks.
 * Mirrors C# WorkbookScenarioFactory from shared test infrastructure.
 */
public final class WorkbookScenarioFactory {

    /**
     * Verifies that workbook scenario factory.
     */
    private WorkbookScenarioFactory() {}

    /**
     * Verifies that create mixed cell workbook.
     */
    public static Workbook createMixedCellWorkbook() {
        return createMixedCellWorkbook(false);
    }

    /**
     * Verifies that create mixed cell workbook.
     * @param useDate1904 use date 1904
     */
    public static Workbook createMixedCellWorkbook(boolean useDate1904) {
        Workbook workbook = new Workbook();
        workbook.getSettings().setDate1904(useDate1904);

        Worksheet sheet = workbook.getWorksheets().get(0);
        sheet.setName("Data");
        sheet.getCells().get("A1").putValue("Hello");
        sheet.getCells().get("B1").putValue(123);
        sheet.getCells().get("C1").putValue(true);
        sheet.getCells().get("D1").putValue(12.5d);
        sheet.getCells().get("E1").putValue(6.02214076E+23);
        sheet.getCells().get("F1").putValue(LocalDateTime.of(2024, 5, 6, 7, 8, 9));
        sheet.getCells().get("G1").putValue(20);
        sheet.getCells().get("G1").setFormula("=B1*2");
        return workbook;
    }

    /**
     * Verifies that assert workbook data equals.
     * @param expected expected
     * @param actual actual
     */
    public static void assertWorkbookDataEquals(Workbook expected, Workbook actual) {
        Worksheet expectedSheet = expected.getWorksheets().get(0);
        Worksheet actualSheet = actual.getWorksheets().get(0);

        // Walk the current collection so every entry is processed consistently.
        for (String cellName : new String[]{"A1", "B1", "C1", "D1", "E1", "F1", "G1"}) {
            AssertEx.assertEqual(expectedSheet.getCells().get(cellName).getType(),
                    actualSheet.getCells().get(cellName).getType(),
                    "Type mismatch for " + cellName + ".");
            AssertEx.assertEqual(expectedSheet.getCells().get(cellName).getStringValue(),
                    actualSheet.getCells().get(cellName).getStringValue(),
                    "Value mismatch for " + cellName + ".");
            AssertEx.assertEqual(expectedSheet.getCells().get(cellName).getFormula(),
                    actualSheet.getCells().get(cellName).getFormula(),
                    "Formula mismatch for " + cellName + ".");
        }
    }
}

