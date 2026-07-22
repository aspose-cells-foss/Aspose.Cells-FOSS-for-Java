package org.aspose.cells_foss;

import static org.junit.jupiter.api.Assertions.*;

/** Factory for chart test workbooks. Mirrors C# ChartScenarioFactory. */
public final class ChartScenarioFactory {

    private static final String[] QUARTER_LABELS = {"Q1", "Q2", "Q3", "Q4", "Q5"};
    private static final int[]    QUARTER_VALUES = {10, 30, 20, 50, 40};

    private ChartScenarioFactory() {}

    public static Workbook createChartWorkbook(ChartType type) {
        return createChartWorkbook(type, "Sheet1");
    }

    public static Workbook createChartWorkbook(ChartType type, String sheetName) {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName(sheetName);

        for (int i = 0; i < 5; i++) {
            sheet.getCells().get(i, 0).putValue(QUARTER_LABELS[i]);
            sheet.getCells().get(i, 1).putValue(QUARTER_VALUES[i]);
        }

        String dataRange = sheetName + "!$B$1:$B$5";
        sheet.getCharts().add(type, dataRange, 7, 0, 22, 10);
        return wb;
    }

    public static void assertChartWorkbook(Workbook wb, ChartType expectedType) {
        Worksheet sheet = wb.getWorksheets().get(0);
        assertEquals(1, sheet.getCharts().getCount(), "Expected exactly 1 chart");
        Chart chart = sheet.getCharts().get(0);
        assertEquals(expectedType, chart.getChartType(), "Wrong chart type after round-trip");
        assertFalse(chart.getName().isBlank(), "Chart name should not be empty");
        assertEquals(7,  chart.getUpperLeftRow(),    "Wrong UpperLeftRow");
        assertEquals(0,  chart.getUpperLeftColumn(), "Wrong UpperLeftColumn");
        assertEquals(22, chart.getLowerRightRow(),   "Wrong LowerRightRow");
        assertEquals(10, chart.getLowerRightColumn(),"Wrong LowerRightColumn");
    }
}

