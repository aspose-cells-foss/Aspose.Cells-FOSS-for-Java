package com.aspose.cells_foss;

import com.aspose.cells_foss.core.AutoFilterModel;
import java.time.LocalDateTime;

/**
 * Factory for auto-filter test workbooks.
 * Mirrors C# AutoFilterScenarioFactory from shared test infrastructure.
 * Uses Apache POI-compatible XLSX output for XML verification.
 */
public final class AutoFilterScenarioFactory {

    /**
     * Verifies that auto filter scenario factory.
     */
    private AutoFilterScenarioFactory() {}

    /**
     * Verifies that create auto filter workbook.
     */
    public static Workbook createAutoFilterWorkbook() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);
        sheet.setName("Filtered");

        sheet.getCells().get(0, 0).putValue("Status");
        sheet.getCells().get(0, 1).putValue("Amount");
        sheet.getCells().get(0, 2).putValue("Color");
        sheet.getCells().get(0, 3).putValue("Date");
        sheet.getCells().get(0, 4).putValue("Score");
        sheet.getCells().get(1, 0).putValue("Open");
        sheet.getCells().get(1, 1).putValue(10);
        sheet.getCells().get(1, 2).putValue("Red");
        sheet.getCells().get(1, 3).putValue(LocalDateTime.of(2024, 5, 1, 0, 0));
        sheet.getCells().get(1, 4).putValue(70);
        sheet.getCells().get(2, 0).putValue("Closed");
        sheet.getCells().get(2, 1).putValue(20);
        sheet.getCells().get(2, 2).putValue("Blue");
        sheet.getCells().get(2, 3).putValue(LocalDateTime.of(2024, 5, 2, 0, 0));
        sheet.getCells().get(2, 4).putValue(80);
        sheet.getCells().get(3, 0).putValue("Open");
        sheet.getCells().get(3, 1).putValue(30);
        sheet.getCells().get(3, 2).putValue("Green");
        sheet.getCells().get(3, 3).putValue(LocalDateTime.of(2024, 5, 3, 0, 0));
        sheet.getCells().get(3, 4).putValue(90);

        sheet.getAutoFilter().setRange("A1:E4");

        AutoFilter.FilterColumnCollection filterColumns = sheet.getAutoFilter().getFilterColumns();

        AutoFilter.FilterColumn statusColumn = filterColumns.get(filterColumns.add(0));
        statusColumn.setHiddenButton(true);
        statusColumn.getFilters().add("Open");
        statusColumn.getFilters().add("Closed");

        AutoFilter.FilterColumn amountColumn = filterColumns.get(filterColumns.add(1));
        amountColumn.getCustomFilters().setMatchAll(true);  // setMatchAll is a setter — OK
        amountColumn.getCustomFilters().add(AutoFilterModel.FilterOperatorType.GREATER_OR_EQUAL, "10");
        amountColumn.getCustomFilters().add(AutoFilterModel.FilterOperatorType.LESS_OR_EQUAL, "30");

        AutoFilter.AutoFilterSortState sortState = sheet.getAutoFilter().getSortState();
        sortState.setRef("A2:E4");
        sortState.setCaseSensitive(true);  // setter is OK

        AutoFilter.AutoFilterSortConditionCollection sortConditions = sortState.getSortConditions();
        AutoFilter.AutoFilterSortCondition valueSort =
                sortConditions.get(sortConditions.add("B2:B4"));
        valueSort.setDescending(true);
        valueSort.setSortBy("value");

        return workbook;
    }

    /**
     * Verifies that assert auto filter.
     * @param workbook workbook to apply
     */
    public static void assertAutoFilter(Workbook workbook) {
        Worksheet sheet = workbook.getWorksheets().get(0);
        AssertEx.assertEqual("A1:E4", sheet.getAutoFilter().getRange());
        AssertEx.assertEqual(2, sheet.getAutoFilter().getFilterColumns().getCount());

        AutoFilter.FilterColumn statusColumn = sheet.getAutoFilter().getFilterColumns().get(0);
        AssertEx.assertEqual(0, statusColumn.getColumnIndex());
        AssertEx.assertTrue(statusColumn.getHiddenButton());
        AssertEx.assertEqual(2, statusColumn.getFilters().getCount());
        AssertEx.assertEqual("Open", statusColumn.getFilters().get(0));
        AssertEx.assertEqual("Closed", statusColumn.getFilters().get(1));

        AutoFilter.FilterColumn amountColumn = sheet.getAutoFilter().getFilterColumns().get(1);
        AssertEx.assertEqual(1, amountColumn.getColumnIndex());
        AssertEx.assertTrue(amountColumn.getCustomFilters().isMatchAll());
        AssertEx.assertEqual(2, amountColumn.getCustomFilters().getCount());
        AssertEx.assertEqual(AutoFilterModel.FilterOperatorType.GREATER_OR_EQUAL,
                amountColumn.getCustomFilters().get(0).getOperator());
        AssertEx.assertEqual("10", amountColumn.getCustomFilters().get(0).getValue());
        AssertEx.assertEqual(AutoFilterModel.FilterOperatorType.LESS_OR_EQUAL,
                amountColumn.getCustomFilters().get(1).getOperator());
        AssertEx.assertEqual("30", amountColumn.getCustomFilters().get(1).getValue());

        AssertEx.assertEqual("A2:E4", sheet.getAutoFilter().getSortState().getRef());
        AssertEx.assertTrue(sheet.getAutoFilter().getSortState().isCaseSensitive());
        AssertEx.assertEqual(1, sheet.getAutoFilter().getSortState().getSortConditions().getCount());

        AutoFilter.AutoFilterSortCondition valueSort =
                sheet.getAutoFilter().getSortState().getSortConditions().get(0);
        AssertEx.assertEqual("B2:B4", valueSort.getRef());
        AssertEx.assertTrue(valueSort.isDescending());
        AssertEx.assertEqual("value", valueSort.getSortBy());
    }
}
