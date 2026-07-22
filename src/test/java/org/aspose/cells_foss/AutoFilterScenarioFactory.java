package org.aspose.cells_foss;

import org.aspose.cells_foss.core.AutoFilterModel;
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
        sheet.getCells().get(4, 0).putValue("Pending");
        sheet.getCells().get(4, 1).putValue(40);
        sheet.getCells().get(4, 2).putValue("Yellow");
        sheet.getCells().get(4, 3).putValue(LocalDateTime.of(2024, 5, 4, 0, 0));
        sheet.getCells().get(4, 4).putValue(60);
        sheet.getCells().get(5, 0).putValue("Closed");
        sheet.getCells().get(5, 1).putValue(50);
        sheet.getCells().get(5, 2).putValue("Black");
        sheet.getCells().get(5, 3).putValue(LocalDateTime.of(2024, 5, 5, 0, 0));
        sheet.getCells().get(5, 4).putValue(50);

        sheet.getAutoFilter().setRange("A1:E6");

        AutoFilter.FilterColumnCollection filterColumns = sheet.getAutoFilter().getFilterColumns();

        AutoFilter.FilterColumn statusColumn = filterColumns.get(filterColumns.add(0));
        statusColumn.setDropdownVisible(false);
        statusColumn.getFilters().add("Open");
        statusColumn.getFilters().add("Closed");

        AutoFilter.FilterColumn amountColumn = filterColumns.get(filterColumns.add(1));
        amountColumn.getCustomFilters().setMatchAll(true);
        amountColumn.getCustomFilters().add(AutoFilterModel.FilterOperatorType.GREATER_OR_EQUAL, "10");
        amountColumn.getCustomFilters().add(AutoFilterModel.FilterOperatorType.LESS_OR_EQUAL, "50");

        // Color, Date, Score columns 鈥?added as plain filter slots
        filterColumns.add(2);
        filterColumns.add(3);
        filterColumns.add(4);

        AutoFilter.AutoFilterSortState sortState = sheet.getAutoFilter().getSortState();
        sortState.setRef("A2:E6");
        sortState.setCaseSensitive(true);

        AutoFilter.AutoFilterSortConditionCollection sortConditions = sortState.getSortConditions();
        AutoFilter.AutoFilterSortCondition valueSort =
                sortConditions.get(sortConditions.add("B2:B6"));
        valueSort.setDescending(true);
        valueSort.setSortBy("value");

        AutoFilter.AutoFilterSortCondition colorSort =
                sortConditions.get(sortConditions.add("C2:C6"));
        colorSort.setSortBy("cellColor");

        AutoFilter.AutoFilterSortCondition iconSort =
                sortConditions.get(sortConditions.add("E2:E6"));
        iconSort.setSortBy("icon");

        return workbook;
    }

    /**
     * Verifies that assert auto filter.
     * @param workbook workbook to apply
     */
    public static void assertAutoFilter(Workbook workbook) {
        Worksheet sheet = workbook.getWorksheets().get(0);
        AssertEx.assertEqual("A1:E6", sheet.getAutoFilter().getRange());
        AssertEx.assertEqual(5, sheet.getAutoFilter().getFilterColumns().getCount());

        AutoFilter.FilterColumn statusColumn = sheet.getAutoFilter().getFilterColumns().get(0);
        AssertEx.assertEqual(0, statusColumn.getFieldIndex());
        AssertEx.assertFalse(statusColumn.isDropdownVisible());
        AssertEx.assertEqual(2, statusColumn.getFilters().getCount());
        AssertEx.assertEqual("Open", statusColumn.getFilters().get(0));
        AssertEx.assertEqual("Closed", statusColumn.getFilters().get(1));

        AutoFilter.FilterColumn amountColumn = sheet.getAutoFilter().getFilterColumns().get(1);
        AssertEx.assertEqual(1, amountColumn.getFieldIndex());
        AssertEx.assertTrue(amountColumn.getCustomFilters().isMatchAll());
        AssertEx.assertEqual(2, amountColumn.getCustomFilters().getCount());
        AssertEx.assertEqual(AutoFilterModel.FilterOperatorType.GREATER_OR_EQUAL,
                amountColumn.getCustomFilters().get(0).getOperator());
        AssertEx.assertEqual("10", amountColumn.getCustomFilters().get(0).getValue());
        AssertEx.assertEqual(AutoFilterModel.FilterOperatorType.LESS_OR_EQUAL,
                amountColumn.getCustomFilters().get(1).getOperator());
        AssertEx.assertEqual("50", amountColumn.getCustomFilters().get(1).getValue());

        AssertEx.assertEqual(2, sheet.getAutoFilter().getFilterColumns().get(2).getFieldIndex());
        AssertEx.assertEqual(3, sheet.getAutoFilter().getFilterColumns().get(3).getFieldIndex());
        AssertEx.assertEqual(4, sheet.getAutoFilter().getFilterColumns().get(4).getFieldIndex());

        AssertEx.assertEqual("A2:E6", sheet.getAutoFilter().getSortState().getRef());
        AssertEx.assertTrue(sheet.getAutoFilter().getSortState().isCaseSensitive());
        AssertEx.assertEqual(3, sheet.getAutoFilter().getSortState().getSortConditions().getCount());

        AutoFilter.AutoFilterSortCondition valueSort =
                sheet.getAutoFilter().getSortState().getSortConditions().get(0);
        AssertEx.assertEqual("B2:B6", valueSort.getRef());
        AssertEx.assertTrue(valueSort.isDescending());
        AssertEx.assertEqual("value", valueSort.getSortBy());

        AssertEx.assertEqual("C2:C6", sheet.getAutoFilter().getSortState().getSortConditions().get(1).getRef());
        AssertEx.assertEqual("E2:E6", sheet.getAutoFilter().getSortState().getSortConditions().get(2).getRef());
    }
}

