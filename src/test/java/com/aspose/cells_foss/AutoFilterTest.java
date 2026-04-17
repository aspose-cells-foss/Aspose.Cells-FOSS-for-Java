package com.aspose.cells_foss;

import com.aspose.cells_foss.core.AutoFilterModel;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 tests for the AutoFilter API — AF-* test cases.
 */
class AutoFilterTest {

    // =========================================================================
    // AF-01 to AF-04: Range set / clear / null
    // =========================================================================

    /** AF-01: A fresh worksheet has a null/empty AutoFilter range. */
    @Test
    void AF_01_newWorksheetHasEmptyRange() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            String range = ws.getAutoFilter().getRange();
            assertTrue(range == null || range.isEmpty());
        }
    }

    /** AF-02: setRange persists via getRange. */
    @Test
    void AF_02_setRangePersists() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:E1");
            assertEquals("A1:E1", ws.getAutoFilter().getRange());
        }
    }

    /** AF-03: setRange(null) clears the range. */
    @Test
    void AF_03_setRangeNullClearsRange() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().setRange(null);
            String range = ws.getAutoFilter().getRange();
            assertTrue(range == null || range.isEmpty());
        }
    }

    /** AF-04: clear() resets the range to empty and removes all filter columns. */
    @Test
    void AF_04_clearResetsRangeAndColumns() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            ws.getAutoFilter().clear();
            String range = ws.getAutoFilter().getRange();
            assertTrue(range == null || range.isEmpty());
            assertEquals(0, ws.getAutoFilter().getFilterColumns().getCount());
        }
    }

    // =========================================================================
    // AF-10 to AF-17: Filter column add / sorted insertion / duplicate / remove
    // =========================================================================

    /** AF-10: add(0) to an empty collection returns index 0 and count becomes 1. */
    @Test
    void AF_10_addFilterColumnReturnsIndex() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            int idx = ws.getAutoFilter().getFilterColumns().add(0);
            assertEquals(0, idx);
            assertEquals(1, ws.getAutoFilter().getFilterColumns().getCount());
        }
    }

    /** AF-11: Filter columns are inserted in sorted order by column index. */
    @Test
    void AF_11_filterColumnsInsertedInSortedOrder() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(2);
            ws.getAutoFilter().getFilterColumns().add(0);
            ws.getAutoFilter().getFilterColumns().add(1);
            assertEquals(0, ws.getAutoFilter().getFilterColumns().get(0).getColumnIndex());
            assertEquals(1, ws.getAutoFilter().getFilterColumns().get(1).getColumnIndex());
            assertEquals(2, ws.getAutoFilter().getFilterColumns().get(2).getColumnIndex());
        }
    }

    /** AF-12: Adding a duplicate column index throws CellsException. */
    @Test
    void AF_12_duplicateColumnIndexThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(1);
            assertThrows(CellsException.class,
                    () -> ws.getAutoFilter().getFilterColumns().add(1));
        }
    }

    /** AF-13: removeAt(0) removes the first filter column and count drops. */
    @Test
    void AF_13_removeAtFilterColumnDecrementsCount() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            ws.getAutoFilter().getFilterColumns().add(1);
            ws.getAutoFilter().getFilterColumns().removeAt(0);
            assertEquals(1, ws.getAutoFilter().getFilterColumns().getCount());
            assertEquals(1, ws.getAutoFilter().getFilterColumns().get(0).getColumnIndex());
        }
    }

    /** AF-14: FilterColumnCollection.clear() removes all filter columns. */
    @Test
    void AF_14_filterColumnsClearRemovesAll() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            ws.getAutoFilter().getFilterColumns().add(1);
            ws.getAutoFilter().getFilterColumns().clear();
            assertEquals(0, ws.getAutoFilter().getFilterColumns().getCount());
        }
    }

    /** AF-15: get() with out-of-range index throws CellsException. */
    @Test
    void AF_15_filterColumnGetOutOfRangeThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class,
                    () -> ws.getAutoFilter().getFilterColumns().get(0));
        }
    }

    /** AF-16: removeAt() with out-of-range index throws CellsException. */
    @Test
    void AF_16_filterColumnRemoveAtOutOfRangeThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class,
                    () -> ws.getAutoFilter().getFilterColumns().removeAt(0));
        }
    }

    /** AF-17: Adding a negative column index throws CellsException. */
    @Test
    void AF_17_negativeColumnIndexThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            assertThrows(CellsException.class,
                    () -> ws.getAutoFilter().getFilterColumns().add(-1));
        }
    }

    // =========================================================================
    // AF-20 to AF-24: Value filters (FilterValueCollection)
    // =========================================================================

    /** AF-20: add(value) to a fresh filter column returns index 0 and count becomes 1. */
    @Test
    void AF_20_addFilterValueReturnsIndex() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            int idx = ws.getAutoFilter().getFilterColumns().get(0).getFilters().add("Apple");
            assertEquals(0, idx);
            assertEquals(1, ws.getAutoFilter().getFilterColumns().get(0).getFilters().getCount());
        }
    }

    /** AF-21: Multiple values can be added and retrieved by index. */
    @Test
    void AF_21_multipleFilterValuesAddedAndRetrieved() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.FilterValueCollection fv =
                    ws.getAutoFilter().getFilterColumns().get(0).getFilters();
            fv.add("Apple");
            fv.add("Banana");
            fv.add("Cherry");
            assertEquals(3, fv.getCount());
            assertEquals("Apple", fv.get(0));
            assertEquals("Banana", fv.get(1));
            assertEquals("Cherry", fv.get(2));
        }
    }

    /** AF-22: removeAt(0) removes the first filter value and shifts remaining entries. */
    @Test
    void AF_22_filterValueRemoveAtShiftsEntries() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.FilterValueCollection fv =
                    ws.getAutoFilter().getFilterColumns().get(0).getFilters();
            fv.add("Apple");
            fv.add("Banana");
            fv.removeAt(0);
            assertEquals(1, fv.getCount());
            assertEquals("Banana", fv.get(0));
        }
    }

    /** AF-23: clear() on FilterValueCollection removes all values. */
    @Test
    void AF_23_filterValueClearRemovesAll() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.FilterValueCollection fv =
                    ws.getAutoFilter().getFilterColumns().get(0).getFilters();
            fv.add("Apple");
            fv.add("Banana");
            fv.clear();
            assertEquals(0, fv.getCount());
        }
    }

    /** AF-24: Blank filter value is accepted (filters for blank/empty cells). */
    @Test
    void AF_24_blankFilterValueAccepted() throws Exception {
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.FilterValueCollection fv =
                    ws.getAutoFilter().getFilterColumns().get(0).getFilters();
            // Empty string is a valid filter value (matches blank cells in Excel)
            assertDoesNotThrow(() -> fv.add(""));
            assertEquals(1, fv.getCount());
        }
    }

    // =========================================================================
    // AF-30 to AF-37: Custom filters (AutoFilterCustomFilterCollection)
    // =========================================================================

    /** AF-30: A single custom filter can be added. */
    @Test
    void AF_30_addSingleCustomFilter() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterCustomFilterCollection cf =
                    ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
            int idx = cf.add(AutoFilterModel.FilterOperatorType.EQUAL, "100");
            assertEquals(0, idx);
            assertEquals(1, cf.getCount());
            assertEquals(AutoFilterModel.FilterOperatorType.EQUAL, cf.get(0).getOperator());
            assertEquals("100", cf.get(0).getValue());
        }
    }

    /** AF-31: Two custom filters with OR logic (isMatchAll = false). */
    @Test
    void AF_31_twoCustomFiltersWithOrLogic() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterCustomFilterCollection cf =
                    ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
            cf.add(AutoFilterModel.FilterOperatorType.GREATER_THAN, "10");
            cf.add(AutoFilterModel.FilterOperatorType.LESS_THAN, "50");
            assertEquals(2, cf.getCount());
            assertFalse(cf.isMatchAll());
        }
    }

    /** AF-32: Two custom filters with AND logic (isMatchAll = true). */
    @Test
    void AF_32_twoCustomFiltersWithAndLogic() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterCustomFilterCollection cf =
                    ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
            cf.add(AutoFilterModel.FilterOperatorType.GREATER_OR_EQUAL, "10");
            cf.add(AutoFilterModel.FilterOperatorType.LESS_OR_EQUAL, "50");
            cf.setMatchAll(true);
            assertTrue(cf.isMatchAll());
            assertEquals(2, cf.getCount());
        }
    }

    /** AF-33: Adding a third custom filter throws CellsException. */
    @Test
    void AF_33_thirdCustomFilterThrows() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterCustomFilterCollection cf =
                    ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
            cf.add(AutoFilterModel.FilterOperatorType.EQUAL, "A");
            cf.add(AutoFilterModel.FilterOperatorType.EQUAL, "B");
            assertThrows(CellsException.class,
                    () -> cf.add(AutoFilterModel.FilterOperatorType.EQUAL, "C"));
        }
    }

    /** AF-34: setOperator on a custom filter updates the operator. */
    @Test
    void AF_34_setOperatorOnCustomFilter() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterCustomFilterCollection cf =
                    ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
            cf.add(AutoFilterModel.FilterOperatorType.EQUAL, "100");
            cf.get(0).setOperator(AutoFilterModel.FilterOperatorType.NOT_EQUAL);
            assertEquals(AutoFilterModel.FilterOperatorType.NOT_EQUAL, cf.get(0).getOperator());
        }
    }

    /** AF-35: setValue on a custom filter updates the value. */
    @Test
    void AF_35_setValueOnCustomFilter() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterCustomFilterCollection cf =
                    ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
            cf.add(AutoFilterModel.FilterOperatorType.EQUAL, "100");
            cf.get(0).setValue("200");
            assertEquals("200", cf.get(0).getValue());
        }
    }

    /** AF-36: clear() on AutoFilterCustomFilterCollection resets count and matchAll. */
    @Test
    void AF_36_customFilterClearResetsState() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterCustomFilterCollection cf =
                    ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
            cf.add(AutoFilterModel.FilterOperatorType.EQUAL, "A");
            cf.add(AutoFilterModel.FilterOperatorType.EQUAL, "B");
            cf.setMatchAll(true);
            cf.clear();
            assertEquals(0, cf.getCount());
            assertFalse(cf.isMatchAll());
        }
    }

    /** AF-37: Custom filter value is accepted even when blank (blank = match empty cells). */
    @Test
    void AF_37_blankCustomFilterValueAccepted() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterCustomFilterCollection cf =
                    ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
            assertDoesNotThrow(
                    () -> cf.add(AutoFilterModel.FilterOperatorType.EQUAL, ""));
            assertEquals(1, cf.getCount());
        }
    }

    // =========================================================================
    // AF-40 to AF-42: Dynamic filter
    // =========================================================================

    /** AF-40: DynamicFilter is disabled by default. */
    @Test
    void AF_40_dynamicFilterDefaultDisabled() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            assertFalse(ws.getAutoFilter().getFilterColumns().get(0).getDynamicFilter().isEnabled());
        }
    }

    /** AF-41: setType() on a dynamic filter enables it and retains the type string. */
    @Test
    void AF_41_setTypeEnablesDynamicFilter() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterDynamicFilter df =
                    ws.getAutoFilter().getFilterColumns().get(0).getDynamicFilter();
            df.setType("aboveAverage");
            assertTrue(df.isEnabled());
            assertEquals("aboveAverage", df.getType());
        }
    }

    /** AF-42: setEnabled(false) disables the dynamic filter (clear). */
    @Test
    void AF_42_setEnabledFalseDisablesDynamicFilter() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterDynamicFilter df =
                    ws.getAutoFilter().getFilterColumns().get(0).getDynamicFilter();
            df.setType("aboveAverage");
            assertTrue(df.isEnabled());
            df.setEnabled(false);
            assertFalse(df.isEnabled());
        }
    }

    // =========================================================================
    // AF-50 to AF-54: Top-10 filter
    // =========================================================================

    /** AF-50: Top10 filter is disabled by default. */
    @Test
    void AF_50_top10DefaultDisabled() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            assertFalse(ws.getAutoFilter().getFilterColumns().get(0).getTop10().isEnabled());
        }
    }

    /** AF-51: setTop(true) enables the Top10 filter and sets isTop. */
    @Test
    void AF_51_setTopEnablesTop10Filter() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterTop10 top10 =
                    ws.getAutoFilter().getFilterColumns().get(0).getTop10();
            top10.setTop(true);
            assertTrue(top10.isEnabled());
            assertTrue(top10.isTop());
        }
    }

    /** AF-52: setPercent(true) sets the percent flag and keeps enabled. */
    @Test
    void AF_52_setPercentSetsPercentFlag() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterTop10 top10 =
                    ws.getAutoFilter().getFilterColumns().get(0).getTop10();
            top10.setPercent(true);
            assertTrue(top10.isEnabled());
            assertTrue(top10.isPercent());
        }
    }

    /** AF-53: setValue(5.0) persists via getValue() and enables the filter. */
    @Test
    void AF_53_setValuePersistsAndEnables() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterTop10 top10 =
                    ws.getAutoFilter().getFilterColumns().get(0).getTop10();
            top10.setValue(5.0);
            assertTrue(top10.isEnabled());
            assertEquals(5.0, top10.getValue(), 1e-9);
        }
    }

    /** AF-54: clear() on Top10 disables the filter. */
    @Test
    void AF_54_top10ClearDisablesFilter() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (Workbook wb = new Workbook()) {
            Worksheet ws = wb.getWorksheets().get(0);
            ws.getAutoFilter().setRange("A1:D1");
            ws.getAutoFilter().getFilterColumns().add(0);
            AutoFilter.AutoFilterTop10 top10 =
                    ws.getAutoFilter().getFilterColumns().get(0).getTop10();
            top10.setValue(5.0);
            top10.clear();
            assertFalse(top10.isEnabled());
        }
    }

    // =========================================================================
    // AF-60 to AF-65: XLSX round-trip tests (in-memory + reload)
    // =========================================================================

    /** AF-60: AutoFilter range survives save/reload round-trip. */
    @Test
    void AF_60_rangeRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af60.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getAutoFilter().setRange("A1:E1");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals("A1:E1", loaded.getWorksheets().get(0).getAutoFilter().getRange());
            }
        }
    }

    /** AF-61: Filter column index survives save/reload round-trip. */
    @Test
    void AF_61_filterColumnIndexRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af61.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(2);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                assertEquals(1, loaded.getWorksheets().get(0).getAutoFilter().getFilterColumns().getCount());
                assertEquals(2, loaded.getWorksheets().get(0).getAutoFilter().getFilterColumns().get(0).getColumnIndex());
            }
        }
    }

    /** AF-62: Filter values survive save/reload round-trip. */
    @Test
    void AF_62_filterValuesRoundtripViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af62.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(0);
                ws.getAutoFilter().getFilterColumns().get(0).getFilters().add("Alpha");
                ws.getAutoFilter().getFilterColumns().get(0).getFilters().add("Beta");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                AutoFilter.FilterValueCollection fv =
                        loaded.getWorksheets().get(0).getAutoFilter().getFilterColumns().get(0).getFilters();
                assertEquals(2, fv.getCount());
                assertEquals("Alpha", fv.get(0));
                assertEquals("Beta", fv.get(1));
            }
        }
    }

    /** AF-63: Custom filter (single, EQUAL) survives save/reload round-trip. */
    @Test
    void AF_63_customFilterSingleRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af63.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(1);
                ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters()
                        .add(AutoFilterModel.FilterOperatorType.EQUAL, "42");
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                AutoFilter.AutoFilterCustomFilterCollection cf =
                        loaded.getWorksheets().get(0).getAutoFilter().getFilterColumns().get(0).getCustomFilters();
                assertEquals(1, cf.getCount());
                assertEquals(AutoFilterModel.FilterOperatorType.EQUAL, cf.get(0).getOperator());
                assertEquals("42", cf.get(0).getValue());
            }
        }
    }

    /** AF-64: Two custom filters with AND logic survive save/reload round-trip. */
    @Test
    void AF_64_customFilterAndLogicRoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af64.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(0);
                AutoFilter.AutoFilterCustomFilterCollection cf =
                        ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
                cf.add(AutoFilterModel.FilterOperatorType.GREATER_THAN, "10");
                cf.add(AutoFilterModel.FilterOperatorType.LESS_THAN, "100");
                cf.setMatchAll(true);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                AutoFilter.AutoFilterCustomFilterCollection cf =
                        loaded.getWorksheets().get(0).getAutoFilter().getFilterColumns().get(0).getCustomFilters();
                assertEquals(2, cf.getCount());
                assertTrue(cf.isMatchAll());
            }
        }
    }

    /** AF-65: Top10 filter settings survive save/reload round-trip. */
    @Test
    void AF_65_top10RoundtripsViaXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af65.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(0);
                AutoFilter.AutoFilterTop10 top10 =
                        ws.getAutoFilter().getFilterColumns().get(0).getTop10();
                top10.setTop(true);
                top10.setValue(5.0);
                wb.save(path);
            }
            try (Workbook loaded = new Workbook(path)) {
                AutoFilter.AutoFilterTop10 top10 =
                        loaded.getWorksheets().get(0).getAutoFilter().getFilterColumns().get(0).getTop10();
                assertTrue(top10.isEnabled());
                assertTrue(top10.isTop());
                assertEquals(5.0, top10.getValue(), 1e-9);
            }
        }
    }

    // =========================================================================
    // AF-70 to AF-76: Integration — raw XML inspection via ZipPackageHelper
    // =========================================================================

    /** AF-70: Saved XLSX contains &lt;autoFilter&gt; element with correct ref attribute. */
    @Test
    void AF_70_xlsxContainsAutoFilterElement() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af70.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getAutoFilter().setRange("A1:D1");
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(xml.contains("<autoFilter"), "Expected <autoFilter> element in sheet XML");
            assertTrue(xml.contains("A1:D1"), "Expected ref attribute value A1:D1 in autoFilter");
        }
    }

    /** AF-71: Saved XLSX contains &lt;filterColumn&gt; element for added column index. */
    @Test
    void AF_71_xlsxContainsFilterColumnElement() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af71.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(2);
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(xml.contains("<filterColumn"), "Expected <filterColumn> element in sheet XML");
            assertTrue(xml.contains("colId=\"2\""), "Expected colId=\"2\" in filterColumn");
        }
    }

    /** AF-72: Saved XLSX contains &lt;filter val="..."&gt; elements for value filters. */
    @Test
    void AF_72_xlsxContainsFilterValueElements() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af72.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(0);
                ws.getAutoFilter().getFilterColumns().get(0).getFilters().add("Apple");
                ws.getAutoFilter().getFilterColumns().get(0).getFilters().add("Mango");
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(xml.contains("<filters"), "Expected <filters> element in sheet XML");
            assertTrue(xml.contains("val=\"Apple\""), "Expected val=\"Apple\" in filter element");
            assertTrue(xml.contains("val=\"Mango\""), "Expected val=\"Mango\" in filter element");
        }
    }

    /** AF-73: Saved XLSX contains &lt;customFilters&gt; element for custom filters. */
    @Test
    void AF_73_xlsxContainsCustomFiltersElement() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af73.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(0);
                ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters()
                        .add(AutoFilterModel.FilterOperatorType.GREATER_THAN, "50");
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(xml.contains("<customFilters"), "Expected <customFilters> element in sheet XML");
            assertTrue(xml.contains("<customFilter"), "Expected <customFilter> element in sheet XML");
            assertTrue(xml.contains("val=\"50\""), "Expected val=\"50\" in customFilter element");
        }
    }

    /** AF-74: Saved XLSX marks two AND-combined custom filters with and="1". */
    @Test
    void AF_74_xlsxAndCombinedCustomFiltersHaveAndAttribute() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af74.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(0);
                AutoFilter.AutoFilterCustomFilterCollection cf =
                        ws.getAutoFilter().getFilterColumns().get(0).getCustomFilters();
                cf.add(AutoFilterModel.FilterOperatorType.GREATER_THAN, "10");
                cf.add(AutoFilterModel.FilterOperatorType.LESS_THAN, "90");
                cf.setMatchAll(true);
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(xml.contains("and=\"1\""),
                    "Expected and=\"1\" attribute on <customFilters> for AND logic");
        }
    }

    /** AF-75: Saved XLSX contains &lt;top10&gt; element with correct val and top attributes. */
    @Test
    void AF_75_xlsxContainsTop10Element() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af75.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(0);
                AutoFilter.AutoFilterTop10 top10 =
                        ws.getAutoFilter().getFilterColumns().get(0).getTop10();
                top10.setTop(true);
                top10.setValue(5.0);
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(xml.contains("<top10"), "Expected <top10> element in sheet XML");
            assertTrue(xml.contains("top=\"1\""), "Expected top=\"1\" in top10 element");
            assertTrue(xml.contains("val=\"5.0\""), "Expected val=\"5.0\" in top10 element");
        }
    }

    /** AF-76: Saved XLSX contains hiddenButton="1" for a filter column with hidden button. */
    @Test
    void AF_76_xlsxContainsHiddenButtonAttribute() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("AutoFilterTest")) {
            String path = tempDir.getPath("af76.xlsx");
            try (Workbook wb = new Workbook()) {
                Worksheet ws = wb.getWorksheets().get(0);
                ws.getAutoFilter().setRange("A1:D1");
                ws.getAutoFilter().getFilterColumns().add(0);
                ws.getAutoFilter().getFilterColumns().get(0).setHiddenButton(true);
                wb.save(path);
            }
            String xml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
            assertTrue(xml.contains("hiddenButton=\"1\""),
                    "Expected hiddenButton=\"1\" in filterColumn element");
        }
    }
}
