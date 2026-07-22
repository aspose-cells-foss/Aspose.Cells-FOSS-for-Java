package org.aspose.cells_foss;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 tests for the Conditional Formatting API 鈥?CF-* test cases.
 *
 * CF-01 to CF-07  : ConditionalFormattingCollection management 鈥?all pass now.
 * CF-10 to CF-1C  : FormatConditionCollection area/condition management 鈥?all pass now.
 * CF-20 to CF-30  : addCondition for each FormatConditionType 鈥?structural counts only.
 * CF-40 to CF-43  : Priority tests 鈥?@Disabled (no public read API yet).
 * CF-50 to CF-57  : FormatCondition property getters 鈥?@Disabled (UnsupportedOperationException).
 * CF-60 to CF-66  : XLSX serialisation 鈥?@Disabled.
 * CF-70 to CF-75  : XLSX round-trip / POI integration 鈥?@Disabled.
 */
class ConditionalFormattingTest {

    // =========================================================================
    // Helpers
    // =========================================================================

    /** Returns a fresh worksheet backed by a new workbook. */
    private static Worksheet freshSheet() {
        // Workbook is intentionally not closed: it is a lightweight in-memory
        // object and the returned Worksheet holds a live reference to it.
        @SuppressWarnings("resource")
        Workbook wb = new Workbook();
        return wb.getWorksheets().get(0);
    }

    /** Shorthand: worksheet 鈫?its ConditionalFormattingCollection. */
    private static ConditionalFormattingCollection cfc(Worksheet ws) {
        return ws.getConditionalFormattings();
    }

    // =========================================================================
    // CF-01  New worksheet has zero conditional formatting collections
    // =========================================================================

    /**
     * Verifies that new worksheet has zero collections.
     */
    @Test
    void CF_01_newWorksheetHasZeroCollections() {
        assertEquals(0, cfc(freshSheet()).getCount());
    }

    // =========================================================================
    // CF-02  add() returns 0 for the first collection
    // =========================================================================

    /**
     * Verifies that add returns zero for first collection.
     */
    @Test
    void CF_02_addReturnsZeroForFirstCollection() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        int idx = cfc.add();
        assertEquals(0, idx);
    }

    // =========================================================================
    // CF-03  add() increments getCount()
    // =========================================================================

    /**
     * Verifies that add increments count.
     */
    @Test
    void CF_03_addIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        cfc.add();
        assertEquals(1, cfc.getCount());
        cfc.add();
        assertEquals(2, cfc.getCount());
    }

    // =========================================================================
    // CF-04  get() returns a non-null FormatConditionCollection for a valid index
    // =========================================================================

    /**
     * Verifies that get returns non null for valid index.
     */
    @Test
    void CF_04_getReturnsNonNullForValidIndex() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        int idx = cfc.add();
        assertNotNull(cfc.get(idx));
    }

    // =========================================================================
    // CF-05  get() throws CellsException for an out-of-range index
    // =========================================================================

    /**
     * Verifies that get throws for out of range index.
     */
    @Test
    void CF_05_getThrowsForOutOfRangeIndex() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        assertThrows(CellsException.class, () -> cfc.get(0));
    }

    // =========================================================================
    // CF-06  removeAt() decrements count
    // =========================================================================

    /**
     * Verifies that remove at decrements count.
     */
    @Test
    void CF_06_removeAtDecrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        cfc.add();
        cfc.add();
        assertEquals(2, cfc.getCount());
        cfc.removeAt(0);
        assertEquals(1, cfc.getCount());
    }

    // =========================================================================
    // CF-07  removeAt() throws CellsException for an out-of-range index
    // =========================================================================

    /**
     * Verifies that remove at throws for out of range index.
     */
    @Test
    void CF_07_removeAtThrowsForOutOfRangeIndex() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        assertThrows(CellsException.class, () -> cfc.removeAt(0));
    }

    // =========================================================================
    // CF-10  addArea() increments getRangeCount()
    // =========================================================================

    /**
     * Verifies that add area increments range count.
     */
    @Test
    void CF_10_addAreaIncrementsRangeCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        assertEquals(0, fcc.getRangeCount());
        fcc.addArea(new CellArea(0, 0, 1, 1));
        assertEquals(1, fcc.getRangeCount());
        fcc.addArea(new CellArea(5, 0, 2, 3));
        assertEquals(2, fcc.getRangeCount());
    }

    // =========================================================================
    // CF-11  getCellArea() returns the area that was added
    // =========================================================================

    /**
     * Verifies that get cell area returns added area.
     */
    @Test
    void CF_11_getCellAreaReturnsAddedArea() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        CellArea area = new CellArea(2, 3, 4, 5);
        fcc.addArea(area);
        CellArea retrieved = fcc.getCellArea(0);
        assertEquals(area.getFirstRow(),     retrieved.getFirstRow());
        assertEquals(area.getFirstColumn(),  retrieved.getFirstColumn());
        assertEquals(area.getTotalRows(),    retrieved.getTotalRows());
        assertEquals(area.getTotalColumns(), retrieved.getTotalColumns());
    }

    // =========================================================================
    // CF-12  getCellArea() throws CellsException for out-of-range index
    // =========================================================================

    /**
     * Verifies that get cell area throws for out of range index.
     */
    @Test
    void CF_12_getCellAreaThrowsForOutOfRangeIndex() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        assertThrows(CellsException.class, () -> fcc.getCellArea(0));
    }

    // =========================================================================
    // CF-13  removeArea(int) removes area by index and decrements getRangeCount()
    // =========================================================================

    /**
     * Verifies that remove area by index decrements range count.
     */
    @Test
    void CF_13_removeAreaByIndexDecrementsRangeCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        fcc.addArea(new CellArea(0, 0, 2, 2));
        fcc.addArea(new CellArea(5, 0, 1, 1));
        assertEquals(2, fcc.getRangeCount());
        fcc.removeArea(0);
        assertEquals(1, fcc.getRangeCount());
    }

    // =========================================================================
    // CF-14  addCondition() increments getCount()
    // =========================================================================

    /**
     * Verifies that add condition increments count.
     */
    @Test
    void CF_14_addConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        assertEquals(0, fcc.getCount());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        assertEquals(1, fcc.getCount());
        fcc.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "A1>0", "");
        assertEquals(2, fcc.getCount());
    }

    // =========================================================================
    // CF-15  removeArea(CellArea) subtracts overlapping region 鈥?3 pieces remain
    //
    // Source area  : CellArea(0,0,4,4) = rows 0-3, cols 0-3
    // Removal area : CellArea(1,0,1,2) = row 1, cols 0-1
    //
    // After subtraction the source rectangle is split into three non-overlapping
    // pieces:
    //   Top strip    : row 0,         cols 0-3  (above the removal)
    //   Bottom strip : rows 2-3,      cols 0-3  (below the removal)
    //   Right slice  : row 1,         cols 2-3  (right of removal in the same row)
    // =========================================================================

    /**
     * Verifies that remove area subtracts overlapping region.
     */
    @Test
    void CF_15_removeAreaSubtractsOverlappingRegion() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        fcc.addArea(new CellArea(0, 0, 4, 4));  // rows 0-3, cols 0-3
        assertEquals(1, fcc.getRangeCount());

        fcc.removeArea(new CellArea(1, 0, 1, 2)); // row 1, cols 0-1
        assertEquals(3, fcc.getRangeCount());
    }

    // =========================================================================
    // CF-16  add(CellArea, type, op, f1, f2) adds area AND condition atomically
    // =========================================================================

    /**
     * Verifies that add cell area type op formulas adds area and condition.
     */
    @Test
    void CF_16_addCellAreaTypeOpFormulasAddsAreaAndCondition() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        int condIdx = fcc.add(
                new CellArea(0, 0, 3, 3),
                FormatConditionType.CELL_VALUE, OperatorType.GREATER_THAN, "5", "");
        assertEquals(0, condIdx);
        assertEquals(1, fcc.getRangeCount());
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-17  removeCondition(int) decrements getCount()
    // =========================================================================

    /**
     * Verifies that remove condition decrements count.
     */
    @Test
    void CF_17_removeConditionDecrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addArea(new CellArea(0, 0, 1, 1));
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        fcc.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "A1>0", "");
        assertEquals(2, fcc.getCount());
        fcc.removeCondition(0);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-18  removeCondition() throws CellsException for out-of-range index
    // =========================================================================

    /**
     * Verifies that remove condition throws for out of range index.
     */
    @Test
    void CF_18_removeConditionThrowsForOutOfRangeIndex() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        assertThrows(CellsException.class, () -> fcc.removeCondition(0));
    }

    // =========================================================================
    // CF-19  removeArea(CellArea) on a fully-covered area leaves 0 ranges
    // =========================================================================

    /**
     * Verifies that remove area full coverage leave zero ranges.
     */
    @Test
    void CF_19_removeAreaFullCoverageLeaveZeroRanges() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        // Keep a condition so the collection isn't auto-removed by the area-drop
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        fcc.addArea(new CellArea(0, 0, 2, 2));
        fcc.removeArea(new CellArea(0, 0, 2, 2));
        assertEquals(0, fcc.getRangeCount());
    }

    // =========================================================================
    // CF-1A  removeArea(int) throws CellsException for out-of-range index
    // =========================================================================

    /**
     * Verifies that remove area by index throws for out of range index.
     */
    @Test
    void CF_1A_removeAreaByIndexThrowsForOutOfRangeIndex() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        assertThrows(CellsException.class, () -> fcc.removeArea(0));
    }

    // =========================================================================
    // CF-1B  Removing all areas auto-removes the collection from its parent
    // =========================================================================

    /**
     * Verifies that removing all areas auto removes collection.
     */
    @Test
    void CF_1B_removingAllAreasAutoRemovesCollection() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        int idx = cfc.add();
        assertEquals(1, cfc.getCount());

        FormatConditionCollection fcc = cfc.get(idx);
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        assertEquals(1, fcc.getCount());

        fcc.addArea(new CellArea(0, 0, 1, 1));
        assertEquals(1, fcc.getRangeCount());

        // Removing the only area triggers auto-remove of the collection
        fcc.removeArea(0);
        assertEquals(0, fcc.getRangeCount());
        assertEquals(0, cfc.getCount());
    }

    // =========================================================================
    // CF-1C  Removing all conditions auto-removes the collection from its parent
    // =========================================================================

    /**
     * Verifies that removing all conditions auto removes collection.
     */
    @Test
    void CF_1C_removingAllConditionsAutoRemovesCollection() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        int idx = cfc.add();

        FormatConditionCollection fcc = cfc.get(idx);
        fcc.addArea(new CellArea(0, 0, 1, 1));
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        assertEquals(1, fcc.getCount());

        // Removing the only condition triggers auto-remove of the collection
        fcc.removeCondition(0);
        assertEquals(0, fcc.getCount());
        assertEquals(0, cfc.getCount());
    }

    // =========================================================================
    // CF-20  addCondition(CELL_VALUE, BETWEEN, "10", "20") returns index 0
    // =========================================================================

    /**
     * Verifies that add cell value condition returns index zero.
     */
    @Test
    void CF_20_addCellValueConditionReturnsIndexZero() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        int condIdx = fcc.addCondition(
                FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "10", "20");
        assertEquals(0, condIdx);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-21  addCondition(EXPRESSION) increments count
    // =========================================================================

    /**
     * Verifies that add expression condition increments count.
     */
    @Test
    void CF_21_addExpressionConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.EXPRESSION);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-22  addCondition(CONTAINS_TEXT) increments count
    // =========================================================================

    /**
     * Verifies that add contains text condition increments count.
     */
    @Test
    void CF_22_addContainsTextConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CONTAINS_TEXT);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-23  addCondition(DUPLICATE_VALUES) increments count
    // =========================================================================

    /**
     * Verifies that add duplicate values condition increments count.
     */
    @Test
    void CF_23_addDuplicateValuesConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.DUPLICATE_VALUES);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-24  addCondition(UNIQUE_VALUES) increments count
    // =========================================================================

    /**
     * Verifies that add unique values condition increments count.
     */
    @Test
    void CF_24_addUniqueValuesConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.UNIQUE_VALUES);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-25  addCondition(TOP_10) increments count
    // =========================================================================

    /**
     * Verifies that add top 10 condition increments count.
     */
    @Test
    void CF_25_addTop10ConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.TOP_10);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-26  addCondition(BOTTOM_10) increments count
    // =========================================================================

    /**
     * Verifies that add bottom 10 condition increments count.
     */
    @Test
    void CF_26_addBottom10ConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.BOTTOM_10);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-27  addCondition(ABOVE_AVERAGE) increments count
    // =========================================================================

    /**
     * Verifies that add above average condition increments count.
     */
    @Test
    void CF_27_addAboveAverageConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.ABOVE_AVERAGE);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-28  addCondition(BELOW_AVERAGE) increments count
    // =========================================================================

    /**
     * Verifies that add below average condition increments count.
     */
    @Test
    void CF_28_addBelowAverageConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.BELOW_AVERAGE);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-29  addCondition(COLOR_SCALE) increments count
    // =========================================================================

    /**
     * Verifies that add color scale condition increments count.
     */
    @Test
    void CF_29_addColorScaleConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.COLOR_SCALE);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-2A  addCondition(DATA_BAR) increments count
    // =========================================================================

    /**
     * Verifies that add data bar condition increments count.
     */
    @Test
    void CF_2A_addDataBarConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.DATA_BAR);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-2B  addCondition(ICON_SET) increments count
    // =========================================================================

    /**
     * Verifies that add icon set condition increments count.
     */
    @Test
    void CF_2B_addIconSetConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.ICON_SET);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-2C  addCondition(TIME_PERIOD) increments count
    // =========================================================================

    /**
     * Verifies that add time period condition increments count.
     */
    @Test
    void CF_2C_addTimePeriodConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.TIME_PERIOD);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-2D  addCondition(NOT_CONTAINS_TEXT) increments count
    // =========================================================================

    /**
     * Verifies that add not contains text condition increments count.
     */
    @Test
    void CF_2D_addNotContainsTextConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.NOT_CONTAINS_TEXT);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-2E  addCondition(begins_with) increments count
    // =========================================================================

    /**
     * Verifies that add begins with condition increments count.
     */
    @Test
    void CF_2E_addBeginsWithConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.begins_with);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-2F  addCondition(ENDS_WITH) increments count
    // =========================================================================

    /**
     * Verifies that add ends with condition increments count.
     */
    @Test
    void CF_2F_addEndsWithConditionIncrementsCount() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.ENDS_WITH);
        assertEquals(1, fcc.getCount());
    }

    // =========================================================================
    // CF-30  Multiple conditions in one collection 鈥?count reflects all added
    // =========================================================================

    /**
     * Verifies that multiple conditions in one collection counts correctly.
     */
    @Test
    void CF_30_multipleConditionsInOneCollectionCountsCorrectly() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.LESS_THAN, "0", "");
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.GREATER_THAN, "100", "");
        fcc.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "ISBLANK(A1)", "");
        assertEquals(3, fcc.getCount());
    }

    // =========================================================================
    // CF-40  Priority: first condition across all collections gets priority 1
    // =========================================================================

    /**
     * Verifies that first condition gets priority one.
     */
    @Test

    void CF_40_firstConditionGetsPriorityOne() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        assertEquals(1, fcc.get(0).getPriority());
    }

    // =========================================================================
    // CF-41  Priority: second condition in same collection gets priority 2
    // =========================================================================

    /**
     * Verifies that second condition gets priority two.
     */
    @Test

    void CF_41_secondConditionGetsPriorityTwo() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        fcc.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "A1>0", "");
        assertEquals(2, fcc.get(1).getPriority());
    }

    // =========================================================================
    // CF-42  Priority: condition in second collection follows on from first collection
    // =========================================================================

    /**
     * Verifies that condition in second collection continues priority.
     */
    @Test

    void CF_42_conditionInSecondCollectionContinuesPriority() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc1 = cfc.get(cfc.add());
        fcc1.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        fcc1.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "A1>0", "");

        FormatConditionCollection fcc2 = cfc.get(cfc.add());
        fcc2.addCondition(FormatConditionType.TOP_10);

        // First condition in fcc2 should get priority 3 (max from fcc1 was 2)
        assertEquals(3, fcc2.get(0).getPriority());
    }

    // =========================================================================
    // CF-43  Priority: priorities are globally unique across all collections
    // =========================================================================

    /**
     * Verifies that priorities are globally unique.
     */
    @Test

    void CF_43_prioritiesAreGloballyUnique() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());

        FormatConditionCollection fcc1 = cfc.get(cfc.add());
        fcc1.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");

        FormatConditionCollection fcc2 = cfc.get(cfc.add());
        fcc2.addCondition(FormatConditionType.TOP_10);
        fcc2.addCondition(FormatConditionType.BOTTOM_10);

        int p1 = fcc1.get(0).getPriority();
        int p2 = fcc2.get(0).getPriority();
        int p3 = fcc2.get(1).getPriority();
        assertNotEquals(p1, p2);
        assertNotEquals(p1, p3);
        assertNotEquals(p2, p3);
    }

    // =========================================================================
    // CF-50  FormatCondition.getType() throws UnsupportedOperationException
    // =========================================================================

    /**
     * Verifies that get type throws unsupported operation exception.
     */
    @Test

    void CF_50_getTypeThrowsUnsupportedOperationException() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
        FormatCondition fc = fcc.get(0);
        // Once wired, should return CELL_VALUE
        assertEquals(FormatConditionType.CELL_VALUE, fc.getType());
    }

    // =========================================================================
    // CF-51  FormatCondition.getOperator() returns the operator that was set
    // =========================================================================

    /**
     * Verifies that get operator returns set value.
     */
    @Test

    void CF_51_getOperatorReturnsSetValue() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.GREATER_THAN, "0", "");
        FormatCondition fc = fcc.get(0);
        assertEquals(OperatorType.GREATER_THAN, fc.getOperator());
    }

    // =========================================================================
    // CF-52  FormatCondition.getFormula1() returns the formula that was set
    // =========================================================================

    /**
     * Verifies that get formula 1 returns set value.
     */
    @Test

    void CF_52_getFormula1ReturnsSetValue() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "10", "20");
        FormatCondition fc = fcc.get(0);
        assertEquals("10", fc.getFormula1());
    }

    // =========================================================================
    // CF-53  FormatCondition.getFormula2() returns the formula that was set
    // =========================================================================

    /**
     * Verifies that get formula 2 returns set value.
     */
    @Test

    void CF_53_getFormula2ReturnsSetValue() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "10", "20");
        FormatCondition fc = fcc.get(0);
        assertEquals("20", fc.getFormula2());
    }

    // =========================================================================
    // CF-54  TOP_10 condition: getTop() == true, getRank() == 10
    // =========================================================================

    /**
     * Verifies that top 10 defaults top true rank 10.
     */
    @Test

    void CF_54_top10DefaultsTopTrueRank10() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.TOP_10);
        FormatCondition fc = fcc.get(0);
        assertTrue(fc.getTop());
        assertEquals(10, fc.getRank());
    }

    // =========================================================================
    // CF-55  BOTTOM_10 condition: getTop() == false, getRank() == 10
    // =========================================================================

    /**
     * Verifies that bottom 10 defaults top false rank 10.
     */
    @Test

    void CF_55_bottom10DefaultsTopFalseRank10() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.BOTTOM_10);
        FormatCondition fc = fcc.get(0);
        assertFalse(fc.getTop());
        assertEquals(10, fc.getRank());
    }

    // =========================================================================
    // CF-56  ABOVE_AVERAGE condition: getAbove() == true
    // =========================================================================

    /**
     * Verifies that above average defaults above true.
     */
    @Test

    void CF_56_aboveAverageDefaultsAboveTrue() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.ABOVE_AVERAGE);
        FormatCondition fc = fcc.get(0);
        assertTrue(fc.getAbove());
    }

    // =========================================================================
    // CF-57  ICON_SET condition: getIconSetType() == "3TrafficLights1"
    // =========================================================================

    /**
     * Verifies that icon set default icon set type.
     */
    @Test

    void CF_57_iconSetDefaultIconSetType() {
        ConditionalFormattingCollection cfc = cfc(freshSheet());
        FormatConditionCollection fcc = cfc.get(cfc.add());
        fcc.addCondition(FormatConditionType.ICON_SET);
        FormatCondition fc = fcc.get(0);
        assertEquals("3TrafficLights1", fc.getIconSetType());
    }

    // =========================================================================
    // CF-60  XLSX serialisation: a CELL_VALUE condition round-trips through save/load
    // =========================================================================

    /**
     * Verifies that cell value condition roundtrips xlsx.
     */
    @Test

    void CF_60_cellValueConditionRoundtripsXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf60.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 10, 1));
                fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "100");
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                ConditionalFormattingCollection cfc = loaded.getWorksheets().get(0).getConditionalFormattings();
                assertEquals(1, cfc.getCount());
                FormatConditionCollection fcc = cfc.get(0);
                assertEquals(1, fcc.getCount());
                assertEquals(1, fcc.getRangeCount());
            }
        }
    }

    // =========================================================================
    // CF-61  XLSX serialisation: TOP_10 condition round-trips
    // =========================================================================

    /**
     * Verifies that top 10 condition roundtrips xlsx.
     */
    @Test

    void CF_61_top10ConditionRoundtripsXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf61.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 10, 1));
                fcc.addCondition(FormatConditionType.TOP_10);
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                ConditionalFormattingCollection cfc = loaded.getWorksheets().get(0).getConditionalFormattings();
                assertEquals(1, cfc.getCount());
                FormatConditionCollection fcc = cfc.get(0);
                assertEquals(1, fcc.getCount());
                FormatCondition fc = fcc.get(0);
                assertEquals(FormatConditionType.TOP_10, fc.getType());
                assertTrue(fc.getTop());
                assertEquals(10, fc.getRank());
            }
        }
    }

    // =========================================================================
    // CF-62  XLSX serialisation: COLOR_SCALE condition round-trips
    // =========================================================================

    /**
     * Verifies that color scale condition roundtrips xlsx.
     */
    @Test

    void CF_62_colorScaleConditionRoundtripsXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf62.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 10, 1));
                fcc.addCondition(FormatConditionType.COLOR_SCALE);
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                ConditionalFormattingCollection cfc = loaded.getWorksheets().get(0).getConditionalFormattings();
                assertEquals(1, cfc.getCount());
                FormatCondition fc = cfc.get(0).get(0);
                assertEquals(FormatConditionType.COLOR_SCALE, fc.getType());
            }
        }
    }

    // =========================================================================
    // CF-63  XLSX serialisation: DATA_BAR condition round-trips
    // =========================================================================

    /**
     * Verifies that data bar condition roundtrips xlsx.
     */
    @Test

    void CF_63_dataBarConditionRoundtripsXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf63.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 10, 1));
                fcc.addCondition(FormatConditionType.DATA_BAR);
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                ConditionalFormattingCollection cfc = loaded.getWorksheets().get(0).getConditionalFormattings();
                assertEquals(1, cfc.getCount());
                FormatCondition fc = cfc.get(0).get(0);
                assertEquals(FormatConditionType.DATA_BAR, fc.getType());
                assertNotNull(fc.getBarColor());
            }
        }
    }

    // =========================================================================
    // CF-64  XLSX serialisation: ICON_SET condition round-trips
    // =========================================================================

    /**
     * Verifies that icon set condition roundtrips xlsx.
     */
    @Test

    void CF_64_iconSetConditionRoundtripsXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf64.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 10, 1));
                fcc.addCondition(FormatConditionType.ICON_SET);
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                FormatCondition fc = loaded.getWorksheets().get(0)
                        .getConditionalFormattings().get(0).get(0);
                assertEquals(FormatConditionType.ICON_SET, fc.getType());
                assertEquals("3TrafficLights1", fc.getIconSetType());
            }
        }
    }

    // =========================================================================
    // CF-65  XLSX serialisation: multiple conditions in one collection all survive
    // =========================================================================

    /**
     * Verifies that multiple conditions survive xlsx roundtrip.
     */
    @Test

    void CF_65_multipleConditionsSurviveXlsxRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf65.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 10, 1));
                fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.LESS_THAN, "0", "");
                fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.GREATER_THAN, "100", "");
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                FormatConditionCollection fcc = loaded.getWorksheets().get(0)
                        .getConditionalFormattings().get(0);
                assertEquals(2, fcc.getCount());
            }
        }
    }

    // =========================================================================
    // CF-66  XLSX serialisation: multiple collections on one sheet all survive
    // =========================================================================

    /**
     * Verifies that multiple collections survive xlsx roundtrip.
     */
    @Test

    void CF_66_multipleCollectionsSurviveXlsxRoundtrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf66.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();

                FormatConditionCollection fcc1 = cfc.get(cfc.add());
                fcc1.addArea(new CellArea(0, 0, 5, 1));
                fcc1.addCondition(FormatConditionType.TOP_10);

                FormatConditionCollection fcc2 = cfc.get(cfc.add());
                fcc2.addArea(new CellArea(10, 0, 5, 1));
                fcc2.addCondition(FormatConditionType.BELOW_AVERAGE);

                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                ConditionalFormattingCollection cfc = loaded.getWorksheets().get(0).getConditionalFormattings();
                assertEquals(2, cfc.getCount());
            }
        }
    }

    // =========================================================================
    // CF-70  POI integration: CELL_VALUE condition written by Aspose is readable by POI
    // =========================================================================

    /**
     * Verifies that poi can read cell value condition written by aspose.
     */
    @Test

    void CF_70_poiCanReadCellValueConditionWrittenByAspose() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf70.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 10, 1));
                fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "100");
                wb.save(path);
            }

            try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook)
                    org.apache.poi.ss.usermodel.WorkbookFactory.create(new java.io.File(path))) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet = poiWb.getSheetAt(0);
                assertFalse(sheet.getSheetConditionalFormatting()
                        .getNumConditionalFormattings() == 0,
                        "Expected at least one conditional formatting rule");
            }
        }
    }

    // =========================================================================
    // CF-71  POI integration: condition written by POI is readable by Aspose
    // =========================================================================

    /**
     * Verifies that aspose can read condition written by poi.
     */
    @Test

    void CF_71_asposeCanReadConditionWrittenByPoi() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf71.xlsx");

            try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb =
                    new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet = poiWb.createSheet("Sheet1");
                org.apache.poi.ss.usermodel.SheetConditionalFormatting scf =
                        sheet.getSheetConditionalFormatting();
                org.apache.poi.ss.usermodel.ConditionalFormattingRule rule =
                        scf.createConditionalFormattingRule(
                                org.apache.poi.ss.usermodel.ComparisonOperator.BETWEEN, "1", "100");
                org.apache.poi.ss.util.CellRangeAddress[] regions =
                        { new org.apache.poi.ss.util.CellRangeAddress(0, 9, 0, 0) };
                scf.addConditionalFormatting(regions, rule);
                try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path)) {
                    poiWb.write(fos);
                }
            }

            try (Workbook wb = new Workbook(path)) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                assertEquals(1, cfc.getCount());
            }
        }
    }

    // =========================================================================
    // CF-72  XLSX serialisation: removeArea(CellArea) result survives round-trip
    // =========================================================================

    /**
     * Verifies that remove area result survives xlsx roundtrip.
     */
    @Test

    void CF_72_removeAreaResultSurvivesXlsxRoundtrip() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf72.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 4, 4));
                fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");
                fcc.removeArea(new CellArea(1, 0, 1, 2));
                // 3 areas remain after the subtraction
                assertEquals(3, fcc.getRangeCount());
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                FormatConditionCollection fcc = loaded.getWorksheets().get(0)
                        .getConditionalFormattings().get(0);
                assertEquals(3, fcc.getRangeCount());
            }
        }
    }

    // =========================================================================
    // CF-73  XLSX serialisation: removeArea(int, int, int, int) on the collection
    //        removes matching areas and the collection is gone when empty
    // =========================================================================

    /**
     * Verifies that collection level remove area empties and auto removes.
     */
    @Test

    void CF_73_collectionLevelRemoveAreaEmptiesAndAutoRemoves() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf73.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 2, 2));
                fcc.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "1", "10");

                // Remove the entire covered area via the collection-level API
                cfc.removeArea(0, 0, 2, 2);
                assertEquals(0, cfc.getCount());
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                assertEquals(0, loaded.getWorksheets().get(0)
                        .getConditionalFormattings().getCount());
            }
        }
    }

    // =========================================================================
    // CF-74  POI integration: DATA_BAR condition written by Aspose is readable by POI
    // =========================================================================

    /**
     * Verifies that poi can read data bar condition written by aspose.
     */
    @Test

    void CF_74_poiCanReadDataBarConditionWrittenByAspose() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf74.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 10, 1));
                fcc.addCondition(FormatConditionType.DATA_BAR);
                wb.save(path);
            }

            try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook)
                    org.apache.poi.ss.usermodel.WorkbookFactory.create(new java.io.File(path))) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet = poiWb.getSheetAt(0);
                assertTrue(sheet.getSheetConditionalFormatting()
                        .getNumConditionalFormattings() > 0);
            }
        }
    }

    // =========================================================================
    // CF-75  POI integration: ICON_SET condition written by Aspose is readable by POI
    // =========================================================================

    /**
     * Verifies that poi can read icon set condition written by aspose.
     */
    @Test

    void CF_75_poiCanReadIconSetConditionWrittenByAspose() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("ConditionalFormattingTest")) {
            String path = tempDir.getPath("cf75.xlsx");

            try (Workbook wb = new Workbook()) {
                ConditionalFormattingCollection cfc = wb.getWorksheets().get(0).getConditionalFormattings();
                FormatConditionCollection fcc = cfc.get(cfc.add());
                fcc.addArea(new CellArea(0, 0, 10, 1));
                fcc.addCondition(FormatConditionType.ICON_SET);
                wb.save(path);
            }

            try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook)
                    org.apache.poi.ss.usermodel.WorkbookFactory.create(new java.io.File(path))) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet = poiWb.getSheetAt(0);
                assertTrue(sheet.getSheetConditionalFormatting()
                        .getNumConditionalFormattings() > 0);
            }
        }
    }
}

