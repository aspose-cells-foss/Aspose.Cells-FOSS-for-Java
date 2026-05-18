package com.aspose.cells_foss.unit;

import com.aspose.cells_foss.*;
import com.aspose.cells_foss.core.AutoFilterModel;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class AutoFilterAdvancedTest {

    private Workbook wb;
    private AutoFilter autoFilter;

    @BeforeEach
    void setUp() {
        wb = new Workbook();
        autoFilter = wb.getWorksheets().get(0).getAutoFilter();
        autoFilter.setRange("A1:E1");
    }

    @AfterEach
    void tearDown() { wb.close(); }

    // ---- ColorFilter ----

    @Test
    void colorFilter_defaultDisabled() {
        int col = autoFilter.getFilterColumns().add(0);
        assertFalse(autoFilter.getFilterColumns().get(col).getColorFilter().isEnabled());
    }

    @Test
    void colorFilter_setDifferentialStyleId_enables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterColorFilter cf = autoFilter.getFilterColumns().get(col).getColorFilter();
        cf.setDifferentialStyleId(3);
        assertTrue(cf.isEnabled());
        assertEquals(Integer.valueOf(3), cf.getDifferentialStyleId());
    }

    @Test
    void colorFilter_setNegativeStyleId_throws() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterColorFilter cf = autoFilter.getFilterColumns().get(col).getColorFilter();
        assertThrows(CellsException.class, () -> cf.setDifferentialStyleId(-1));
    }

    @Test
    void colorFilter_setCellColor_enables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterColorFilter cf = autoFilter.getFilterColumns().get(col).getColorFilter();
        cf.setCellColor(true);
        assertTrue(cf.isEnabled());
        assertTrue(cf.isCellColor());
    }

    @Test
    void colorFilter_clear_disables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterColorFilter cf = autoFilter.getFilterColumns().get(col).getColorFilter();
        cf.setDifferentialStyleId(1);
        cf.clear();
        assertFalse(cf.isEnabled());
        assertNull(cf.getDifferentialStyleId());
    }

    @Test
    void colorFilter_setEnabledFalse_clears() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterColorFilter cf = autoFilter.getFilterColumns().get(col).getColorFilter();
        cf.setDifferentialStyleId(5);
        cf.setEnabled(false);
        assertFalse(cf.isEnabled());
    }

    // ---- DynamicFilter ----

    @Test
    void dynamicFilter_defaultDisabled() {
        int col = autoFilter.getFilterColumns().add(0);
        assertFalse(autoFilter.getFilterColumns().get(col).getDynamicFilter().isEnabled());
    }

    @Test
    void dynamicFilter_setType_enables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterDynamicFilter df = autoFilter.getFilterColumns().get(col).getDynamicFilter();
        df.setType("aboveAverage");
        assertTrue(df.isEnabled());
        assertEquals("aboveAverage", df.getType());
    }

    @Test
    void dynamicFilter_setValue_enables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterDynamicFilter df = autoFilter.getFilterColumns().get(col).getDynamicFilter();
        df.setValue(100.0);
        assertTrue(df.isEnabled());
        assertEquals(Double.valueOf(100.0), df.getValue());
    }

    @Test
    void dynamicFilter_setMaxValue_enables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterDynamicFilter df = autoFilter.getFilterColumns().get(col).getDynamicFilter();
        df.setMaxValue(999.9);
        assertTrue(df.isEnabled());
        assertEquals(Double.valueOf(999.9), df.getMaxValue());
    }

    @Test
    void dynamicFilter_clear_disables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterDynamicFilter df = autoFilter.getFilterColumns().get(col).getDynamicFilter();
        df.setType("aboveAverage");
        df.clear();
        assertFalse(df.isEnabled());
        assertTrue(df.getType().isEmpty());
        assertNull(df.getValue());
    }

    // ---- Top10 ----

    @Test
    void top10_defaultDisabled() {
        int col = autoFilter.getFilterColumns().add(0);
        assertFalse(autoFilter.getFilterColumns().get(col).getTop10().isEnabled());
    }

    @Test
    void top10_setTop_enables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterTop10 t10 = autoFilter.getFilterColumns().get(col).getTop10();
        t10.setTop(true);
        assertTrue(t10.isEnabled());
        assertTrue(t10.isTop());
    }

    @Test
    void top10_setValue_enables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterTop10 t10 = autoFilter.getFilterColumns().get(col).getTop10();
        t10.setValue(10.0);
        assertTrue(t10.isEnabled());
        assertEquals(Double.valueOf(10.0), t10.getValue());
    }

    @Test
    void top10_setPercent_enables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterTop10 t10 = autoFilter.getFilterColumns().get(col).getTop10();
        t10.setPercent(true);
        assertTrue(t10.isEnabled());
        assertTrue(t10.isPercent());
    }

    @Test
    void top10_setFilterValue_enables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterTop10 t10 = autoFilter.getFilterColumns().get(col).getTop10();
        t10.setFilterValue(500.0);
        assertTrue(t10.isEnabled());
        assertEquals(Double.valueOf(500.0), t10.getFilterValue());
    }

    @Test
    void top10_clear_disables() {
        int col = autoFilter.getFilterColumns().add(0);
        AutoFilter.AutoFilterTop10 t10 = autoFilter.getFilterColumns().get(col).getTop10();
        t10.setValue(10.0);
        t10.clear();
        assertFalse(t10.isEnabled());
        assertNull(t10.getValue());
    }

    // ---- SortState ----

    @Test
    void sortState_addCondition_increasesCount() {
        autoFilter.getSortState().getSortConditions().add("A1:A10");
        assertEquals(1, autoFilter.getSortState().getSortConditions().getCount());
    }

    @Test
    void sortState_condition_properties_roundTrip() {
        autoFilter.getSortState().getSortConditions().add("A1:A10");
        AutoFilter.AutoFilterSortCondition cond = autoFilter.getSortState().getSortConditions().get(0);
        cond.setDescending(true);
        cond.setSortBy("value");
        cond.setDifferentialStyleId(2);

        assertTrue(cond.isDescending());
        assertEquals("value", cond.getSortBy());
        assertEquals(Integer.valueOf(2), cond.getDifferentialStyleId());
    }

    @Test
    void sortState_condition_negativeDiffStyleId_throws() {
        autoFilter.getSortState().getSortConditions().add("A1:A10");
        AutoFilter.AutoFilterSortCondition cond = autoFilter.getSortState().getSortConditions().get(0);
        assertThrows(CellsException.class, () -> cond.setDifferentialStyleId(-1));
    }

    @Test
    void sortState_removeCondition_decreasesCount() {
        autoFilter.getSortState().getSortConditions().add("A1:A10");
        autoFilter.getSortState().getSortConditions().removeAt(0);
        assertEquals(0, autoFilter.getSortState().getSortConditions().getCount());
    }

    @Test
    void sortState_clear_removesAllConditions() {
        autoFilter.getSortState().getSortConditions().add("A1:A5");
        autoFilter.getSortState().getSortConditions().add("B1:B5");
        autoFilter.getSortState().clear();
        assertEquals(0, autoFilter.getSortState().getSortConditions().getCount());
    }

    @Test
    void sortState_columnSort_roundTrips() {
        autoFilter.getSortState().setColumnSort(true);
        assertTrue(autoFilter.getSortState().isColumnSort());
    }

    @Test
    void sortState_caseSensitive_roundTrips() {
        autoFilter.getSortState().setCaseSensitive(true);
        assertTrue(autoFilter.getSortState().isCaseSensitive());
    }

    // ---- Custom filter operator enum ----

    @Test
    void filterOperatorType_allValues_exist() {
        for (AutoFilterModel.FilterOperatorType op : AutoFilterModel.FilterOperatorType.values()) {
            assertNotNull(op);
        }
    }
}
