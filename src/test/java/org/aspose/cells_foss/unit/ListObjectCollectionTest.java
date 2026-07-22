package org.aspose.cells_foss.unit;

import org.aspose.cells_foss.*;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class ListObjectCollectionTest {

    private Workbook wb;
    private ListObjectCollection tables;

    @BeforeEach
    void setUp() {
        wb = new Workbook();
        tables = wb.getWorksheets().get(0).getListObjects();
    }

    @AfterEach
    void tearDown() { wb.close(); }

    // ---- Add by coordinates ----

    @Test
    void add_validRange_returnsZeroIndex() {
        assertEquals(0, tables.add(0, 0, 4, 3, true));
        assertEquals(1, tables.getCount());
    }

    @Test
    void add_byCellName_parsesRange() {
        int idx = tables.add("A1", "D5", true);
        assertEquals(0, idx);
        ListObject t = tables.get(0);
        assertEquals(0, t.getStartRow());
        assertEquals(0, t.getStartColumn());
        assertEquals(4, t.getEndRow());
        assertEquals(3, t.getEndColumn());
    }

    @Test
    void add_negativeStart_throws() {
        assertThrows(CellsException.class, () -> tables.add(-1, 0, 4, 3, true));
    }

    @Test
    void add_endRowSmallerThanStart_throws() {
        assertThrows(CellsException.class, () -> tables.add(5, 0, 3, 3, true));
    }

    @Test
    void add_overlappingRanges_throws() {
        tables.add(0, 0, 5, 5, true);
        assertThrows(CellsException.class, () -> tables.add(3, 3, 8, 8, true));
    }

    // ---- Get ----

    @Test
    void get_byIndex_returnsTable() {
        tables.add(0, 0, 4, 3, true);
        ListObject t = tables.get(0);
        assertNotNull(t);
    }

    @Test
    void get_byName_caseInsensitive() {
        tables.add(0, 0, 4, 3, true);
        tables.get(0).setDisplayName("Products");
        assertNotNull(tables.get("products"));
        assertNotNull(tables.get("PRODUCTS"));
    }

    @Test
    void get_byName_notFound_throws() {
        assertThrows(CellsException.class, () -> tables.get("NonExistent"));
    }

    @Test
    void get_outOfBounds_throws() {
        assertThrows(CellsException.class, () -> tables.get(0));
    }

    // ---- Remove ----

    @Test
    void removeAt_removesEntry() {
        tables.add(0, 0, 4, 3, true);
        tables.removeAt(0);
        assertEquals(0, tables.getCount());
    }

    // ---- ListObject properties ----

    @Test
    void displayName_defaultGenerated() {
        tables.add(0, 0, 4, 3, true);
        assertFalse(tables.get(0).getDisplayName().isBlank());
    }

    @Test
    void displayName_set_roundTrips() {
        tables.add(0, 0, 4, 3, true);
        tables.get(0).setDisplayName("Inventory");
        assertEquals("Inventory", tables.get(0).getDisplayName());
    }

    @Test
    void displayName_withSpace_throws() {
        tables.add(0, 0, 4, 3, true);
        assertThrows(CellsException.class, () -> tables.get(0).setDisplayName("My Table"));
    }

    @Test
    void displayName_blank_throws() {
        tables.add(0, 0, 4, 3, true);
        assertThrows(CellsException.class, () -> tables.get(0).setDisplayName(""));
    }

    @Test
    void showHeaderRow_defaultTrue() {
        tables.add(0, 0, 4, 3, true);
        assertTrue(tables.get(0).isShowHeaderRow());
    }

    @Test
    void showTotals_roundTrips() {
        tables.add(0, 0, 4, 3, true);
        tables.get(0).setShowTotals(true);
        assertTrue(tables.get(0).isShowTotals());
    }

    @Test
    void tableStyleType_set_roundTrips() {
        tables.add(0, 0, 4, 3, true);
        ListObject t = tables.get(0);
        t.setTableStyleType(TableStyleType.TABLE_STYLE_MEDIUM_2);
        assertEquals(TableStyleType.TABLE_STYLE_MEDIUM_2, t.getTableStyleType());
    }

    @Test
    void comment_roundTrips() {
        tables.add(0, 0, 4, 3, true);
        tables.get(0).setComment("Sales data");
        assertEquals("Sales data", tables.get(0).getComment());
    }

    // ---- ListColumn ----

    @Test
    void listColumns_countMatchesColumnSpan() {
        tables.add(0, 0, 4, 3, true);
        assertEquals(4, tables.get(0).getListColumns().getCount());
    }

    @Test
    void listColumn_name_set_roundTrips() {
        tables.add(0, 0, 4, 3, true);
        ListColumn col = tables.get(0).getListColumns().get(0);
        col.setName("ProductName");
        assertEquals("ProductName", tables.get(0).getListColumns().get(0).getName());
    }

    @Test
    void listColumn_totalsCalculation_roundTrips() {
        tables.add(0, 0, 4, 3, true);
        ListColumn col = tables.get(0).getListColumns().get(1);
        col.setTotalsCalculation(TotalsCalculation.SUM);
        assertEquals(TotalsCalculation.SUM, col.getTotalsCalculation());
    }

    // ---- Resize ----

    @Test
    void resize_updatesRange() {
        tables.add(0, 0, 4, 3, true);
        tables.get(0).resize(1, 1, 6, 5, true);
        ListObject t = tables.get(0);
        assertEquals(1, t.getStartRow());
        assertEquals(5, t.getEndColumn());
    }

    // ---- TableStyleType helpers ----

    @Test
    void parseTableStyleType_roundTrip() {
        for (TableStyleType t : TableStyleType.values()) {
            if (t == TableStyleType.NONE || t == TableStyleType.CUSTOM) continue;
            String name = ListObjectCollection.tableStyleTypeName(t);
            assertEquals(t, ListObjectCollection.parseTableStyleType(name),
                "Round-trip failed for " + t);
        }
    }

    @Test
    void parseTableStyleType_blank_returnsNone() {
        assertEquals(TableStyleType.NONE, ListObjectCollection.parseTableStyleType(""));
    }

    @Test
    void parseTableStyleType_unknown_returnsCustom() {
        assertEquals(TableStyleType.CUSTOM, ListObjectCollection.parseTableStyleType("MyCustomStyle"));
    }
}

