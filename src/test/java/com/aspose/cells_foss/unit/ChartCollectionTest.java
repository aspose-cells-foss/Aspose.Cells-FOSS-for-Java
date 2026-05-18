package com.aspose.cells_foss.unit;

import com.aspose.cells_foss.*;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class ChartCollectionTest {

    @Test
    void parseChartType_knownValues_roundTrip() {
        for (ChartType t : ChartType.values()) {
            if (t == ChartType.UNKNOWN) continue;
            String name = ChartCollection.chartTypeToString(t);
            assertEquals(t, ChartCollection.parseChartType(name),
                "Round-trip failed for " + t);
        }
    }

    @Test
    void parseChartType_null_returnsUnknown() {
        assertEquals(ChartType.UNKNOWN, ChartCollection.parseChartType(null));
    }

    @Test
    void parseChartType_emptyString_returnsUnknown() {
        assertEquals(ChartType.UNKNOWN, ChartCollection.parseChartType(""));
    }

    @Test
    void parseChartType_caseInsensitive() {
        assertEquals(ChartType.COLUMN, ChartCollection.parseChartType("COLUMN"));
        assertEquals(ChartType.PIE,    ChartCollection.parseChartType("Pie"));
    }

    @Test
    void parseChartType_unrecognised_returnsUnknown() {
        assertEquals(ChartType.UNKNOWN, ChartCollection.parseChartType("foobar"));
    }

    @Test
    void chartCollection_newWorkbook_isEmpty() {
        try (Workbook wb = new Workbook()) {
            assertEquals(0, wb.getWorksheets().get(0).getCharts().getCount());
        }
    }

    @Test
    void chartCollection_get_outOfBounds_throws() {
        try (Workbook wb = new Workbook()) {
            ChartCollection charts = wb.getWorksheets().get(0).getCharts();
            assertThrows(CellsException.class, () -> charts.get(0));
        }
    }

    @Test
    void chartTypeToString_allKnownTypes_nonEmpty() {
        for (ChartType t : ChartType.values()) {
            if (t == ChartType.UNKNOWN) continue;
            String s = ChartCollection.chartTypeToString(t);
            assertFalse(s.isBlank(), "chartTypeToString returned blank for " + t);
        }
    }

    @Test
    void chartTypeEnum_containsExpectedValues() {
        assertNotNull(ChartType.BAR);
        assertNotNull(ChartType.COLUMN);
        assertNotNull(ChartType.LINE);
        assertNotNull(ChartType.PIE);
        assertNotNull(ChartType.WATERFALL);
        assertNotNull(ChartType.UNKNOWN);
    }
}
