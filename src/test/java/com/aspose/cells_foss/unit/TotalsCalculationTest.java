package com.aspose.cells_foss.unit;

import com.aspose.cells_foss.ListColumn;
import com.aspose.cells_foss.TotalsCalculation;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class TotalsCalculationTest {

    @Test
    void allValues_roundTrip_viaTotalsCalculationString() {
        for (TotalsCalculation tc : TotalsCalculation.values()) {
            String raw = ListColumn.totalsCalculationToString(tc);
            TotalsCalculation recovered = ListColumn.parseTotalsCalculation(raw);
            assertEquals(tc, recovered, "Round-trip failed for " + tc);
        }
    }

    @Test
    void parseTotalsCalculation_null_returnsNone() {
        assertEquals(TotalsCalculation.NONE, ListColumn.parseTotalsCalculation(null));
    }

    @Test
    void parseTotalsCalculation_unknown_returnsNone() {
        assertEquals(TotalsCalculation.NONE, ListColumn.parseTotalsCalculation("bogus"));
    }

    @Test
    void parseTotalsCalculation_sum_caseInsensitive() {
        assertEquals(TotalsCalculation.SUM, ListColumn.parseTotalsCalculation("SUM"));
        assertEquals(TotalsCalculation.SUM, ListColumn.parseTotalsCalculation("sum"));
        assertEquals(TotalsCalculation.SUM, ListColumn.parseTotalsCalculation("Sum"));
    }

    @Test
    void totalsCalculationToString_none_returnsNoneString() {
        assertEquals("none", ListColumn.totalsCalculationToString(TotalsCalculation.NONE));
    }

    @Test
    void totalsCalculationToString_allValues_nonBlank() {
        for (TotalsCalculation tc : TotalsCalculation.values()) {
            assertFalse(ListColumn.totalsCalculationToString(tc).isBlank(),
                "Expected non-blank string for " + tc);
        }
    }
}
