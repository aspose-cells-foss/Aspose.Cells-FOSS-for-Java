package org.aspose.cells_foss.unit;

import org.aspose.cells_foss.NumberFormat;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class NumberFormatTest {

    @Test
    void getBuiltInFormat_zero_returnsGeneral() {
        assertEquals("General", NumberFormat.getBuiltInFormat(0));
    }

    @Test
    void getBuiltInFormat_knownIds_returnExpected() {
        assertEquals("0",       NumberFormat.getBuiltInFormat(1));
        assertEquals("0.00",    NumberFormat.getBuiltInFormat(2));
        assertEquals("#,##0",   NumberFormat.getBuiltInFormat(3));
        assertEquals("0%",      NumberFormat.getBuiltInFormat(9));
        assertEquals("0.00%",   NumberFormat.getBuiltInFormat(10));
        assertEquals("m/d/yy",  NumberFormat.getBuiltInFormat(14));
        assertEquals("@",       NumberFormat.getBuiltInFormat(49));
    }

    @Test
    void getBuiltInFormat_unknownId_returnsGeneral() {
        assertEquals("General", NumberFormat.getBuiltInFormat(999));
    }

    @Test
    void isBuiltInFormat_knownCode_returnsTrue() {
        assertTrue(NumberFormat.isBuiltInFormat("General"));
        assertTrue(NumberFormat.isBuiltInFormat("0.00"));
        assertTrue(NumberFormat.isBuiltInFormat("0%"));
        assertTrue(NumberFormat.isBuiltInFormat("@"));
    }

    @Test
    void isBuiltInFormat_customCode_returnsFalse() {
        assertFalse(NumberFormat.isBuiltInFormat("##0.000"));
        assertFalse(NumberFormat.isBuiltInFormat("dd/MM/yyyy HH:mm:ss"));
    }

    @Test
    void getBuiltInFormatId_general_returnsZero() {
        assertEquals(Integer.valueOf(0), NumberFormat.getBuiltInFormatId("General"));
    }

    @Test
    void getBuiltInFormatId_knownCode_returnsId() {
        assertEquals(Integer.valueOf(1),  NumberFormat.getBuiltInFormatId("0"));
        assertEquals(Integer.valueOf(9),  NumberFormat.getBuiltInFormatId("0%"));
        assertEquals(Integer.valueOf(49), NumberFormat.getBuiltInFormatId("@"));
    }

    @Test
    void getBuiltInFormatId_unknownCode_returnsNull() {
        assertNull(NumberFormat.getBuiltInFormatId("##0.000"));
    }

    @Test
    void getBuiltInFormatId_nullInput_returnsZero() {
        assertEquals(Integer.valueOf(0), NumberFormat.getBuiltInFormatId(null));
    }

    @Test
    void getBuiltInFormatId_blankInput_returnsZero() {
        assertEquals(Integer.valueOf(0), NumberFormat.getBuiltInFormatId("  "));
    }

    @Test
    void resolveFormatCode_withCustom_returnsCustom() {
        assertEquals("dd/MM/yyyy", NumberFormat.resolveFormatCode(0, "dd/MM/yyyy"));
    }

    @Test
    void resolveFormatCode_withNullCustom_returnsBuiltIn() {
        assertEquals("0.00", NumberFormat.resolveFormatCode(2, null));
    }

    @Test
    void resolveFormatCode_withEmptyCustom_returnsBuiltIn() {
        assertEquals("General", NumberFormat.resolveFormatCode(0, ""));
    }

    @Test
    void roundTrip_allBuiltInIds_consistent() {
        for (int id = 0; id <= 49; id++) {
            String code = NumberFormat.getBuiltInFormat(id);
            if ("General".equals(code) && id != 0 && id != 50) continue;
            Integer recovered = NumberFormat.getBuiltInFormatId(code);
            if (recovered != null) {
                assertEquals(code, NumberFormat.getBuiltInFormat(recovered),
                    "Inconsistency for id=" + id);
            }
        }
    }
}

