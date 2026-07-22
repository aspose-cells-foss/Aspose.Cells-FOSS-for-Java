package org.aspose.cells_foss.unit;

import org.aspose.cells_foss.ListObjectCollection;
import org.aspose.cells_foss.TableStyleType;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class TableStyleTypeTest {

    @Test
    void allNamedStyles_haveNonBlankName() {
        for (TableStyleType t : TableStyleType.values()) {
            if (t == TableStyleType.NONE || t == TableStyleType.CUSTOM) continue;
            String name = ListObjectCollection.tableStyleTypeName(t);
            assertFalse(name.isBlank(), "Expected non-blank name for " + t);
        }
    }

    @Test
    void noneStyle_hasEmptyName() {
        assertEquals("", ListObjectCollection.tableStyleTypeName(TableStyleType.NONE));
    }

    @Test
    void customStyle_hasEmptyName() {
        assertEquals("", ListObjectCollection.tableStyleTypeName(TableStyleType.CUSTOM));
    }

    @Test
    void spotCheck_medium2_name() {
        assertEquals("TableStyleMedium2",
            ListObjectCollection.tableStyleTypeName(TableStyleType.TABLE_STYLE_MEDIUM_2));
    }

    @Test
    void spotCheck_dark1_name() {
        assertEquals("TableStyleDark1",
            ListObjectCollection.tableStyleTypeName(TableStyleType.TABLE_STYLE_DARK_1));
    }

    @Test
    void spotCheck_light1_name() {
        assertEquals("TableStyleLight1",
            ListObjectCollection.tableStyleTypeName(TableStyleType.TABLE_STYLE_LIGHT_1));
    }

    @Test
    void parseTableStyleType_medium2_roundTrip() {
        TableStyleType t = ListObjectCollection.parseTableStyleType("TableStyleMedium2");
        assertEquals(TableStyleType.TABLE_STYLE_MEDIUM_2, t);
    }

    @Test
    void parseTableStyleType_allBuiltin_roundTrip() {
        for (TableStyleType t : TableStyleType.values()) {
            if (t == TableStyleType.NONE || t == TableStyleType.CUSTOM) continue;
            String name = ListObjectCollection.tableStyleTypeName(t);
            assertEquals(t, ListObjectCollection.parseTableStyleType(name),
                "Round-trip failed for " + t);
        }
    }

    @Test
    void parseTableStyleType_unknownName_returnsCustom() {
        assertEquals(TableStyleType.CUSTOM,
            ListObjectCollection.parseTableStyleType("MyPrivateStyle"));
    }

    @Test
    void parseTableStyleType_blankName_returnsNone() {
        assertEquals(TableStyleType.NONE,
            ListObjectCollection.parseTableStyleType(""));
    }

    @Test
    void parseTableStyleType_null_returnsNone() {
        assertEquals(TableStyleType.NONE,
            ListObjectCollection.parseTableStyleType(null));
    }
}

