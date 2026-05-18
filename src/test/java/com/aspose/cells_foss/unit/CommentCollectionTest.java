package com.aspose.cells_foss.unit;

import com.aspose.cells_foss.*;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class CommentCollectionTest {

    private Workbook wb;
    private CommentCollection comments;

    @BeforeEach
    void setUp() {
        wb = new Workbook();
        comments = wb.getWorksheets().get(0).getComments();
    }

    @AfterEach
    void tearDown() { wb.close(); }

    @Test
    void add_byRowColumn_createsComment() {
        Comment c = comments.add(0, 0);
        assertEquals(0, c.getRow());
        assertEquals(0, c.getColumn());
        assertEquals(1, comments.getCount());
    }

    @Test
    void add_byCellName_parsesAddress() {
        Comment c = comments.add("B3");
        assertEquals(2, c.getRow());
        assertEquals(1, c.getColumn());
    }

    @Test
    void add_duplicate_throws() {
        comments.add(1, 2);
        assertThrows(CellsException.class, () -> comments.add(1, 2));
    }

    @Test
    void add_negativeRow_throws() {
        assertThrows(CellsException.class, () -> comments.add(-1, 0));
    }

    @Test
    void add_negativeColumn_throws() {
        assertThrows(CellsException.class, () -> comments.add(0, -1));
    }

    @Test
    void get_byIndex_returnsCorrectComment() {
        comments.add(0, 0);
        comments.add(1, 1);
        assertEquals(1, comments.get(1).getRow());
    }

    @Test
    void get_outOfBounds_throws() {
        assertThrows(CellsException.class, () -> comments.get(0));
    }

    @Test
    void get_byCellName_returnsMatch() {
        comments.add(2, 3);
        Comment c = comments.get("D3");
        assertNotNull(c);
        assertEquals(2, c.getRow());
        assertEquals(3, c.getColumn());
    }

    @Test
    void get_byCellName_notFound_returnsNull() {
        assertNull(comments.get("A1"));
    }

    @Test
    void removeAt_byIndex_removesEntry() {
        comments.add(0, 0);
        comments.removeAt(0);
        assertEquals(0, comments.getCount());
    }

    @Test
    void removeAt_byIndex_outOfBounds_throws() {
        assertThrows(CellsException.class, () -> comments.removeAt(0));
    }

    @Test
    void removeAt_byCellName_removesEntry() {
        comments.add(0, 0);
        comments.removeAt("A1");
        assertEquals(0, comments.getCount());
    }

    @Test
    void removeAt_byCellName_notFound_isNoOp() {
        comments.add(0, 0);
        assertDoesNotThrow(() -> comments.removeAt("B2"));
        assertEquals(1, comments.getCount());
    }

    @Test
    void comment_propertiesRoundTrip() {
        Comment c = comments.add(0, 0);
        c.setAuthor("Alice");
        c.setNote("Review this cell");
        c.setVisible(true);
        c.setWidth(200);
        c.setHeight(100);

        Comment same = comments.get("A1");
        assertNotNull(same);
        assertEquals("Alice",            same.getAuthor());
        assertEquals("Review this cell", same.getNote());
        assertTrue(same.isVisible());
        assertEquals(200, same.getWidth());
        assertEquals(100, same.getHeight());
    }

    @Test
    void comment_setWidth_lessThanOne_throws() {
        Comment c = comments.add(0, 0);
        assertThrows(CellsException.class, () -> c.setWidth(0));
    }

    @Test
    void comment_setHeight_lessThanOne_throws() {
        Comment c = comments.add(0, 0);
        assertThrows(CellsException.class, () -> c.setHeight(-5));
    }

    @Test
    void comment_defaultValues_areReasonable() {
        Comment c = comments.add(0, 0);
        assertEquals("", c.getAuthor());
        assertEquals("", c.getNote());
        assertFalse(c.isVisible());
        assertTrue(c.getWidth() > 0);
        assertTrue(c.getHeight() > 0);
    }

    @Test
    void multipleSheets_commentsAreIndependent() {
        wb.getWorksheets().add("Sheet2");
        Worksheet sheet2 = wb.getWorksheets().get("Sheet2");
        comments.add(0, 0).setNote("Sheet1 note");
        sheet2.getComments().add(0, 0).setNote("Sheet2 note");

        assertEquals("Sheet1 note", comments.get("A1").getNote());
        assertEquals("Sheet2 note", sheet2.getComments().get("A1").getNote());
    }
}
