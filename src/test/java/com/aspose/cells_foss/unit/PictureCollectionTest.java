package com.aspose.cells_foss.unit;

import com.aspose.cells_foss.*;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class PictureCollectionTest {

    static final byte[] PNG_BYTES  = {(byte)0x89,'P','N','G',0,0,0,0};
    static final byte[] JPEG_BYTES = {(byte)0xFF,(byte)0xD8,0,0};
    static final byte[] GIF_BYTES  = {'G','I','F','8','9','a',0,0};
    static final byte[] BMP_BYTES  = {'B','M',0,0};

    private Workbook wb;
    private PictureCollection pictures;

    @BeforeEach
    void setUp() {
        wb = new Workbook();
        pictures = wb.getWorksheets().get(0).getPictures();
    }

    @AfterEach
    void tearDown() { wb.close(); }

    @Test
    void add_validPng_returnsZeroIndex() {
        int idx = pictures.add(0, 0, 5, 5, PNG_BYTES);
        assertEquals(0, idx);
        assertEquals(1, pictures.getCount());
    }

    @Test
    void add_multipleImages_incrementsIndex() {
        assertEquals(0, pictures.add(0, 0, 2, 2, PNG_BYTES));
        assertEquals(1, pictures.add(3, 0, 5, 2, JPEG_BYTES));
    }

    @Test
    void add_nullData_throws() {
        assertThrows(CellsException.class, () -> pictures.add(0, 0, 1, 1, (byte[])null));
    }

    @Test
    void add_emptyData_throws() {
        assertThrows(CellsException.class, () -> pictures.add(0, 0, 1, 1, new byte[0]));
    }

    @Test
    void add_lowerRightSmallerThanUpperLeft_throws() {
        assertThrows(CellsException.class, () -> pictures.add(5, 5, 3, 3, PNG_BYTES));
    }

    @Test
    void add_negativeRow_throws() {
        assertThrows(CellsException.class, () -> pictures.add(-1, 0, 1, 1, PNG_BYTES));
    }

    @Test
    void get_anchors_roundTrip() {
        pictures.add(1, 2, 3, 4, PNG_BYTES);
        Picture p = pictures.get(0);
        assertEquals(1, p.getUpperLeftRow());
        assertEquals(2, p.getUpperLeftColumn());
        assertEquals(3, p.getLowerRightRow());
        assertEquals(4, p.getLowerRightColumn());
    }

    @Test
    void get_outOfBounds_throws() {
        assertThrows(CellsException.class, () -> pictures.get(0));
    }

    @Test
    void removeAt_removesEntry() {
        pictures.add(0, 0, 1, 1, PNG_BYTES);
        pictures.removeAt(0);
        assertEquals(0, pictures.getCount());
    }

    @Test
    void removeAt_outOfBounds_throws() {
        assertThrows(CellsException.class, () -> pictures.removeAt(0));
    }

    @Test
    void picture_setName_roundTrips() {
        pictures.add(0, 0, 1, 1, PNG_BYTES);
        Picture p = pictures.get(0);
        p.setName("Logo");
        assertEquals("Logo", pictures.get(0).getName());
    }

    @Test
    void picture_setUpperLeftRow_negative_throws() {
        pictures.add(0, 0, 1, 1, PNG_BYTES);
        assertThrows(CellsException.class, () -> pictures.get(0).setUpperLeftRow(-1));
    }

    // ---- Static detection helpers ----

    @Test
    void detectImageType_png()     { assertEquals(ImageType.PNG,     PictureCollection.detectImageType(PNG_BYTES)); }
    @Test
    void detectImageType_jpeg()    { assertEquals(ImageType.JPEG,    PictureCollection.detectImageType(JPEG_BYTES)); }
    @Test
    void detectImageType_gif()     { assertEquals(ImageType.GIF,     PictureCollection.detectImageType(GIF_BYTES)); }
    @Test
    void detectImageType_bmp()     { assertEquals(ImageType.BMP,     PictureCollection.detectImageType(BMP_BYTES)); }
    @Test
    void detectImageType_unknown() { assertEquals(ImageType.UNKNOWN, PictureCollection.detectImageType(new byte[]{1,2,3,4})); }
    @Test
    void detectImageType_tooShort(){ assertEquals(ImageType.UNKNOWN, PictureCollection.detectImageType(new byte[]{1})); }

    @Test
    void picture_imageType_reflectsData() {
        pictures.add(0, 0, 1, 1, JPEG_BYTES);
        assertEquals(ImageType.JPEG, pictures.get(0).getImageType());
    }

    @Test
    void contentType_png()     { assertEquals("image/png",                PictureCollection.contentTypeFromExtension("png")); }
    @Test
    void contentType_jpeg()    { assertEquals("image/jpeg",               PictureCollection.contentTypeFromExtension("jpeg")); }
    @Test
    void contentType_unknown() { assertEquals("application/octet-stream", PictureCollection.contentTypeFromExtension("xyz")); }
}
