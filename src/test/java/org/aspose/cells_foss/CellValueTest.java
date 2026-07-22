package org.aspose.cells_foss;

import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.File;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 tests for Cell Value API 鈥?CV-* test cases.
 */
class CellValueTest {

    // =========================================================================
    // 1.1 Primitive Write and Read
    // =========================================================================

    /**
     * Verifies that put value string.
     */
    @Test
    void CV_01_putValueString() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue("Hello");
        assertEquals("Hello", cell.getValue());
        assertEquals(CellValueType.STRING, cell.getType());
    }

    /**
     * Verifies that put value int.
     */
    @Test
    void CV_02_putValueInt() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(42);
        assertEquals(42, cell.getValue());
        assertEquals(CellValueType.NUMBER, cell.getType());
    }

    /**
     * Verifies that put value double.
     */
    @Test
    void CV_03_putValueDouble() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(3.14);
        assertEquals(3.14, (Double) cell.getValue(), 1e-9);
        assertEquals(CellValueType.NUMBER, cell.getType());
    }

    /**
     * Verifies that put value boolean true.
     */
    @Test
    void CV_04_putValueBooleanTrue() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(true);
        assertEquals(true, cell.getValue());
        assertEquals(CellValueType.BOOLEAN, cell.getType());
    }

    /**
     * Verifies that put value boolean false.
     */
    @Test
    void CV_05_putValueBooleanFalse() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(false);
        assertEquals(false, cell.getValue());
        assertEquals(CellValueType.BOOLEAN, cell.getType());
    }

    /**
     * Verifies that put value local date time.
     */
    @Test
    void CV_06_putValueLocalDateTime() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        LocalDateTime dt = LocalDateTime.of(2024, 5, 6, 7, 8, 9);
        cell.putValue(dt);
        assertEquals(CellValueType.DATE_TIME, cell.getType());
        assertEquals(dt, cell.getValue());
    }

    /**
     * Verifies that set value null clears cell.
     */
    @Test
    void CV_07_setValueNullClearsCell() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue("something");
        cell.setValue(null);
        assertEquals(CellValueType.BLANK, cell.getType());
        assertEquals("", cell.getStringValue());
    }

    /**
     * Verifies that set value object dispatches.
     */
    @Test
    void CV_08_setValueObjectDispatches() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.setValue(99);
        cell.setValue("x");
        assertEquals(CellValueType.STRING, cell.getType());
        assertEquals("x", cell.getValue());
    }

    // =========================================================================
    // 1.2 Formula
    // =========================================================================

    /**
     * Verifies that set formula without equal normalised.
     */
    @Test
    void CV_10_setFormulaWithoutEqualNormalised() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.setFormula("A1+B1");
        assertEquals("=A1+B1", cell.getFormula());
    }

    /**
     * Verifies that set formula with equal accepted.
     */
    @Test
    void CV_11_setFormulaWithEqualAccepted() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.setFormula("=A1+B1");
        assertEquals("=A1+B1", cell.getFormula());
    }

    /**
     * Verifies that get type returns formula after formula set.
     */
    @Test
    void CV_12_getTypeReturnsFormulaAfterFormulaSet() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(0);
        cell.setFormula("=1+1");
        assertEquals(CellValueType.FORMULA, cell.getType());
    }

    /**
     * Verifies that cached formula value persists.
     */
    @Test
    void CV_13_cachedFormulaValuePersists() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(20);
        cell.setFormula("=B1*2");
        // Cached value from putValue(20) should remain as string value
        assertEquals("20", cell.getStringValue());
    }

    /**
     * Verifies that clear formula via empty string.
     */
    @Test
    void CV_14_clearFormulaViaEmptyString() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.setFormula("=1+1");
        cell.setFormula("");
        cell.putValue(5);
        assertEquals("", cell.getFormula());
        assertEquals(CellValueType.NUMBER, cell.getType());
    }

    // =========================================================================
    // 1.3 String Value Formatting
    // =========================================================================

    /**
     * Verifies that integer formats without decimal.
     */
    @Test
    void CV_20_integerFormatsWithoutDecimal() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(123);
        assertEquals("123", cell.getStringValue());
    }

    /**
     * Verifies that double formats with decimal.
     */
    @Test
    void CV_21_doubleFormatsWithDecimal() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(1.5);
        assertEquals("1.5", cell.getStringValue());
    }

    /**
     * Verifies that boolean true formats as true.
     */
    @Test
    void CV_22_booleanTrueFormatsAsTRUE() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(true);
        assertEquals("TRUE", cell.getStringValue());
    }

    /**
     * Verifies that boolean false formats as false.
     */
    @Test
    void CV_23_booleanFalseFormatsAsFALSE() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(false);
        assertEquals("FALSE", cell.getStringValue());
    }

    /**
     * Verifies that blank cell string value is empty.
     */
    @Test
    void CV_24_blankCellStringValueIsEmpty() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        assertEquals("", cell.getStringValue());
    }

    /**
     * Verifies that large double scientific notation.
     */
    @Test
    void CV_25_largeDoubleScientificNotation() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        double val = 6.02214076E+23;
        cell.putValue(val);
        assertEquals(CellValueType.NUMBER, cell.getType());
        assertEquals(val, (Double) cell.getValue(), 1e9);
    }

    // =========================================================================
    // 1.4 Cell Address API
    // =========================================================================

    /**
     * Verifies that get cells get string accepts 1 notation.
     */
    @Test
    void CV_30_getCellsGetStringAcceptsA1Notation() {
        Workbook wb = new Workbook();
        Cells cells = wb.getWorksheets().get(0).getCells();
        cells.get("A1").putValue(1);
        assertEquals(1, cells.get("A1").getValue());
    }

    /**
     * Verifies that get cells get row col accepts zero based.
     */
    @Test
    void CV_31_getCellsGetRowColAcceptsZeroBased() {
        Workbook wb = new Workbook();
        Cells cells = wb.getWorksheets().get(0).getCells();
        cells.get(0, 0).putValue(1);
        assertEquals(1, cells.get("A1").getValue());
        assertEquals(cells.get(0, 0).getValue(), cells.get("A1").getValue());
    }

    /**
     * Verifies that invalid address throws cells exception.
     */
    @Test
    void CV_32_invalidAddressThrowsCellsException() {
        Workbook wb = new Workbook();
        Cells cells = wb.getWorksheets().get(0).getCells();
        assertThrows(CellsException.class, () -> cells.get("1A"));
    }

    // =========================================================================
    // 1.5 XLSX Round-Trip 鈥?Cell Values
    // =========================================================================

    /**
     * Verifies that string round trips through xlsx.
     */
    @Test
    void CV_40_stringRoundTripsThroughXlsx() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CellValueTest")) {
            String path = tempDir.getPath("basic.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().get("A1").putValue("RoundTrip");
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertEquals("RoundTrip", loaded.getWorksheets().get(0).getCells().get("A1").getValue());
        }
    }

    /**
     * Verifies that integer round trips without becoming double.
     */
    @Test
    void CV_41_integerRoundTripsWithoutBecomingDouble() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CellValueTest")) {
            String path = tempDir.getPath("basic.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().get("A1").putValue(777);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            String sv = loaded.getWorksheets().get(0).getCells().get("A1").getStringValue();
            assertFalse(sv.contains("."), "Integer string value should not contain a decimal point, was: " + sv);
        }
    }

    /**
     * Verifies that double round trips.
     */
    @Test
    void CV_42_doubleRoundTrips() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CellValueTest")) {
            String path = tempDir.getPath("basic.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().get("A1").putValue(2.71828);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertEquals(2.71828, (Double) loaded.getWorksheets().get(0).getCells().get("A1").getValue(), 1e-9);
        }
    }

    /**
     * Verifies that boolean round trips.
     */
    @Test
    void CV_43_booleanRoundTrips() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CellValueTest")) {
            String path = tempDir.getPath("basic.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().get("A1").putValue(true);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            Cell cell = loaded.getWorksheets().get(0).getCells().get("A1");
            assertEquals(CellValueType.BOOLEAN, cell.getType());
            assertEquals(true, cell.getValue());
        }
    }

    /**
     * Verifies that date time round trips 1900 system.
     */
    @Test
    void CV_44_dateTimeRoundTrips1900System() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CellValueTest")) {
            String path = tempDir.getPath("dates.xlsx");
            LocalDateTime dt = LocalDateTime.of(2024, 3, 15, 10, 30, 0);
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().get("A1").putValue(dt);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            Cell cell = loaded.getWorksheets().get(0).getCells().get("A1");
            assertEquals(CellValueType.DATE_TIME, cell.getType());
            LocalDateTime loaded_dt = (LocalDateTime) cell.getValue();
            assertEquals(dt.getYear(), loaded_dt.getYear());
            assertEquals(dt.getMonth(), loaded_dt.getMonth());
            assertEquals(dt.getDayOfMonth(), loaded_dt.getDayOfMonth());
            assertEquals(dt.getHour(), loaded_dt.getHour());
            assertEquals(dt.getMinute(), loaded_dt.getMinute());
            assertEquals(dt.getSecond(), loaded_dt.getSecond());
        }
    }

    /**
     * Verifies that date time round trips 1904 system.
     */
    @Test
    void CV_45_dateTimeRoundTrips1904System() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CellValueTest")) {
            String path = tempDir.getPath("dates-1904.xlsx");
            LocalDateTime dt = LocalDateTime.of(2024, 6, 1, 12, 0, 0);
            try (Workbook wb = new Workbook()) {
                wb.getSettings().setDate1904(true);
                wb.getWorksheets().get(0).getCells().get("A1").putValue(dt);
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertTrue(loaded.getSettings().getDate1904());
            Cell cell = loaded.getWorksheets().get(0).getCells().get("A1");
            assertEquals(CellValueType.DATE_TIME, cell.getType());
            LocalDateTime loaded_dt = (LocalDateTime) cell.getValue();
            assertEquals(dt.getYear(), loaded_dt.getYear());
            assertEquals(dt.getMonth(), loaded_dt.getMonth());
            assertEquals(dt.getDayOfMonth(), loaded_dt.getDayOfMonth());
        }
    }

    /**
     * Verifies that formula text round trips.
     */
    @Test
    void CV_46_formulaTextRoundTrips() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CellValueTest")) {
            String path = tempDir.getPath("formulas.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).getCells().get("A1").putValue(10);
                wb.getWorksheets().get(0).getCells().get("B1").setFormula("=A1*2");
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertEquals("=A1*2", loaded.getWorksheets().get(0).getCells().get("B1").getFormula());
        }
    }

    /**
     * Verifies that multi sheet values round trip.
     */
    @Test
    void CV_47_multiSheetValuesRoundTrip() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CellValueTest")) {
            String path = tempDir.getPath("multi.xlsx");
            try (Workbook wb = new Workbook()) {
                wb.getWorksheets().get(0).setName("Sheet1");
                wb.getWorksheets().get(0).getCells().get("A1").putValue("Alpha");
                wb.getWorksheets().add("Sheet2");
                wb.getWorksheets().get(1).getCells().get("A1").putValue("Beta");
                wb.save(path);
            }
            Workbook loaded = new Workbook(path);
            assertEquals("Alpha", loaded.getWorksheets().get(0).getCells().get("A1").getValue());
            assertEquals("Beta", loaded.getWorksheets().get(1).getCells().get("A1").getValue());
        }
    }

    // =========================================================================
    // 1.6 Integration: Write All Cell Types 鈥?Verify via API and POI
    // =========================================================================

    /**
     * Verifies that to cv 59 integration all cell types.
     */
    @Test
    void CV_50_to_CV_59_integrationAllCellTypes() throws Exception {
        try (TemporaryDirectory tempDir = new TemporaryDirectory("CellValueTest")) {
            String path = tempDir.getPath("cell-values.xlsx");
            LocalDateTime dt = LocalDateTime.of(2024, 3, 15, 10, 30, 0);

            try (Workbook wb = new Workbook()) {
                Cells cells = wb.getWorksheets().get(0).getCells();
                cells.get("A1").putValue("Hello");          // CV-50 String
                cells.get("A2").putValue(42);               // CV-51 Integer
                cells.get("A3").putValue(3.14);             // CV-52 Double
                cells.get("A4").putValue(true);             // CV-53 Boolean true
                cells.get("A5").putValue(false);            // CV-54 Boolean false
                cells.get("A6").putValue(dt);               // CV-55 DateTime
                cells.get("A7").setFormula("=A2*2");        // CV-56 Formula
                // A8 never written                         // CV-57 Blank
                cells.get("A9").putValue(1_000_000);        // CV-58 Large integer
                cells.get("A10").putValue(-0.001);          // CV-59 Negative double
                wb.save(path);
            }

            // --- API verification ---
            Workbook loaded = new Workbook(path);
            Cells cells = loaded.getWorksheets().get(0).getCells();

            // CV-50
            assertEquals("Hello", cells.get("A1").getValue());
            assertEquals(CellValueType.STRING, cells.get("A1").getType());

            // CV-51
            assertEquals(CellValueType.NUMBER, cells.get("A2").getType());
            // Value should be 42 (stored as int or double, check numeric equality)
            assertEquals(42.0, ((Number) cells.get("A2").getValue()).doubleValue(), 1e-9);

            // CV-52
            assertEquals(3.14, ((Number) cells.get("A3").getValue()).doubleValue(), 1e-9);

            // CV-53
            assertEquals(CellValueType.BOOLEAN, cells.get("A4").getType());
            assertEquals(true, cells.get("A4").getValue());

            // CV-54
            assertEquals(false, cells.get("A5").getValue());

            // CV-55
            assertEquals(CellValueType.DATE_TIME, cells.get("A6").getType());
            LocalDateTime loaded_dt = (LocalDateTime) cells.get("A6").getValue();
            assertEquals(dt.getYear(), loaded_dt.getYear());
            assertEquals(dt.getMonth(), loaded_dt.getMonth());
            assertEquals(dt.getDayOfMonth(), loaded_dt.getDayOfMonth());

            // CV-56
            assertEquals("=A2*2", cells.get("A7").getFormula());
            assertEquals(CellValueType.FORMULA, cells.get("A7").getType());

            // CV-57 Blank
            assertEquals(CellValueType.BLANK, cells.get("A8").getType());

            // CV-58
            assertEquals(1000000.0, ((Number) cells.get("A9").getValue()).doubleValue(), 1e-9);

            // CV-59
            assertEquals(-0.001, ((Number) cells.get("A10").getValue()).doubleValue(), 1e-12);

            // --- POI verification ---
            try (org.apache.poi.ss.usermodel.Workbook poiWb = WorkbookFactory.create(new File(path))) {
                org.apache.poi.ss.usermodel.Sheet poiSheet = poiWb.getSheetAt(0);

                // CV-50: String
                org.apache.poi.ss.usermodel.Row row0 = poiSheet.getRow(0);
                assertEquals(org.apache.poi.ss.usermodel.CellType.STRING, row0.getCell(0).getCellType());
                assertEquals("Hello", row0.getCell(0).getStringCellValue());

                // CV-51: Integer
                org.apache.poi.ss.usermodel.Row row1 = poiSheet.getRow(1);
                assertEquals(org.apache.poi.ss.usermodel.CellType.NUMERIC, row1.getCell(0).getCellType());
                assertEquals(42, (int) row1.getCell(0).getNumericCellValue());

                // CV-52: Double
                org.apache.poi.ss.usermodel.Row row2 = poiSheet.getRow(2);
                assertEquals(3.14, row2.getCell(0).getNumericCellValue(), 1e-9);

                // CV-53: Boolean true
                org.apache.poi.ss.usermodel.Row row3 = poiSheet.getRow(3);
                assertEquals(org.apache.poi.ss.usermodel.CellType.BOOLEAN, row3.getCell(0).getCellType());
                assertTrue(row3.getCell(0).getBooleanCellValue());

                // CV-54: Boolean false
                org.apache.poi.ss.usermodel.Row row4 = poiSheet.getRow(4);
                assertFalse(row4.getCell(0).getBooleanCellValue());

                // CV-55: DateTime
                org.apache.poi.ss.usermodel.Row row5 = poiSheet.getRow(5);
                org.apache.poi.ss.usermodel.Cell dateCell = row5.getCell(0);
                assertTrue(DateUtil.isCellDateFormatted(dateCell), "Date cell should be formatted as date");
                Date javaDate = DateUtil.getJavaDate(dateCell.getNumericCellValue());
                LocalDateTime poiDt = javaDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
                assertEquals(dt.getYear(), poiDt.getYear());
                assertEquals(dt.getMonth(), poiDt.getMonth());
                assertEquals(dt.getDayOfMonth(), poiDt.getDayOfMonth());

                // CV-56: Formula
                org.apache.poi.ss.usermodel.Row row6 = poiSheet.getRow(6);
                assertEquals(org.apache.poi.ss.usermodel.CellType.FORMULA, row6.getCell(0).getCellType());
                assertEquals("A2*2", row6.getCell(0).getCellFormula());

                // CV-57: Blank
                org.apache.poi.ss.usermodel.Row row7 = poiSheet.getRow(7);
                assertTrue(row7 == null || row7.getCell(0) == null
                        || row7.getCell(0).getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK,
                        "Cell A8 should be blank");

                // CV-58: Large integer
                org.apache.poi.ss.usermodel.Row row8 = poiSheet.getRow(8);
                assertEquals(1000000, (int) row8.getCell(0).getNumericCellValue());

                // CV-59: Negative double
                org.apache.poi.ss.usermodel.Row row9 = poiSheet.getRow(9);
                assertEquals(-0.001, row9.getCell(0).getNumericCellValue(), 1e-12);
            }
        }
    }
}

