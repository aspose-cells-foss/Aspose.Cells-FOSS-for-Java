package org.aspose.cells_foss;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 tests for the Data Validation API 鈥?DV-* test cases.
 *
 * DV-01 to DV-08 exercise the current (stub) behaviour and must pass now.
 * DV-10 to DV-35 are @Disabled pending full implementation.
 */
class DataValidationTest {

    // =========================================================================
    // Helper
    // =========================================================================

    /**
     * Verifies that fresh sheet.
     */
    private static Worksheet freshSheet() {
        // Workbook is intentionally not closed here: it is a lightweight in-memory
        // object with no external resources, and the returned Worksheet holds a
        // reference to it for the duration of the test.
        @SuppressWarnings("resource")
        Workbook wb = new Workbook();
        return wb.getWorksheets().get(0);
    }

    // =========================================================================
    // DV-01  getCount() on a new worksheet returns 0
    // =========================================================================

    /**
     * Verifies that new worksheet has zero validations.
     */
    @Test
    void DV_01_newWorksheetHasZeroValidations() {
        ValidationCollection vc = freshSheet().getValidations();
        assertEquals(0, vc.getCount());
    }

    // =========================================================================
    // DV-02  add() with a valid CellArea returns 0
    // =========================================================================

    /**
     * Verifies that add valid area returns zero.
     */
    @Test
    void DV_02_addValidAreaReturnsZero() {
        ValidationCollection vc = freshSheet().getValidations();
        CellArea area = new CellArea(0, 0, 1, 1);
        int index = vc.add(area);
        assertEquals(0, index);
    }

    // =========================================================================
    // DV-03  add() increments getCount()
    // =========================================================================

    /**
     * Verifies that add increments count.
     */
    @Test
    void DV_03_addIncrementsCount() {
        ValidationCollection vc = freshSheet().getValidations();
        vc.add(new CellArea(0, 0, 1, 1));
        assertEquals(1, vc.getCount());
    }

    // =========================================================================
    // DV-04  getValidationInCell() returns null for any valid cell
    // =========================================================================

    /**
     * Verifies that get validation in cell returns null.
     */
    @Test
    void DV_04_getValidationInCellReturnsNull() {
        ValidationCollection vc = freshSheet().getValidations();
        Object result = vc.getValidationInCell(5, 3);
        assertNull(result);
    }

    // =========================================================================
    // DV-05  getValidationInCell() with negative row throws CellsException
    // =========================================================================

    /**
     * Verifies that get validation in cell negative row throws.
     */
    @Test
    void DV_05_getValidationInCellNegativeRowThrows() {
        ValidationCollection vc = freshSheet().getValidations();
        assertThrows(CellsException.class, () -> vc.getValidationInCell(-1, 0));
    }

    // =========================================================================
    // DV-06  removeACell() silently accepts any valid non-negative coordinates
    // =========================================================================

    /**
     * Verifies that remove a cell silently accepts valid coordinates.
     */
    @Test
    void DV_06_removeACellSilentlyAcceptsValidCoordinates() {
        ValidationCollection vc = freshSheet().getValidations();
        assertDoesNotThrow(() -> vc.removeACell(0, 0));
        assertDoesNotThrow(() -> vc.removeACell(100, 200));
    }

    // =========================================================================
    // DV-07  Negative firstRow is rejected by CellArea constructor
    // =========================================================================

    /**
     * Verifies that add with negative first row throws.
     */
    @Test
    void DV_07_addWithNegativeFirstRowThrows() {
        // CellArea constructor throws IllegalArgumentException for negative firstRow
        assertThrows(IllegalArgumentException.class, () -> new CellArea(-1, 0, 1, 1));
    }

    // =========================================================================
    // DV-08  Zero totalRows is rejected by CellArea constructor
    // =========================================================================

    /**
     * Verifies that add with zero total rows throws.
     */
    @Test
    void DV_08_addWithZeroTotalRowsThrows() {
        // CellArea constructor throws IllegalArgumentException for zero totalRows
        assertThrows(IllegalArgumentException.class, () -> new CellArea(0, 0, 0, 1));
    }

    // =========================================================================
    // DV-10  In-memory: set/get ValidationType rounds trip
    // =========================================================================

    /**
     * Verifies that validation type roundtrip.
     */
    @Test
    void DV_10_validationTypeRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setType(ValidationType.WHOLE_NUMBER);
        assertEquals(ValidationType.WHOLE_NUMBER, v.getType());
    }

    // =========================================================================
    // DV-11  In-memory: set/get ValidationAlertType rounds trip
    // =========================================================================

    /**
     * Verifies that alert style roundtrip.
     */
    @Test
    void DV_11_alertStyleRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setAlertStyle(ValidationAlertType.WARNING);
        assertEquals(ValidationAlertType.WARNING, v.getAlertStyle());
    }

    // =========================================================================
    // DV-12  In-memory: set/get OperatorType rounds trip
    // =========================================================================

    /**
     * Verifies that operator type roundtrip.
     */
    @Test
    void DV_12_operatorTypeRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setOperator(OperatorType.BETWEEN);
        assertEquals(OperatorType.BETWEEN, v.getOperator());
    }

    // =========================================================================
    // DV-13  In-memory: set/get formula1 rounds trip
    // =========================================================================

    /**
     * Verifies that formula 1 roundtrip.
     */
    @Test
    void DV_13_formula1Roundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setFormula1("1");
        assertEquals("1", v.getFormula1());
    }

    // =========================================================================
    // DV-14  In-memory: set/get formula2 rounds trip
    // =========================================================================

    /**
     * Verifies that formula 2 roundtrip.
     */
    @Test
    void DV_14_formula2Roundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setFormula2("100");
        assertEquals("100", v.getFormula2());
    }

    // =========================================================================
    // DV-15  In-memory: setIgnoreBlank / getIgnoreBlank rounds trip
    // =========================================================================

    /**
     * Verifies that ignore blank roundtrip.
     */
    @Test
    void DV_15_ignoreBlankRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setIgnoreBlank(true);
        assertTrue(v.getIgnoreBlank());
        v.setIgnoreBlank(false);
        assertFalse(v.getIgnoreBlank());
    }

    // =========================================================================
    // DV-16  In-memory: setInCellDropDown / getInCellDropDown rounds trip
    // =========================================================================

    /**
     * Verifies that in cell drop down roundtrip.
     */
    @Test
    void DV_16_inCellDropDownRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setInCellDropDown(true);
        assertTrue(v.getInCellDropDown());
    }

    // =========================================================================
    // DV-17  In-memory: setInputTitle / getInputTitle rounds trip
    // =========================================================================

    /**
     * Verifies that input title roundtrip.
     */
    @Test
    void DV_17_inputTitleRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setInputTitle("Enter a value");
        assertEquals("Enter a value", v.getInputTitle());
    }

    // =========================================================================
    // DV-18  In-memory: setInputMessage / getInputMessage rounds trip
    // =========================================================================

    /**
     * Verifies that input message roundtrip.
     */
    @Test
    void DV_18_inputMessageRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setInputMessage("Value must be between 1 and 100");
        assertEquals("Value must be between 1 and 100", v.getInputMessage());
    }

    // =========================================================================
    // DV-19  In-memory: setErrorTitle / getErrorTitle rounds trip
    // =========================================================================

    /**
     * Verifies that error title roundtrip.
     */
    @Test
    void DV_19_errorTitleRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setErrorTitle("Invalid input");
        assertEquals("Invalid input", v.getErrorTitle());
    }

    // =========================================================================
    // DV-1A  In-memory: setErrorMessage / getErrorMessage rounds trip
    // =========================================================================

    /**
     * Verifies that error message roundtrip.
     */
    @Test
    void DV_1A_errorMessageRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setErrorMessage("Please enter a number between 1 and 100.");
        assertEquals("Please enter a number between 1 and 100.", v.getErrorMessage());
    }

    // =========================================================================
    // DV-1B  In-memory: setShowInput / getShowInput rounds trip
    // =========================================================================

    /**
     * Verifies that show input roundtrip.
     */
    @Test
    void DV_1B_showInputRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setShowInput(true);
        assertTrue(v.getShowInput());
    }

    // =========================================================================
    // DV-1C  In-memory: setShowError / getShowError rounds trip
    // =========================================================================

    /**
     * Verifies that show error roundtrip.
     */
    @Test
    void DV_1C_showErrorRoundtrip() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 1, 1));
        Validation v = (Validation) vc.get(idx);
        v.setShowError(true);
        assertTrue(v.getShowError());
    }

    // =========================================================================
    // DV-1D  In-memory: add() increments count once implemented
    // =========================================================================

    /**
     * Verifies that add increments counts.
     */
    @Test
    void DV_1D_addIncrementsCounts() {
        ValidationCollection vc = freshSheet().getValidations();
        vc.add(new CellArea(0, 0, 1, 1));
        assertEquals(1, vc.getCount());
        vc.add(new CellArea(2, 0, 1, 1));
        assertEquals(2, vc.getCount());
    }

    // =========================================================================
    // DV-20  XLSX round-trip: WholeNumber BETWEEN validation survives save/load
    // =========================================================================

    /**
     * Verifies that xlsx roundtrip whole number between.
     */
    @Test
    void DV_20_xlsxRoundtripWholeNumberBetween() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv20.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 10, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.WHOLE_NUMBER);
                v.setOperator(OperatorType.BETWEEN);
                v.setFormula1("1");
                v.setFormula2("100");
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                ValidationCollection vc = loaded.getWorksheets().get(0).getValidations();
                assertEquals(1, vc.getCount());
                Validation v = (Validation) vc.get(0);
                assertEquals(ValidationType.WHOLE_NUMBER, v.getType());
                assertEquals(OperatorType.BETWEEN, v.getOperator());
                assertEquals("1", v.getFormula1());
                assertEquals("100", v.getFormula2());
            }
        }
    }

    // =========================================================================
    // DV-21  XLSX round-trip: Decimal GREATER_THAN validation survives save/load
    // =========================================================================

    /**
     * Verifies that xlsx roundtrip decimal greater than.
     */
    @Test
    void DV_21_xlsxRoundtripDecimalGreaterThan() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv21.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(1, 0, 5, 3));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.DECIMAL);
                v.setOperator(OperatorType.GREATER_THAN);
                v.setFormula1("0.5");
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                Validation v = (Validation) loaded.getWorksheets().get(0).getValidations().get(0);
                assertEquals(ValidationType.DECIMAL, v.getType());
                assertEquals(OperatorType.GREATER_THAN, v.getOperator());
                assertEquals("0.5", v.getFormula1());
            }
        }
    }

    // =========================================================================
    // DV-22  XLSX round-trip: List validation survives save/load
    // =========================================================================

    /**
     * Verifies that xlsx roundtrip list.
     */
    @Test
    void DV_22_xlsxRoundtripList() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv22.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 1, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.LIST);
                v.setFormula1("\"Apple,Banana,Cherry\"");
                v.setInCellDropDown(true);
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                Validation v = (Validation) loaded.getWorksheets().get(0).getValidations().get(0);
                assertEquals(ValidationType.LIST, v.getType());
                assertTrue(v.getInCellDropDown());
            }
        }
    }

    // =========================================================================
    // DV-23  XLSX round-trip: TextLength validation survives save/load
    // =========================================================================

    /**
     * Verifies that xlsx roundtrip text length.
     */
    @Test
    void DV_23_xlsxRoundtripTextLength() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv23.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 1, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.TEXT_LENGTH);
                v.setOperator(OperatorType.LESS_OR_EQUAL);
                v.setFormula1("50");
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                Validation v = (Validation) loaded.getWorksheets().get(0).getValidations().get(0);
                assertEquals(ValidationType.TEXT_LENGTH, v.getType());
                assertEquals(OperatorType.LESS_OR_EQUAL, v.getOperator());
                assertEquals("50", v.getFormula1());
            }
        }
    }

    // =========================================================================
    // DV-24  XLSX round-trip: input title and message survive save/load
    // =========================================================================

    /**
     * Verifies that xlsx roundtrip input prompt.
     */
    @Test
    void DV_24_xlsxRoundtripInputPrompt() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv24.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 1, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.ANY_VALUE);
                v.setInputTitle("Hint");
                v.setInputMessage("Type anything here.");
                v.setShowInput(true);
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                Validation v = (Validation) loaded.getWorksheets().get(0).getValidations().get(0);
                assertEquals("Hint", v.getInputTitle());
                assertEquals("Type anything here.", v.getInputMessage());
                assertTrue(v.getShowInput());
            }
        }
    }

    // =========================================================================
    // DV-25  XLSX round-trip: error title, message and alert style survive save/load
    // =========================================================================

    /**
     * Verifies that xlsx roundtrip error alert.
     */
    @Test
    void DV_25_xlsxRoundtripErrorAlert() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv25.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 1, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.WHOLE_NUMBER);
                v.setOperator(OperatorType.GREATER_OR_EQUAL);
                v.setFormula1("0");
                v.setAlertStyle(ValidationAlertType.STOP);
                v.setErrorTitle("Bad value");
                v.setErrorMessage("Only non-negative integers are allowed.");
                v.setShowError(true);
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                Validation v = (Validation) loaded.getWorksheets().get(0).getValidations().get(0);
                assertEquals(ValidationAlertType.STOP, v.getAlertStyle());
                assertEquals("Bad value", v.getErrorTitle());
                assertEquals("Only non-negative integers are allowed.", v.getErrorMessage());
                assertTrue(v.getShowError());
            }
        }
    }

    // =========================================================================
    // DV-26  XLSX round-trip: ignoreBlank flag survives save/load
    // =========================================================================

    /**
     * Verifies that xlsx roundtrip ignore blank.
     */
    @Test
    void DV_26_xlsxRoundtripIgnoreBlank() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv26.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 1, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.WHOLE_NUMBER);
                v.setOperator(OperatorType.BETWEEN);
                v.setFormula1("1");
                v.setFormula2("10");
                v.setIgnoreBlank(false);
                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                Validation v = (Validation) loaded.getWorksheets().get(0).getValidations().get(0);
                assertFalse(v.getIgnoreBlank());
            }
        }
    }

    // =========================================================================
    // DV-27  XLSX round-trip: multiple validations on one sheet survive save/load
    // =========================================================================

    /**
     * Verifies that xlsx roundtrip multiple validations.
     */
    @Test
    void DV_27_xlsxRoundtripMultipleValidations() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv27.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();

                int idx1 = vc.add(new CellArea(0, 0, 1, 1));
                Validation v1 = (Validation) vc.get(idx1);
                v1.setType(ValidationType.WHOLE_NUMBER);
                v1.setOperator(OperatorType.BETWEEN);
                v1.setFormula1("1");
                v1.setFormula2("10");

                int idx2 = vc.add(new CellArea(5, 0, 1, 1));
                Validation v2 = (Validation) vc.get(idx2);
                v2.setType(ValidationType.DECIMAL);
                v2.setOperator(OperatorType.GREATER_THAN);
                v2.setFormula1("0");

                wb.save(path);
            }

            try (Workbook loaded = new Workbook(path)) {
                ValidationCollection vc = loaded.getWorksheets().get(0).getValidations();
                assertEquals(2, vc.getCount());
            }
        }
    }

    // =========================================================================
    // DV-30  POI integration: validation written by Aspose is readable by POI
    // =========================================================================

    /**
     * Verifies that poi can read validation written by aspose.
     */
    @Test
    void DV_30_poiCanReadValidationWrittenByAspose() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv30.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 1, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.WHOLE_NUMBER);
                v.setOperator(OperatorType.BETWEEN);
                v.setFormula1("1");
                v.setFormula2("100");
                wb.save(path);
            }

            try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook)
                    org.apache.poi.ss.usermodel.WorkbookFactory.create(new java.io.File(path))) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet = poiWb.getSheetAt(0);
                java.util.List<? extends org.apache.poi.ss.usermodel.DataValidation> dvs =
                        sheet.getDataValidations();
                assertEquals(1, dvs.size());
            }
        }
    }

    // =========================================================================
    // DV-31  POI integration: validation written by POI is readable by Aspose
    // =========================================================================

    /**
     * Verifies that aspose can read validation written by poi.
     */
    @Test
    void DV_31_asposeCanReadValidationWrittenByPoi() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv31.xlsx");

            try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb =
                    new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet = poiWb.createSheet("Sheet1");
                org.apache.poi.ss.util.CellRangeAddressList addressList =
                        new org.apache.poi.ss.util.CellRangeAddressList(0, 9, 0, 0);
                org.apache.poi.ss.usermodel.DataValidationHelper dvh = sheet.getDataValidationHelper();
                org.apache.poi.ss.usermodel.DataValidationConstraint constraint =
                        dvh.createIntegerConstraint(
                                org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.BETWEEN,
                                "1", "50");
                org.apache.poi.ss.usermodel.DataValidation dv =
                        dvh.createValidation(constraint, addressList);
                sheet.addValidationData(dv);
                try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path)) {
                    poiWb.write(fos);
                }
            }

            try (Workbook wb = new Workbook(path)) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                assertEquals(1, vc.getCount());
                Validation v = (Validation) vc.get(0);
                assertEquals(ValidationType.WHOLE_NUMBER, v.getType());
            }
        }
    }

    // =========================================================================
    // DV-32  POI integration: LIST validation written by Aspose is readable by POI
    // =========================================================================

    /**
     * Verifies that poi can read list validation written by aspose.
     */
    @Test
    void DV_32_poiCanReadListValidationWrittenByAspose() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv32.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 10, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.LIST);
                v.setFormula1("\"Yes,No,Maybe\"");
                v.setInCellDropDown(true);
                wb.save(path);
            }

            try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook)
                    org.apache.poi.ss.usermodel.WorkbookFactory.create(new java.io.File(path))) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet = poiWb.getSheetAt(0);
                assertEquals(1, sheet.getDataValidations().size());
            }
        }
    }

    // =========================================================================
    // DV-33  In-memory: getValidationInCell returns non-null for cells in area
    // =========================================================================

    /**
     * Verifies that get validation in cell returns non null after add.
     */
    @Test
    void DV_33_getValidationInCellReturnsNonNullAfterAdd() {
        ValidationCollection vc = freshSheet().getValidations();
        int idx = vc.add(new CellArea(0, 0, 5, 3));
        Validation v = (Validation) vc.get(idx);
        v.setType(ValidationType.WHOLE_NUMBER);
        v.setOperator(OperatorType.GREATER_THAN);
        v.setFormula1("0");

        // Cell within the area should return the validation
        assertNotNull(vc.getValidationInCell(2, 1));
        // Cell outside the area should return null
        assertNull(vc.getValidationInCell(10, 10));
    }

    // =========================================================================
    // DV-34  POI integration: CustomFormula validation written by Aspose is readable by POI
    // =========================================================================

    /**
     * Verifies that poi can read custom formula validation written by aspose.
     */
    @Test
    void DV_34_poiCanReadCustomFormulaValidationWrittenByAspose() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv34.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 1, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.CUSTOM);
                v.setFormula1("ISNUMBER(A1)");
                wb.save(path);
            }

            try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook)
                    org.apache.poi.ss.usermodel.WorkbookFactory.create(new java.io.File(path))) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet = poiWb.getSheetAt(0);
                assertEquals(1, sheet.getDataValidations().size());
            }
        }
    }

    // =========================================================================
    // DV-35  POI integration: Date validation written by Aspose is readable by POI
    // =========================================================================

    /**
     * Verifies that poi can read date validation written by aspose.
     */
    @Test
    void DV_35_poiCanReadDateValidationWrittenByAspose() throws Exception {
        // Wrap lower-level failures in the library-specific exception flow.
        try (TemporaryDirectory tempDir = new TemporaryDirectory("DataValidationTest")) {
            String path = tempDir.getPath("dv35.xlsx");

            try (Workbook wb = new Workbook()) {
                ValidationCollection vc = wb.getWorksheets().get(0).getValidations();
                int idx = vc.add(new CellArea(0, 0, 1, 1));
                Validation v = (Validation) vc.get(idx);
                v.setType(ValidationType.DATE);
                v.setOperator(OperatorType.GREATER_OR_EQUAL);
                v.setFormula1("\"2020-01-01\"");
                wb.save(path);
            }

            try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb =
                    (org.apache.poi.xssf.usermodel.XSSFWorkbook)
                    org.apache.poi.ss.usermodel.WorkbookFactory.create(new java.io.File(path))) {
                org.apache.poi.xssf.usermodel.XSSFSheet sheet = poiWb.getSheetAt(0);
                assertEquals(1, sheet.getDataValidations().size());
            }
        }
    }
}

