package org.aspose.cells_foss;

import static org.junit.jupiter.api.Assertions.*;

/** Factory for data-validation test workbooks. Mirrors C# ValidationScenarioFactory. */
public final class ValidationScenarioFactory {

    private ValidationScenarioFactory() {}

    public static Workbook createValidationWorkbook() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Validated");

        // Whole-number validation on A1:B5
        int idx = sheet.getValidations().add(CellArea.createCellArea("A1", "B5"));
        Validation v1 = sheet.getValidations().get(idx);
        v1.setType(ValidationType.WHOLE_NUMBER);
        v1.setOperator(OperatorType.BETWEEN);
        v1.setFormula1("1");
        v1.setFormula2("100");
        v1.setShowError(true);
        v1.setErrorTitle("Invalid");
        v1.setErrorMessage("Enter 1-100");

        // List validation on C1:C10
        int idx2 = sheet.getValidations().add(CellArea.createCellArea("C1", "C10"));
        Validation v2 = sheet.getValidations().get(idx2);
        v2.setType(ValidationType.LIST);
        v2.setFormula1("\"Yes,No,Maybe\"");

        // Decimal validation on D1
        int idx3 = sheet.getValidations().add(CellArea.createCellArea("D1", "D1"));
        Validation v3 = sheet.getValidations().get(idx3);
        v3.setType(ValidationType.DECIMAL);
        v3.setOperator(OperatorType.GREATER_THAN);
        v3.setFormula1("0");

        return wb;
    }

    public static void assertValidations(Workbook wb) {
        Worksheet sheet = wb.getWorksheets().get(0);
        assertEquals(3, sheet.getValidations().getCount());

        Validation v1 = sheet.getValidations().getValidationInCell(0, 0);
        assertNotNull(v1);
        assertEquals(ValidationType.WHOLE_NUMBER, v1.getType());
        assertEquals(OperatorType.BETWEEN, v1.getOperator());
        assertEquals("1", v1.getFormula1());
        assertEquals("100", v1.getFormula2());
        assertTrue(v1.getShowError());

        Validation v2 = sheet.getValidations().getValidationInCell(0, 2);
        assertNotNull(v2);
        assertEquals(ValidationType.LIST, v2.getType());

        Validation v3 = sheet.getValidations().getValidationInCell(0, 3);
        assertNotNull(v3);
        assertEquals(ValidationType.DECIMAL, v3.getType());
        assertEquals(OperatorType.GREATER_THAN, v3.getOperator());
    }
}

