package com.aspose.cells_foss.validation;

import java.util.Collections;
import java.util.List;

/**
 * A validator for workbook models that produces validation messages.
 */
public final class WorkbookValidator {

    /**
     * Validates the workbook model for saving.
     *
     * @param workbookModel the workbook model to validate
     * @param packageModel  the package model to validate (nullable)
     * @return an empty list of validation messages (no validation is performed)
     */
    public List<ValidationMessage> validateForSave(Object workbookModel, Object packageModel) {
        return Collections.emptyList();
    }
}