package com.aspose.cells_foss;

import com.aspose.cells_foss.core.DefinedNameModel;

/**
 * Represents a defined name in the workbook.
 */
public final class DefinedName {
    private final Workbook workbook;
    private DefinedNameModel model;

    /**
     * Initializes a new DefinedName instance.
     * @param workbook workbook to apply
     * @param model model
     */
    DefinedName(Workbook workbook, DefinedNameModel model) {
        this.workbook = workbook;
        this.model = model;
    }

    /**
     * Returns the name.
     * @return the requested result
     */
    public String getName() {
        return model.getName();
    }

    /**
     * Sets the name.
     * @param value value to apply
     */
    public void setName(String value) {
        String normalized = DefinedNameUtility.normalizeName(value);
        workbook.ensureUniqueDefinedName(model, normalized, model.getLocalSheetIndex());
        model.setName(normalized);
    }

    /**
     * Returns the formula.
     * @return the requested result
     */
    public String getFormula() {
        return model.getFormula();
    }

    /**
     * Sets the formula.
     * @param value value to apply
     */
    public void setFormula(String value) {
        model.setFormula(DefinedNameUtility.normalizeFormula(value));
    }

    /**
     * Returns the local sheet index.
     * @return the requested result
     */
    public Integer getLocalSheetIndex() {
        return model.getLocalSheetIndex();
    }

    /**
     * Sets the local sheet index.
     * @param value value to apply
     */
    public void setLocalSheetIndex(Integer value) {
        workbook.ensureValidDefinedNameScope(value);
        workbook.ensureUniqueDefinedName(model, model.getName(), value);
        model.setLocalSheetIndex(value);
    }

    /**
     * Returns the hidden.
     * @return the requested result
     */
    public boolean getHidden() {
        return model.getHidden();
    }

    /**
     * Sets the hidden.
     * @param value value to apply
     */
    public void setHidden(boolean value) {
        model.setHidden(value);
    }

    /**
     * Returns the comment.
     * @return the requested result
     */
    public String getComment() {
        return model.getComment();
    }

    /**
     * Sets the comment.
     * @param value value to apply
     */
    public void setComment(String value) {
        model.setComment(DefinedNameUtility.normalizeComment(value));
    }
}