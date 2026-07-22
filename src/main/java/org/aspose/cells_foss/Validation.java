package org.aspose.cells_foss;

import org.aspose.cells_foss.core.ValidationModel;
import java.util.List;

/**
 * Represents a data validation rule applied to one or more cell areas.
 */
public final class Validation {
    private final List<ValidationModel> owner;
    private final ValidationModel model;

    /**
     * Initializes a new Validation instance.
     * @param owner owner
     * @param model model
     */
    Validation(List<ValidationModel> owner, ValidationModel model) {
        this.owner = owner;
        this.model = model;
    }

    /** Returns the list of cell areas covered by this validation. */
    public List<CellArea> getAreas() {
        return model.getAreas();
    }

    /**
     * Returns the type.
     * @return the requested result
     */
    public ValidationType getType() {
        return model.getType() != null ? model.getType() : ValidationType.ANY_VALUE;
    }

    /**
     * Sets the type.
     * @param type type
     */
    public void setType(ValidationType type) {
        model.setType(type);
    }

    /**
     * Returns the alert style.
     * @return the requested result
     */
    public ValidationAlertType getAlertStyle() {
        return model.getAlertStyle() != null ? model.getAlertStyle() : ValidationAlertType.STOP;
    }

    /**
     * Sets the alert style.
     * @param alertStyle alert style
     */
    public void setAlertStyle(ValidationAlertType alertStyle) {
        model.setAlertStyle(alertStyle);
    }

    /**
     * Returns the operator.
     * @return the requested result
     */
    public OperatorType getOperator() {
        return model.getOperator() != null ? model.getOperator() : OperatorType.NONE;
    }

    /**
     * Sets the operator.
     * @param operator operator
     */
    public void setOperator(OperatorType operator) {
        model.setOperator(operator);
    }

    /**
     * Returns the formula 1.
     * @return the requested result
     */
    public String getFormula1() {
        String v = model.getFormula1();
        if (v == null) return "";
        return v.startsWith("=") ? v.substring(1) : v;
    }

    /**
     * Sets the formula 1.
     * @param value value to apply
     */
    public void setFormula1(String value) {
        model.setFormula1(value);
    }

    /**
     * Returns the formula 2.
     * @return the requested result
     */
    public String getFormula2() {
        String v = model.getFormula2();
        if (v == null) return "";
        return v.startsWith("=") ? v.substring(1) : v;
    }

    /**
     * Sets the formula 2.
     * @param value value to apply
     */
    public void setFormula2(String value) {
        model.setFormula2(value);
    }

    /**
     * Returns the ignore blank.
     * @return the requested result
     */
    public boolean getIgnoreBlank() {
        return model.getIgnoreBlank();
    }

    /**
     * Sets the ignore blank.
     * @param ignoreBlank ignore blank
     */
    public void setIgnoreBlank(boolean ignoreBlank) {
        model.setIgnoreBlank(ignoreBlank);
    }

    /**
     * Returns the in cell drop down.
     * @return the requested result
     */
    public boolean getInCellDropDown() {
        return model.getInCellDropDown();
    }

    /**
     * Sets the in cell drop down.
     * @param inCellDropDown in cell drop down
     */
    public void setInCellDropDown(boolean inCellDropDown) {
        model.setInCellDropDown(inCellDropDown);
    }

    /**
     * Returns the input title.
     * @return the requested result
     */
    public String getInputTitle() {
        return model.getInputTitle() != null ? model.getInputTitle() : "";
    }

    /**
     * Sets the input title.
     * @param value value to apply
     */
    public void setInputTitle(String value) {
        model.setInputTitle(value);
    }

    /**
     * Returns the input message.
     * @return the requested result
     */
    public String getInputMessage() {
        return model.getInputMessage() != null ? model.getInputMessage() : "";
    }

    /**
     * Sets the input message.
     * @param value value to apply
     */
    public void setInputMessage(String value) {
        model.setInputMessage(value);
    }

    /**
     * Returns the error title.
     * @return the requested result
     */
    public String getErrorTitle() {
        return model.getErrorTitle() != null ? model.getErrorTitle() : "";
    }

    /**
     * Sets the error title.
     * @param value value to apply
     */
    public void setErrorTitle(String value) {
        model.setErrorTitle(value);
    }

    /**
     * Returns the error message.
     * @return the requested result
     */
    public String getErrorMessage() {
        return model.getErrorMessage() != null ? model.getErrorMessage() : "";
    }

    /**
     * Sets the error message.
     * @param value value to apply
     */
    public void setErrorMessage(String value) {
        model.setErrorMessage(value);
    }

    /**
     * Returns the show input.
     * @return the requested result
     */
    public boolean getShowInput() {
        return model.getShowInput();
    }

    /**
     * Sets the show input.
     * @param showInput show input
     */
    public void setShowInput(boolean showInput) {
        model.setShowInput(showInput);
    }

    /**
     * Returns the show error.
     * @return the requested result
     */
    public boolean getShowError() {
        return model.getShowError();
    }

    /**
     * Sets the show error.
     * @param showError show error
     */
    public void setShowError(boolean showError) {
        model.setShowError(showError);
    }

    /**
     * Adds area.
     * @param area area
     */
    public void addArea(CellArea area) {
        model.getAreas().add(area);
    }

    /**
     * Removes area.
     * @param area area
     */
    public void removeArea(CellArea area) {
        model.getAreas().remove(area);
    }

    /** Removes this validation rule from the collection. */
    public void delete() {
        owner.remove(model);
    }
}

