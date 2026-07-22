package org.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.List;
import org.aspose.cells_foss.ValidationType;
import org.aspose.cells_foss.CellArea;
import org.aspose.cells_foss.OperatorType;
import org.aspose.cells_foss.ValidationAlertType;

/**
 * Represents a data validation model in the Excel file.
 */
public final class ValidationModel {
    private final List<CellArea> areas = new ArrayList<>();
    private ValidationType type;
    private ValidationAlertType alertStyle = ValidationAlertType.STOP;
    private OperatorType operator = OperatorType.NONE;
    private String formula1;
    private String formula2;
    private boolean ignoreBlank;
    private boolean inCellDropDown = true;
    private String inputTitle;
    private String inputMessage;
    private String errorTitle;
    private String errorMessage;
    private boolean showInput;
    private boolean showError;

    /**
     * Initializes a new ValidationModel instance.
     */
    public ValidationModel() {
        // Default constructor initializes the areas list
    }

    /**
     * Returns the areas.
     * @return the requested result
     */
    public List<CellArea> getAreas() { return areas; }

    /**
     * Returns the type.
     * @return the requested result
     */
    public ValidationType getType() { return type; }
    /**
     * Sets the type.
     * @param type type
     */
    public void setType(ValidationType type) { this.type = type; }

    /**
     * Returns the alert style.
     * @return the requested result
     */
    public ValidationAlertType getAlertStyle() { return alertStyle; }
    /**
     * Sets the alert style.
     * @param alertStyle alert style
     */
    public void setAlertStyle(ValidationAlertType alertStyle) { this.alertStyle = alertStyle; }

    /**
     * Returns the operator.
     * @return the requested result
     */
    public OperatorType getOperator() { return operator; }
    /**
     * Sets the operator.
     * @param operator operator
     */
    public void setOperator(OperatorType operator) { this.operator = operator; }

    /**
     * Returns the formula 1.
     * @return the requested result
     */
    public String getFormula1() { return formula1; }
    /**
     * Sets the formula 1.
     * @param formula1 formula 1
     */
    public void setFormula1(String formula1) { this.formula1 = formula1; }

    /**
     * Returns the formula 2.
     * @return the requested result
     */
    public String getFormula2() { return formula2; }
    /**
     * Sets the formula 2.
     * @param formula2 formula 2
     */
    public void setFormula2(String formula2) { this.formula2 = formula2; }

    /**
     * Returns the ignore blank.
     * @return the requested result
     */
    public boolean getIgnoreBlank() { return ignoreBlank; }
    /**
     * Sets the ignore blank.
     * @param ignoreBlank ignore blank
     */
    public void setIgnoreBlank(boolean ignoreBlank) { this.ignoreBlank = ignoreBlank; }

    /**
     * Returns the in cell drop down.
     * @return the requested result
     */
    public boolean getInCellDropDown() { return inCellDropDown; }
    /**
     * Sets the in cell drop down.
     * @param inCellDropDown in cell drop down
     */
    public void setInCellDropDown(boolean inCellDropDown) { this.inCellDropDown = inCellDropDown; }

    /**
     * Returns the input title.
     * @return the requested result
     */
    public String getInputTitle() { return inputTitle; }
    /**
     * Sets the input title.
     * @param inputTitle input title
     */
    public void setInputTitle(String inputTitle) { this.inputTitle = inputTitle; }

    /**
     * Returns the input message.
     * @return the requested result
     */
    public String getInputMessage() { return inputMessage; }
    /**
     * Sets the input message.
     * @param inputMessage input message
     */
    public void setInputMessage(String inputMessage) { this.inputMessage = inputMessage; }

    /**
     * Returns the error title.
     * @return the requested result
     */
    public String getErrorTitle() { return errorTitle; }
    /**
     * Sets the error title.
     * @param errorTitle error title
     */
    public void setErrorTitle(String errorTitle) { this.errorTitle = errorTitle; }

    /**
     * Returns the error message.
     * @return the requested result
     */
    public String getErrorMessage() { return errorMessage; }
    /**
     * Sets the error message.
     * @param errorMessage error message
     */
    public void setErrorMessage(String errorMessage) { this.errorMessage = errorMessage; }

    /**
     * Returns the show input.
     * @return the requested result
     */
    public boolean getShowInput() { return showInput; }
    /**
     * Sets the show input.
     * @param showInput show input
     */
    public void setShowInput(boolean showInput) { this.showInput = showInput; }

    /**
     * Returns the show error.
     * @return the requested result
     */
    public boolean getShowError() { return showError; }
    /**
     * Sets the show error.
     * @param showError show error
     */
    public void setShowError(boolean showError) { this.showError = showError; }
}
