package com.aspose.cells_foss.core;

/**
 * Represents a cell record in the Excel file.
 */
public final class CellRecord {
    private Object value;
    private CellValueKind kind;
    private String formula;
    private StyleValue style;
    private boolean isExplicitlyStored;

    /**
     * Initializes a new CellRecord instance.
     */
    public CellRecord() {
        this.kind = CellValueKind.BLANK;
        this.style = StyleValue.getDefault();
    }

    /**
     * Returns the value.
     * @return the requested result
     */
    public Object getValue() { return value; }
    /**
     * Sets the value.
     * @param value value to apply
     */
    public void setValue(Object value) { this.value = value; }

    /**
     * Returns the kind.
     * @return the requested result
     */
    public CellValueKind getKind() { return kind; }
    /**
     * Sets the kind.
     * @param kind kind
     */
    public void setKind(CellValueKind kind) { this.kind = kind; }

    /**
     * Returns the formula.
     * @return the requested result
     */
    public String getFormula() { return formula; }
    /**
     * Sets the formula.
     * @param formula formula
     */
    public void setFormula(String formula) { this.formula = formula; }

    /**
     * Returns the style.
     * @return the requested result
     */
    public StyleValue getStyle() { return style; }
    /**
     * Sets the style.
     * @param style style to apply
     */
    public void setStyle(StyleValue style) { this.style = style; }

    /**
     * Returns the explicitly stored.
     * @return the requested result
     */
    public boolean getIsExplicitlyStored() { return isExplicitlyStored; }
    /**
     * Sets the explicitly stored.
     * @param isExplicitlyStored is explicitly stored
     */
    public void setIsExplicitlyStored(boolean isExplicitlyStored) { this.isExplicitlyStored = isExplicitlyStored; }
}