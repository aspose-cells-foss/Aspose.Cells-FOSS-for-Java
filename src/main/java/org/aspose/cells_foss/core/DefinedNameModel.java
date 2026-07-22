package org.aspose.cells_foss.core;

/**
 * Represents a defined name model in the Excel file.
 */
public final class DefinedNameModel {
    private String name = "";
    private String formula = "";
    private Integer localSheetIndex;
    private boolean hidden;
    private String comment = "";

    /**
     * Returns the name.
     * @return the requested result
     */
    public String getName() { return name; }
    /**
     * Sets the name.
     * @param name name
     */
    public void setName(String name) { this.name = name; }

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
     * Returns the local sheet index.
     * @return the requested result
     */
    public Integer getLocalSheetIndex() { return localSheetIndex; }
    /**
     * Sets the local sheet index.
     * @param localSheetIndex zero-based local sheet index
     */
    public void setLocalSheetIndex(Integer localSheetIndex) { this.localSheetIndex = localSheetIndex; }

    /**
     * Returns the hidden.
     * @return the requested result
     */
    public boolean getHidden() { return hidden; }
    /**
     * Sets the hidden.
     * @param hidden hidden
     */
    public void setHidden(boolean hidden) { this.hidden = hidden; }

    /**
     * Returns the comment.
     * @return the requested result
     */
    public String getComment() { return comment; }
    /**
     * Sets the comment.
     * @param comment comment
     */
    public void setComment(String comment) { this.comment = comment; }
}
