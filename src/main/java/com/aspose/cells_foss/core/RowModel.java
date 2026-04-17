package com.aspose.cells_foss.core;

/**
 * Represents a row model in the Excel file.
 */
public final class RowModel {
    private Double height;
    private boolean hidden;
    private Integer styleIndex;
    private int outlineLevel;   // 0 = ungrouped, 1–7 = nesting depth
    private boolean collapsed;  // true = group is collapsed

    /**
     * Returns the height.
     * @return the requested result
     */
    public Double getHeight() { return height; }
    /**
     * Sets the height.
     * @param height height
     */
    public void setHeight(Double height) { this.height = height; }

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
     * Returns the style index.
     * @return the requested result
     */
    public Integer getStyleIndex() { return styleIndex; }
    /**
     * Sets the style index.
     * @param styleIndex zero-based style index
     */
    public void setStyleIndex(Integer styleIndex) { this.styleIndex = styleIndex; }

    /**
     * Returns the outline level.
     * @return the requested result
     */
    public int getOutlineLevel() { return outlineLevel; }
    /**
     * Sets the outline level.
     * @param outlineLevel outline level
     */
    public void setOutlineLevel(int outlineLevel) { this.outlineLevel = outlineLevel; }

    /**
     * Returns the collapsed.
     * @return the requested result
     */
    public boolean getCollapsed() { return collapsed; }
    /**
     * Sets the collapsed.
     * @param collapsed collapsed
     */
    public void setCollapsed(boolean collapsed) { this.collapsed = collapsed; }
}