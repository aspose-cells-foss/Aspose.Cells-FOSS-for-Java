package com.aspose.cells_foss.core;

/**
 * Represents a range of columns with formatting properties.
 */
public final class ColumnRangeModel {
    private int minColumnIndex;
    private int maxColumnIndex;
    private Double width;
    private boolean hidden;
    private Integer styleIndex;
    private int outlineLevel;   // 0 = ungrouped, 1–7 = nesting depth
    private boolean collapsed;  // true = group is collapsed
    private boolean bestFit;

    /**
     * Returns the min column index.
     * @return the requested result
     */
    public int getMinColumnIndex() { return minColumnIndex; }
    /**
     * Sets the min column index.
     * @param minColumnIndex zero-based min column index
     */
    public void setMinColumnIndex(int minColumnIndex) { this.minColumnIndex = minColumnIndex; }

    /**
     * Returns the max column index.
     * @return the requested result
     */
    public int getMaxColumnIndex() { return maxColumnIndex; }
    /**
     * Sets the max column index.
     * @param maxColumnIndex zero-based max column index
     */
    public void setMaxColumnIndex(int maxColumnIndex) { this.maxColumnIndex = maxColumnIndex; }

    /**
     * Returns the width.
     * @return the requested result
     */
    public Double getWidth() { return width; }
    /**
     * Sets the width.
     * @param width width
     */
    public void setWidth(Double width) { this.width = width; }

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

    public boolean getBestFit() { return bestFit; }
    public void setBestFit(boolean bestFit) { this.bestFit = bestFit; }
}