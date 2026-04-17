package com.aspose.cells_foss.core;

/**
 * Represents print options for a worksheet.
 */
public final class PrintOptionsModel {
    private boolean gridLines;
    private boolean headings;
    private boolean horizontalCentered;
    private boolean verticalCentered;

    /**
     * Returns the grid lines.
     * @return the requested result
     */
    public boolean getGridLines() { return gridLines; }
    /**
     * Sets the grid lines.
     * @param gridLines grid lines
     */
    public void setGridLines(boolean gridLines) { this.gridLines = gridLines; }

    /**
     * Returns the headings.
     * @return the requested result
     */
    public boolean getHeadings() { return headings; }
    /**
     * Sets the headings.
     * @param headings headings
     */
    public void setHeadings(boolean headings) { this.headings = headings; }

    /**
     * Returns the horizontal centered.
     * @return the requested result
     */
    public boolean getHorizontalCentered() { return horizontalCentered; }
    /**
     * Sets the horizontal centered.
     * @param horizontalCentered horizontal centered
     */
    public void setHorizontalCentered(boolean horizontalCentered) { this.horizontalCentered = horizontalCentered; }

    /**
     * Returns the vertical centered.
     * @return the requested result
     */
    public boolean getVerticalCentered() { return verticalCentered; }
    /**
     * Sets the vertical centered.
     * @param verticalCentered vertical centered
     */
    public void setVerticalCentered(boolean verticalCentered) { this.verticalCentered = verticalCentered; }
}