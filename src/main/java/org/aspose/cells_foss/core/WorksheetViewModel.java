package org.aspose.cells_foss.core;

/**
 * Represents a view model for a worksheet with display settings.
 */
public final class WorksheetViewModel {
    private boolean showGridLines = true;
    private boolean showRowColumnHeaders = true;
    private boolean showZeros = true;
    private boolean rightToLeft;
    private int zoomScale = 100;
    /** Number of rows frozen from the top (ySplit). Zero means no row freeze. */
    private int freezeRow;
    /** Number of columns frozen from the left (xSplit). Zero means no column freeze. */
    private int freezeColumn;

    /**
     * Returns the show grid lines.
     * @return the requested result
     */
    public boolean getShowGridLines() { return showGridLines; }
    /**
     * Sets the show grid lines.
     * @param showGridLines show grid lines
     */
    public void setShowGridLines(boolean showGridLines) { this.showGridLines = showGridLines; }

    /**
     * Returns the show row column headers.
     * @return the requested result
     */
    public boolean getShowRowColumnHeaders() { return showRowColumnHeaders; }
    /**
     * Sets the show row column headers.
     * @param showRowColumnHeaders show row column headers
     */
    public void setShowRowColumnHeaders(boolean showRowColumnHeaders) { this.showRowColumnHeaders = showRowColumnHeaders; }

    /**
     * Returns the show zeros.
     * @return the requested result
     */
    public boolean getShowZeros() { return showZeros; }
    /**
     * Sets the show zeros.
     * @param showZeros show zeros
     */
    public void setShowZeros(boolean showZeros) { this.showZeros = showZeros; }

    /**
     * Returns the right to left.
     * @return the requested result
     */
    public boolean getRightToLeft() { return rightToLeft; }
    /**
     * Sets the right to left.
     * @param rightToLeft right to left
     */
    public void setRightToLeft(boolean rightToLeft) { this.rightToLeft = rightToLeft; }

    /**
     * Returns the zoom scale.
     * @return the requested result
     */
    public int getZoomScale() { return zoomScale; }
    /**
     * Sets the zoom scale.
     * @param zoomScale zoom scale
     */
    public void setZoomScale(int zoomScale) { this.zoomScale = zoomScale; }

    /**
     * Returns the number of rows frozen from the top (ySplit). Zero means no row freeze.
     * @return the requested result
     */
    public int getFreezeRow() { return freezeRow; }
    /**
     * Sets the number of rows frozen from the top.
     * @param freezeRow freeze row count
     */
    public void setFreezeRow(int freezeRow) { this.freezeRow = freezeRow; }

    /**
     * Returns the number of columns frozen from the left (xSplit). Zero means no column freeze.
     * @return the requested result
     */
    public int getFreezeColumn() { return freezeColumn; }
    /**
     * Sets the number of columns frozen from the left.
     * @param freezeColumn freeze column count
     */
    public void setFreezeColumn(int freezeColumn) { this.freezeColumn = freezeColumn; }
}
