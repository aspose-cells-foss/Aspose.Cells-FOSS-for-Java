package org.aspose.cells_foss.core;

/**
 * Represents a hyperlink model with its properties.
 */
public final class HyperlinkModel {
    private int firstRow;
    private int firstColumn;
    private int totalRows = 1;
    private int totalColumns = 1;
    private String address;
    private String subAddress;
    private String screenTip;
    private String textToDisplay;

    /**
     * Returns the first row.
     * @return the requested result
     */
    public int getFirstRow() { return firstRow; }
    /**
     * Sets the first row.
     * @param firstRow first row
     */
    public void setFirstRow(int firstRow) { this.firstRow = firstRow; }

    /**
     * Returns the first column.
     * @return the requested result
     */
    public int getFirstColumn() { return firstColumn; }
    /**
     * Sets the first column.
     * @param firstColumn first column
     */
    public void setFirstColumn(int firstColumn) { this.firstColumn = firstColumn; }

    /**
     * Returns the total rows.
     * @return the requested result
     */
    public int getTotalRows() { return totalRows; }
    /**
     * Sets the total rows.
     * @param totalRows total rows
     */
    public void setTotalRows(int totalRows) { this.totalRows = totalRows; }

    /**
     * Returns the total columns.
     * @return the requested result
     */
    public int getTotalColumns() { return totalColumns; }
    /**
     * Sets the total columns.
     * @param totalColumns total columns
     */
    public void setTotalColumns(int totalColumns) { this.totalColumns = totalColumns; }

    /**
     * Returns the address.
     * @return the requested result
     */
    public String getAddress() { return address; }
    /**
     * Sets the address.
     * @param address address
     */
    public void setAddress(String address) { this.address = address; }

    /**
     * Returns the sub address.
     * @return the requested result
     */
    public String getSubAddress() { return subAddress; }
    /**
     * Sets the sub address.
     * @param subAddress sub address
     */
    public void setSubAddress(String subAddress) { this.subAddress = subAddress; }

    /**
     * Returns the screen tip.
     * @return the requested result
     */
    public String getScreenTip() { return screenTip; }
    /**
     * Sets the screen tip.
     * @param screenTip screen tip
     */
    public void setScreenTip(String screenTip) { this.screenTip = screenTip; }

    /**
     * Returns the text to display.
     * @return the requested result
     */
    public String getTextToDisplay() { return textToDisplay; }
    /**
     * Sets the text to display.
     * @param textToDisplay text to display
     */
    public void setTextToDisplay(String textToDisplay) { this.textToDisplay = textToDisplay; }
}
