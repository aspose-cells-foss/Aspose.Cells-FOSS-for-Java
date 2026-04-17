package com.aspose.cells_foss.core;

/**
 * Represents a merge region in an Excel worksheet.
 */
public final class MergeRegion {
    private final int firstRow;
    private final int firstColumn;
    private final int totalRows;
    private final int totalColumns;

    /**
     * Initializes a new MergeRegion instance.
     * @param firstRow first row
     * @param firstColumn first column
     * @param totalRows total rows
     * @param totalColumns total columns
     */
    public MergeRegion(int firstRow, int firstColumn, int totalRows, int totalColumns) {
        this.firstRow = firstRow;
        this.firstColumn = firstColumn;
        this.totalRows = totalRows;
        this.totalColumns = totalColumns;
    }

    /**
     * Returns the first row.
     * @return the requested result
     */
    public int getFirstRow() { return firstRow; }
    /**
     * Returns the first column.
     * @return the requested result
     */
    public int getFirstColumn() { return firstColumn; }
    /**
     * Returns the total rows.
     * @return the requested result
     */
    public int getTotalRows() { return totalRows; }
    /**
     * Returns the total columns.
     * @return the requested result
     */
    public int getTotalColumns() { return totalColumns; }
}