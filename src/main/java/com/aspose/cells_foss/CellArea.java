package com.aspose.cells_foss;

import com.aspose.cells_foss.core.CellAddress;

/**
 * Represents a cell area with row and column bounds.
 */
public final class CellArea {
    private final int firstRow;
    private final int firstColumn;
    private final int totalRows;
    private final int totalColumns;

    /**
     * Creates a new CellArea with the specified bounds.
     *
     * @param firstRow      the first row index (zero-based)
     * @param firstColumn   the first column index (zero-based)
     * @param totalRows     the total number of rows (must be > 0)
     * @param totalColumns  the total number of columns (must be > 0)
     */
    public CellArea(int firstRow, int firstColumn, int totalRows, int totalColumns) {
        // Validate the caller input before continuing.
        if (firstRow < 0) throw new IllegalArgumentException("firstRow must be non-negative");
        if (firstColumn < 0) throw new IllegalArgumentException("firstColumn must be non-negative");
        if (totalRows <= 0) throw new IllegalArgumentException("totalRows must be positive");
        if (totalColumns <= 0) throw new IllegalArgumentException("totalColumns must be positive");

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

    /**
     * Creates a CellArea from start and end cell coordinates.
     *
     * @param startRow     the start row index (zero-based)
     * @param startColumn  the start column index (zero-based)
     * @param endRow       the end row index (zero-based)
     * @param endColumn    the end column index (zero-based)
     * @return a new CellArea
     */
    public static CellArea createCellArea(int startRow, int startColumn, int endRow, int endColumn) {
        // Validate the caller input before continuing.
        if (endRow < startRow) throw new IllegalArgumentException("endRow must be >= startRow");
        if (endColumn < startColumn) throw new IllegalArgumentException("endColumn must be >= startColumn");
        return new CellArea(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
    }

    /**
     * Creates a CellArea from start and end cell names (e.g., "A1", "B2").
     *
     * @param startCellName  the start cell name
     * @param endCellName    the end cell name
     * @return a new CellArea
     */
    public static CellArea createCellArea(String startCellName, String endCellName) {
        CellAddress start = CellAddress.parse(startCellName);
        CellAddress end = CellAddress.parse(endCellName);
        return createCellArea(
            Math.min(start.getRowIndex(), end.getRowIndex()),
            Math.min(start.getColumnIndex(), end.getColumnIndex()),
            Math.max(start.getRowIndex(), end.getRowIndex()),
            Math.max(start.getColumnIndex(), end.getColumnIndex()));
    }
}