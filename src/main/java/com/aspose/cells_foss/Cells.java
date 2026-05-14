package com.aspose.cells_foss;

import com.aspose.cells_foss.core.MergeRegion;

/**
 * Represents a collection of cells in a worksheet.
 */
public class Cells {
    private final Worksheet worksheet;

    /**
     * Initializes a new Cells instance.
     * @param worksheet worksheet to apply
     */
    Cells(Worksheet worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Returns the requested item.
     * @param cellName name to use
     * @return the requested result
     */
    public Cell get(String cellName) {
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            return new Cell(worksheet, com.aspose.cells_foss.core.CellAddress.parse(cellName));
        } catch (IllegalArgumentException exception) {
            throw new CellsException(exception.getMessage(), exception);
        }
    }

    /**
     * Returns the requested item.
     * @param row row
     * @param column column
     * @return the requested result
     */
    public Cell get(int row, int column) {
        // Handle the relevant branch before the state changes.
        if (row < 0 || column < 0) {
            throw new CellsException("Row and column indices must be non-negative.");
        }
        return new Cell(worksheet, new com.aspose.cells_foss.core.CellAddress(row, column));
    }

    /**
     * Returns the rows.
     * @return the requested result
     */
    public RowCollection getRows() {
        return new RowCollection(worksheet);
    }

    /**
     * Returns the columns.
     * @return the requested result
     */
    public ColumnCollection getColumns() {
        return new ColumnCollection(worksheet);
    }

    /**
     * Returns the merged cells.
     * @return the requested result
     */
    public java.util.List<CellArea> getMergedCells() {
        java.util.List<CellArea> mergedCells = new java.util.ArrayList<>();
        // Walk the current collection so every entry is processed consistently.
        for (MergeRegion region : worksheet.getModel().getMergeRegions()) {
            mergedCells.add(new CellArea(region.getFirstRow(), region.getFirstColumn(),
                    region.getTotalRows(), region.getTotalColumns()));
        }
        return mergedCells;
    }

    /**
     * Merges the requested cell range.
     * @param firstRow first row
     * @param firstColumn first column
     * @param totalRows total rows
     * @param totalColumns total columns
     */
    public void merge(int firstRow, int firstColumn, int totalRows, int totalColumns) {
        if (firstRow < 0 || firstColumn < 0) {
            throw new CellsException("Merge origin indices must be non-negative.");
        }
        if (totalRows <= 0 || totalColumns <= 0) {
            throw new CellsException("Merge range dimensions must be positive.");
        }
        int lastRow = firstRow + totalRows - 1;
        int lastCol = firstColumn + totalColumns - 1;
        for (MergeRegion r : worksheet.getModel().getMergeRegions()) {
            int rLastRow = r.getFirstRow() + r.getTotalRows() - 1;
            int rLastCol = r.getFirstColumn() + r.getTotalColumns() - 1;
            if (firstRow <= rLastRow && lastRow >= r.getFirstRow()
                    && firstColumn <= rLastCol && lastCol >= r.getFirstColumn()) {
                throw new CellsException("Merge range overlaps with an existing merged region.");
            }
        }
        worksheet.getModel().getMergeRegions().add(
                new MergeRegion(firstRow, firstColumn, totalRows, totalColumns));
    }
}