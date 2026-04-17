package com.aspose.cells_foss;

/**
 * Represents a collection of columns in a worksheet.
 */
public final class ColumnCollection {

    private final Worksheet worksheet;

    /**
     * Initializes a new ColumnCollection instance.
     * @param worksheet worksheet to apply
     */
    ColumnCollection(Worksheet worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Gets the column at the specified index.
     *
     * @param index the zero-based column index
     * @return the column at the specified index
     * @throws CellsException if the index is negative
     */
    public Column get(int index) {
        // Validate the caller input before continuing.
        if (index < 0) throw new CellsException("Column index must be non-negative.");
        return new Column(worksheet, index);
    }
}