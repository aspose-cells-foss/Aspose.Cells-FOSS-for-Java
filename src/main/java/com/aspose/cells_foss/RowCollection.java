package com.aspose.cells_foss;

/**
 * Represents a collection of rows in a worksheet.
 */
public final class RowCollection {
    private final Worksheet worksheet;

    /**
     * Initializes a new RowCollection instance.
     * @param worksheet worksheet to apply
     */
    RowCollection(Worksheet worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Returns the requested item.
     * @param index index
     * @return the requested result
     */
    public Row get(int index) {
        // Validate the caller input before continuing.
        if (index < 0) throw new CellsException("Row index must be non-negative.");
        return new Row(worksheet, index);
    }
}