package com.aspose.cells_foss.core;


/**
 * Represents a cell address with row and column indices.
 */
public final class CellAddress {
    private final int rowIndex;
    private final int columnIndex;

    /**
     * Creates a new CellAddress with the specified row and column indices.
     * 
     * @param rowIndex the zero-based row index
     * @param columnIndex the zero-based column index
     * @throws IllegalArgumentException if rowIndex or columnIndex is negative
     */
    public CellAddress(int rowIndex, int columnIndex) {
        // Validate the caller input before continuing.
        if (rowIndex < 0) throw new IllegalArgumentException("rowIndex must be non-negative");
        if (columnIndex < 0) throw new IllegalArgumentException("columnIndex must be non-negative");
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
    }

    /**
     * Gets the zero-based row index.
     * 
     * @return the row index
     */
    public int getRowIndex() {
        return rowIndex;
    }

    /**
     * Gets the zero-based column index.
     * 
     * @return the column index
     */
    public int getColumnIndex() {
        return columnIndex;
    }

    /**
     * Parses a cell reference string (e.g., "A1", "Z100") into a CellAddress.
     * 
     * @param cellReference the cell reference string to parse
     * @return a CellAddress representing the parsed reference
     * @throws IllegalArgumentException if the cell reference is null, empty, or invalid
     */
    public static CellAddress parse(String cellReference) {
        // Handle the relevant branch before the state changes.
        if (cellReference == null || cellReference.isBlank())
            throw new IllegalArgumentException("Cell reference must be non-empty.");
        String reference = cellReference.trim();
        int index = 0;
        int column = 0;

        while (index < reference.length() && Character.isLetter(reference.charAt(index))) {
            char letter = Character.toUpperCase(reference.charAt(index));
            if (letter < 'A' || letter > 'Z')
                throw new IllegalArgumentException("Cell reference '" + cellReference + "' is invalid.");
            column = (column * 26) + (letter - 'A' + 1);
            index++;
        }

        if (column == 0 || index == reference.length())
            throw new IllegalArgumentException("Cell reference '" + cellReference + "' is invalid.");

        int row = 0;
        while (index < reference.length() && Character.isDigit(reference.charAt(index))) {
            row = (row * 10) + (reference.charAt(index) - '0');
            index++;
        }

        if (index != reference.length() || row <= 0)
            throw new IllegalArgumentException("Cell reference '" + cellReference + "' is invalid.");
        return new CellAddress(row - 1, column - 1);
    }

    /**
     * Compares this instance with the provided value.
     * @param obj obj
     * @return true when the condition is satisfied
     */
    @Override
    public boolean equals(Object obj) {
        // Handle the relevant branch before the state changes.
        if (this == obj) return true;
        if (obj == null || getClass() != obj.getClass()) return false;
        CellAddress other = (CellAddress) obj;
        return rowIndex == other.rowIndex && columnIndex == other.columnIndex;
    }

    /**
     * Indicates whether this instance has h code.
     * @return true when the condition is satisfied
     */
    @Override
    public int hashCode() {
        return (rowIndex * 397) ^ columnIndex;
    }

    /**
     * Returns the string representation of this instance.
     * @return the computed result
     */
    @Override
    public String toString() {
        return columnIndexToName(columnIndex) + (rowIndex + 1);
    }

    /**
     * Processes column index to name.
     * @param columnIndex zero-based column index
     * @return the computed result
     */
    private static String columnIndexToName(int columnIndex) {
        int index = columnIndex + 1;
        StringBuilder result = new StringBuilder();
        while (index > 0) {
            index--;
            result.append((char) ('A' + (index % 26)));
            index /= 26;
        }
        return result.reverse().toString();
    }
}