package com.aspose.cells_foss;

/**
 * Represents an exception that occurs when loading a workbook.
 */
public class WorkbookLoadException extends CellsException {
    /**
     * Initializes a new WorkbookLoadException instance.
     * @param message message
     */
    public WorkbookLoadException(String message) {
        super(message);
    }

    /**
     * Initializes a new WorkbookLoadException instance.
     * @param message message
     * @param innerException inner exception
     */
    public WorkbookLoadException(String message, Exception innerException) {
        super(message, innerException);
    }
}