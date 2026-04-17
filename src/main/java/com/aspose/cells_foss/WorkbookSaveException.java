package com.aspose.cells_foss;

/**
 * Represents an exception that occurs when saving a workbook.
 */
public class WorkbookSaveException extends CellsException {
    /**
     * Initializes a new WorkbookSaveException instance.
     * @param message message
     */
    public WorkbookSaveException(String message) {
        super(message);
    }

    /**
     * Initializes a new WorkbookSaveException instance.
     * @param message message
     * @param innerException inner exception
     */
    public WorkbookSaveException(String message, Exception innerException) {
        super(message, innerException);
    }
}