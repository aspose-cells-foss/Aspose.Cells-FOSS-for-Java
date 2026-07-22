package org.aspose.cells_foss;

/**
 * Represents an exception thrown by the Aspose.Cells library.
 * This class extends RuntimeException to maintain consistency with the project's
 * exception handling model, where all custom exceptions are unchecked.
 */
public class CellsException extends RuntimeException {
    /**
     * Initializes a new CellsException instance.
     * @param message message
     */
    public CellsException(String message) {
        super(message);
    }

    /**
     * Initializes a new CellsException instance.
     * @param message message
     * @param innerException inner exception
     */
    public CellsException(String message, Throwable innerException) {
        super(message, innerException);
    }
}
