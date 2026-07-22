package org.aspose.cells_foss;

/**
 * Represents an exception thrown when an invalid file format is encountered.
 */
public class InvalidFileFormatException extends CellsException {
    /**
     * Initializes a new InvalidFileFormatException instance.
     * @param message message
     */
    public InvalidFileFormatException(String message) {
        super(message);
    }

    /**
     * Initializes a new InvalidFileFormatException instance.
     * @param message message
     * @param innerException inner exception
     */
    public InvalidFileFormatException(String message, Throwable innerException) {
        super(message, innerException);
    }
}
