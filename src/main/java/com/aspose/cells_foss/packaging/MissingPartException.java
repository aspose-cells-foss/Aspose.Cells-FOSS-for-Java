package com.aspose.cells_foss.packaging;

/**
 * Thrown when a required part is missing from the package structure.
 */
public class MissingPartException extends PackageStructureException {
    /**
     * Initializes a new MissingPartException instance.
     * @param message message
     */
    public MissingPartException(String message) {
        super(message);
    }
}