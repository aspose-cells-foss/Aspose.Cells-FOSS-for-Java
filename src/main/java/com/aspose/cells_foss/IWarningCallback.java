package com.aspose.cells_foss;

/**
 * Provides a callback mechanism for reporting warnings during workbook operations.
 */
public interface IWarningCallback {
    /**
     * Called when a warning occurs.
     *
     * @param warningInfo information about the warning
     */
    void warning(WarningInfo warningInfo);
}