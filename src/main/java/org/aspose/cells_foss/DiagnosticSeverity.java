package org.aspose.cells_foss;

/**
 * Represents the severity level of a diagnostic message.
 */
public enum DiagnosticSeverity {
    /** A warning that does not prevent processing. */
    WARNING,
    /** A recoverable issue that can be handled. */
    RECOVERABLE,
    /** A lossy recoverable issue where some data may be lost. */
    LOSSY_RECOVERABLE,
    /** A fatal issue that stops processing. */
    FATAL
}
