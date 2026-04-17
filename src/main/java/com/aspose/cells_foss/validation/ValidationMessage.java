package com.aspose.cells_foss.validation;

/**
 * Represents a validation message with code, severity, and message text.
 */
public final class ValidationMessage {
    private String code = "";
    private ValidationMessageSeverity severity;
    private String message = "";

    /**
     * Returns the code.
     * @return the requested result
     */
    public String getCode() { return code; }
    /**
     * Sets the code.
     * @param code code
     */
    public void setCode(String code) { this.code = code; }

    /**
     * Returns the severity.
     * @return the requested result
     */
    public ValidationMessageSeverity getSeverity() { return severity; }
    /**
     * Sets the severity.
     * @param severity severity
     */
    public void setSeverity(ValidationMessageSeverity severity) { this.severity = severity; }

    /**
     * Returns the message.
     * @return the requested result
     */
    public String getMessage() { return message; }
    /**
     * Sets the message.
     * @param message message
     */
    public void setMessage(String message) { this.message = message; }
}