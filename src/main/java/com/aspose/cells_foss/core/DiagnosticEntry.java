package com.aspose.cells_foss.core;

/**
 * Represents a diagnostic entry with details about a problem or warning.
 */
public final class DiagnosticEntry {
    private String code = "";
    private DiagnosticSeverity severity = DiagnosticSeverity.WARNING;
    private String message = "";
    private boolean repairApplied;
    private boolean dataLossRisk;

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
    public DiagnosticSeverity getSeverity() { return severity; }
    /**
     * Sets the severity.
     * @param severity severity
     */
    public void setSeverity(DiagnosticSeverity severity) { this.severity = severity; }

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

    /**
     * Returns the repair applied.
     * @return the requested result
     */
    public boolean getRepairApplied() { return repairApplied; }
    /**
     * Sets the repair applied.
     * @param repairApplied repair applied
     */
    public void setRepairApplied(boolean repairApplied) { this.repairApplied = repairApplied; }

    /**
     * Returns the data loss risk.
     * @return the requested result
     */
    public boolean getDataLossRisk() { return dataLossRisk; }
    /**
     * Sets the data loss risk.
     * @param dataLossRisk data loss risk
     */
    public void setDataLossRisk(boolean dataLossRisk) { this.dataLossRisk = dataLossRisk; }
}