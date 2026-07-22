package org.aspose.cells_foss;

/**
 * Represents an issue that occurred during workbook loading.
 */
public final class LoadIssue {
    private final String code;
    private final DiagnosticSeverity severity;
    private final String message;
    private final boolean repairApplied;
    private final boolean dataLossRisk;
    private String partUri;
    private String sheetName;
    private String cellRef;
    private Integer rowIndex;

    /**
     * Initializes a new LoadIssue instance.
     * @param code code
     * @param severity severity
     * @param message message
     * @param repairApplied repair applied
     * @param dataLossRisk data loss risk
     */
    public LoadIssue(String code, DiagnosticSeverity severity, String message, boolean repairApplied, boolean dataLossRisk) {
        // Validate the caller input before continuing.
        if (code == null) throw new IllegalArgumentException("code");
        if (message == null) throw new IllegalArgumentException("message");
        this.code = code;
        this.severity = severity;
        this.message = message;
        this.repairApplied = repairApplied;
        this.dataLossRisk = dataLossRisk;
    }

    /**
     * Returns the code.
     * @return the requested result
     */
    public String getCode() { return code; }
    /**
     * Returns the severity.
     * @return the requested result
     */
    public DiagnosticSeverity getSeverity() { return severity; }
    /**
     * Returns the message.
     * @return the requested result
     */
    public String getMessage() { return message; }
    /**
     * Returns the repair applied.
     * @return the requested result
     */
    public boolean getRepairApplied() { return repairApplied; }
    /**
     * Returns the data loss risk.
     * @return the requested result
     */
    public boolean getDataLossRisk() { return dataLossRisk; }
    /**
     * Returns the part uri.
     * @return the requested result
     */
    public String getPartUri() { return partUri; }
    /**
     * Sets the part uri.
     * @param partUri part uri
     */
    public void setPartUri(String partUri) { this.partUri = partUri; }
    /**
     * Returns the sheet name.
     * @return the requested result
     */
    public String getSheetName() { return sheetName; }
    /**
     * Sets the sheet name.
     * @param sheetName name to use
     */
    public void setSheetName(String sheetName) { this.sheetName = sheetName; }
    /**
     * Returns the cell ref.
     * @return the requested result
     */
    public String getCellRef() { return cellRef; }
    /**
     * Sets the cell ref.
     * @param cellRef cell ref
     */
    public void setCellRef(String cellRef) { this.cellRef = cellRef; }
    /**
     * Returns the row index.
     * @return the requested result
     */
    public Integer getRowIndex() { return rowIndex; }
    /**
     * Sets the row index.
     * @param rowIndex zero-based row index
     */
    public void setRowIndex(Integer rowIndex) { this.rowIndex = rowIndex; }
}
