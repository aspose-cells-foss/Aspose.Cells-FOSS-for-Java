package org.aspose.cells_foss;

/**
 * Represents information about a warning that occurred during workbook operations.
 */
public final class WarningInfo {
    private String code = "";
    private DiagnosticSeverity severity = DiagnosticSeverity.WARNING;
    private String message = "";
    private boolean dataLossRisk;
    private String partUri;
    private String sheetName;
    private String cellRef;
    private Integer rowIndex;

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
     * Returns the data loss risk.
     * @return the requested result
     */
    public boolean getDataLossRisk() { return dataLossRisk; }
    /**
     * Sets the data loss risk.
     * @param dataLossRisk data loss risk
     */
    public void setDataLossRisk(boolean dataLossRisk) { this.dataLossRisk = dataLossRisk; }

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
