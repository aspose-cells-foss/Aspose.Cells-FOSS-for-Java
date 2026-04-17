package com.aspose.cells_foss;

/**
 * Represents save options for workbook saving.
 */
public final class SaveOptions {
    private SaveFormat saveFormat = SaveFormat.XLSX;
    private boolean useSharedStrings = true;
    private boolean validateBeforeSave = true;
    private boolean compactStyles = true;
    private boolean preserveRecoveryMetadata;

    /**
     * Returns the save format.
     * @return the requested result
     */
    public SaveFormat getSaveFormat() { return saveFormat; }
    /**
     * Sets the save format.
     * @param saveFormat save format
     */
    public void setSaveFormat(SaveFormat saveFormat) { this.saveFormat = saveFormat; }

    /**
     * Returns the use shared strings.
     * @return the requested result
     */
    public boolean getUseSharedStrings() { return useSharedStrings; }
    /**
     * Sets the use shared strings.
     * @param useSharedStrings use shared strings
     */
    public void setUseSharedStrings(boolean useSharedStrings) { this.useSharedStrings = useSharedStrings; }

    /**
     * Returns the validate before save.
     * @return the requested result
     */
    public boolean getValidateBeforeSave() { return validateBeforeSave; }
    /**
     * Sets the validate before save.
     * @param validateBeforeSave validate before save
     */
    public void setValidateBeforeSave(boolean validateBeforeSave) { this.validateBeforeSave = validateBeforeSave; }

    /**
     * Returns the compact styles.
     * @return the requested result
     */
    public boolean getCompactStyles() { return compactStyles; }
    /**
     * Sets the compact styles.
     * @param compactStyles compact styles
     */
    public void setCompactStyles(boolean compactStyles) { this.compactStyles = compactStyles; }

    /**
     * Returns the preserve recovery metadata.
     * @return the requested result
     */
    public boolean getPreserveRecoveryMetadata() { return preserveRecoveryMetadata; }
    /**
     * Sets the preserve recovery metadata.
     * @param preserveRecoveryMetadata preserve recovery metadata
     */
    public void setPreserveRecoveryMetadata(boolean preserveRecoveryMetadata) { this.preserveRecoveryMetadata = preserveRecoveryMetadata; }
}