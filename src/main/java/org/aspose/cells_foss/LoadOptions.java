package org.aspose.cells_foss;

/**
 * Represents options for loading a workbook.
 */
public final class LoadOptions {
    private LoadFormat loadFormat = LoadFormat.AUTO;
    private boolean strictMode;
    private boolean tryRepairPackage = true;
    private boolean tryRepairXml = true;
    private boolean preserveUnsupportedParts = true;
    private IWarningCallback warningCallback;

    /**
     * Returns the load format.
     * @return the requested result
     */
    public LoadFormat getLoadFormat() { return loadFormat; }
    /**
     * Sets the load format.
     * @param loadFormat load format
     */
    public void setLoadFormat(LoadFormat loadFormat) { this.loadFormat = loadFormat; }

    /**
     * Indicates whether strict mode.
     * @return true when the condition is satisfied
     */
    public boolean isStrictMode() { return strictMode; }
    /**
     * Sets the strict mode.
     * @param strictMode strict mode
     */
    public void setStrictMode(boolean strictMode) { this.strictMode = strictMode; }

    /**
     * Indicates whether try repair package.
     * @return true when the condition is satisfied
     */
    public boolean isTryRepairPackage() { return tryRepairPackage; }
    /**
     * Sets the try repair package.
     * @param tryRepairPackage try repair package
     */
    public void setTryRepairPackage(boolean tryRepairPackage) { this.tryRepairPackage = tryRepairPackage; }

    /**
     * Indicates whether try repair xml.
     * @return true when the condition is satisfied
     */
    public boolean isTryRepairXml() { return tryRepairXml; }
    /**
     * Sets the try repair xml.
     * @param tryRepairXml try repair xml
     */
    public void setTryRepairXml(boolean tryRepairXml) { this.tryRepairXml = tryRepairXml; }

    /**
     * Indicates whether preserve unsupported parts.
     * @return true when the condition is satisfied
     */
    public boolean isPreserveUnsupportedParts() { return preserveUnsupportedParts; }
    /**
     * Sets the preserve unsupported parts.
     * @param preserveUnsupportedParts preserve unsupported parts
     */
    public void setPreserveUnsupportedParts(boolean preserveUnsupportedParts) { this.preserveUnsupportedParts = preserveUnsupportedParts; }

    /**
     * Returns the warning callback.
     * @return the requested result
     */
    public IWarningCallback getWarningCallback() { return warningCallback; }
    /**
     * Sets the warning callback.
     * @param warningCallback warning callback
     */
    public void setWarningCallback(IWarningCallback warningCallback) { this.warningCallback = warningCallback; }
}
