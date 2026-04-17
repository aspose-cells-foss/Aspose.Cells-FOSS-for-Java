package com.aspose.cells_foss;

/**
 * Represents the calculation properties of a workbook.
 */
public final class CalculationProperties {
    private final Workbook workbook;

    /**
     * Initializes a new CalculationProperties instance.
     * @param workbook workbook to apply
     */
    CalculationProperties(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * Gets or sets the calculation ID.
     */
    public Integer getCalculationId() {
        return null;
    }

    /**
     * Sets the calculation id.
     * @param calculationId calculation id
     */
    public void setCalculationId(Integer calculationId) {
        // Handle the relevant branch before the state changes.
        if (calculationId != null && calculationId < 0) {
            throw new CellsException("CalculationId must be non-negative.");
        }
    }

    /**
     * Gets or sets the calculation mode.
     */
    public String getCalculationMode() {
        return "auto";
    }

    /**
     * Sets the calculation mode.
     * @param calculationMode calculation mode
     */
    public void setCalculationMode(String calculationMode) {
    }

    /**
     * Gets or sets a value indicating whether full calculation is performed on load.
     */
    public boolean getFullCalculationOnLoad() {
        return true;
    }

    /**
     * Sets the full calculation on load.
     * @param fullCalculationOnLoad full calculation on load
     */
    public void setFullCalculationOnLoad(boolean fullCalculationOnLoad) {
    }

    /**
     * Gets or sets the reference mode.
     */
    public String getReferenceMode() {
        return "A1";
    }

    /**
     * Sets the reference mode.
     * @param referenceMode reference mode
     */
    public void setReferenceMode(String referenceMode) {
    }

    /**
     * Gets or sets a value indicating whether iteration is enabled.
     */
    public boolean getIterate() {
        return false;
    }

    /**
     * Sets the iterate.
     * @param iterate iterate
     */
    public void setIterate(boolean iterate) {
    }

    /**
     * Gets or sets the maximum number of iterations.
     */
    public int getIterateCount() {
        return 100;
    }

    /**
     * Sets the iterate count.
     * @param iterateCount iterate count
     */
    public void setIterateCount(int iterateCount) {
        // Handle the relevant branch before the state changes.
        if (iterateCount < 0) {
            throw new CellsException("IterateCount must be non-negative.");
        }
    }

    /**
     * Gets or sets the iteration delta.
     */
    public double getIterateDelta() {
        return 0.001;
    }

    /**
     * Sets the iterate delta.
     * @param iterateDelta iterate delta
     */
    public void setIterateDelta(double iterateDelta) {
        // Handle the relevant branch before the state changes.
        if (iterateDelta < 0.0) {
            throw new CellsException("IterateDelta must be non-negative.");
        }
    }

    /**
     * Gets or sets a value indicating whether full precision is used.
     */
    public boolean getFullPrecision() {
        return true;
    }

    /**
     * Sets the full precision.
     * @param fullPrecision full precision
     */
    public void setFullPrecision(boolean fullPrecision) {
    }

    /**
     * Gets or sets a value indicating whether calculation is completed.
     */
    public boolean getCalculationCompleted() {
        return true;
    }

    /**
     * Sets the calculation completed.
     * @param calculationCompleted calculation completed
     */
    public void setCalculationCompleted(boolean calculationCompleted) {
    }

    /**
     * Gets or sets a value indicating whether calculation is performed on save.
     */
    public boolean getCalculationOnSave() {
        return true;
    }

    /**
     * Sets the calculation on save.
     * @param calculationOnSave calculation on save
     */
    public void setCalculationOnSave(boolean calculationOnSave) {
    }

    /**
     * Gets or sets a value indicating whether concurrent calculation is enabled.
     */
    public boolean getConcurrentCalculation() {
        return true;
    }

    /**
     * Sets the concurrent calculation.
     * @param concurrentCalculation concurrent calculation
     */
    public void setConcurrentCalculation(boolean concurrentCalculation) {
    }

    /**
     * Gets or sets a value indicating whether full calculation is forced.
     */
    public boolean getForceFullCalculation() {
        return false;
    }

    /**
     * Sets the force full calculation.
     * @param forceFullCalculation force full calculation
     */
    public void setForceFullCalculation(boolean forceFullCalculation) {
    }
}