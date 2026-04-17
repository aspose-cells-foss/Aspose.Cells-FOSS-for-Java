package com.aspose.cells_foss;

/**
 * Holds metadata about a single section of a number format code string.
 */
public final class DisplayFormatSectionInfo {
    private String raw;
    private boolean hasCondition;
    private String conditionOperator = "";
    private double conditionValue;

    /**
     * Returns the raw.
     * @return the requested result
     */
    public String getRaw() { return raw; }
    /**
     * Sets the raw.
     * @param raw raw
     */
    public void setRaw(String raw) { this.raw = raw; }

    /**
     * Returns the condition.
     * @return the requested result
     */
    public boolean getHasCondition() { return hasCondition; }
    /**
     * Sets the condition.
     * @param hasCondition has condition
     */
    public void setHasCondition(boolean hasCondition) { this.hasCondition = hasCondition; }

    /**
     * Returns the condition operator.
     * @return the requested result
     */
    public String getConditionOperator() { return conditionOperator; }
    /**
     * Sets the condition operator.
     * @param conditionOperator condition operator
     */
    public void setConditionOperator(String conditionOperator) { this.conditionOperator = conditionOperator; }

    /**
     * Returns the condition value.
     * @return the requested result
     */
    public double getConditionValue() { return conditionValue; }
    /**
     * Sets the condition value.
     * @param conditionValue condition value
     */
    public void setConditionValue(double conditionValue) { this.conditionValue = conditionValue; }
}
