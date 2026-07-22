package org.aspose.cells_foss.core;

/**
 * Represents a number format value with its number format index and custom format string.
 */
public final class NumberFormatValue {
    private int number;
    private String custom;

    /**
     * Returns the number.
     * @return the requested result
     */
    public int getNumber() { return number; }
    /**
     * Sets the number.
     * @param number number
     */
    public void setNumber(int number) { this.number = number; }

    /**
     * Returns the custom.
     * @return the requested result
     */
    public String getCustom() { return custom; }
    /**
     * Sets the custom.
     * @param custom custom
     */
    public void setCustom(String custom) { this.custom = custom; }

    /**
     * Creates a copy of this instance.
     * @return the computed result
     */
    public NumberFormatValue clone() {
        NumberFormatValue cloned = new NumberFormatValue();
        cloned.number = this.number;
        cloned.custom = this.custom;
        return cloned;
    }
}
