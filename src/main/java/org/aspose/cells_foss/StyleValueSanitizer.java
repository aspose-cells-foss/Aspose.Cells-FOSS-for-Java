package org.aspose.cells_foss;

/**
 * Sanitizes style values to ensure they fall within valid ranges.
 * This is a utility class with only static methods.
 */
public final class StyleValueSanitizer {

    /**
     * Initializes a new StyleValueSanitizer instance.
     */
    private StyleValueSanitizer() {}

    /**
     * Normalizes the indent level to the valid range [0, 250].
     * Returns 0 if the value is null or out of range.
     *
     * @param value the indent level value (nullable)
     * @return the normalized indent level
     */
    public static int normalizeIndentLevel(Integer value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value < 0 || value > 250) {
            return 0;
        }
        return value;
    }

    /**
     * Normalizes the text rotation angle to the valid range [0, 180] or 255.
     * Returns 0 if the value is null or out of range.
     *
     * @param value the text rotation value (nullable)
     * @return the normalized text rotation
     */
    public static int normalizeTextRotation(Integer value) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            return 0;
        }
        if (value == 255) {
            return 255;
        }
        if (value < 0 || value > 180) {
            return 0;
        }
        return value;
    }

    /**
     * Normalizes the reading order value to the valid range [0, 2].
     * Returns 0 if the value is null or out of range.
     *
     * @param value the reading order value (nullable)
     * @return the normalized reading order
     */
    public static int normalizeReadingOrder(Integer value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value < 0 || value > 2) {
            return 0;
        }
        return value;
    }
}
