package com.aspose.cells_foss;

/**
 * A support class for normalizing workbook property values.
 * This class provides helper methods to validate and normalize string values
 * used in various workbook properties.
 */
final class WorkbookPropertySupport {

    /**
     * Normalizes the "showObjects" property value.
     *
     * @param value the input value to normalize
     * @return the normalized value, or an empty string if input is null/whitespace
     */
    static String normalizeShowObjects(String value) {
        return normalizeChoice(value, "showObjects", "all", "placeholders", "none");
    }

    /**
     * Normalizes the "updateLinks" property value.
     *
     * @param value the input value to normalize
     * @return the normalized value, or an empty string if input is null/whitespace
     */
    static String normalizeUpdateLinks(String value) {
        return normalizeChoice(value, "updateLinks", "userSet", "never", "always");
    }

    /**
     * Normalizes the "visibility" property value.
     *
     * @param value the input value to normalize
     * @return the normalized value, or an empty string if input is null/whitespace
     */
    static String normalizeVisibility(String value) {
        return normalizeChoice(value, "visibility", "visible", "hidden", "veryHidden");
    }

    /**
     * Normalizes the "calculationMode" property value.
     *
     * @param value the input value to normalize
     * @return the normalized value, or an empty string if input is null/whitespace
     */
    static String normalizeCalculationMode(String value) {
        return normalizeChoice(value, "calcMode", "auto", "manual", "autoNoTable");
    }

    /**
     * Normalizes the "referenceMode" property value.
     *
     * @param value the input value to normalize
     * @return the normalized value, or an empty string if input is null/whitespace
     */
    static String normalizeReferenceMode(String value) {
        return normalizeChoice(value, "refMode", "A1", "R1C1");
    }

    /**
     * Normalizes a string value by comparing it (case-insensitive) against a list
     * of allowed values. If the value matches, the canonical form (from the allowed
     * list) is returned. If the input is null or whitespace, an empty string is returned.
     *
     * @param value     the value to normalize
     * @param propertyName the name of the property (used in error messages)
     * @param allowed   the list of allowed values
     * @return the normalized value, or an empty string if input is null/whitespace
     * @throws CellsException if the value does not match any allowed value
     */
    private static String normalizeChoice(String value, String propertyName, String... allowed) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isBlank()) {
            return "";
        }

        String trimmed = value.trim();
        for (String allowedValue : allowed) {
            if (allowedValue.equalsIgnoreCase(trimmed)) {
                return allowedValue;
            }
        }

        throw new CellsException("Unsupported " + propertyName + " value '" + value + "'.");
    }
}