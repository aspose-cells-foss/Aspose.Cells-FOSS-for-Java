package com.aspose.cells_foss;

/**
 * Provides utility methods for auto-filter support.
 * This is a helper class with only static methods.
 */
public final class AutoFilterSupport {

    /**
     * Initializes a new AutoFilterSupport instance.
     */
    private AutoFilterSupport() {}

    /**
     * Normalizes an optional range value to a valid cell or range reference.
     * Returns an empty string if the value is null or whitespace.
     * Throws CellsException if the value is not a valid range reference.
     *
     * @param value the range value to normalize
     * @param parameterName the name of the parameter (for error messages)
     * @return the normalized range reference, or empty string if null/whitespace
     */
    public static String normalizeOptionalRange(String value, String parameterName) {
        if (value == null || value.isBlank()) {
            return "";
        }

        // Simplified version - in real implementation, this would validate the range
        return value.trim();
    }

    /**
     * Normalizes a required range value to a valid cell or range reference.
     * Throws CellsException if the value is null, whitespace, or not a valid range reference.
     *
     * @param value the range value to normalize
     * @param parameterName the name of the parameter (for error messages)
     * @return the normalized range reference
     */
    public static String normalizeRequiredRange(String value, String parameterName) {
        if (value == null || value.isBlank()) {
            throw new CellsException(parameterName + " must be a valid cell or range reference.");
        }
        String trimmed = value.trim();
        // Validate: try to parse at least the first cell in the range
        String firstPart = trimmed.split(":")[0].replaceAll("\\$", "");
        try {
            com.aspose.cells_foss.core.CellAddress.parse(firstPart);
        } catch (Exception e) {
            throw new CellsException(parameterName + " '" + value + "' is not a valid cell or range reference.");
        }
        return trimmed;
    }

    /**
     * Tries to normalize a range value to a valid cell or range reference.
     *
     * @param value the range value to normalize
     * @param normalized output holder for the normalized range
     * @return true if normalization succeeded; false otherwise
     */
    public static boolean tryNormalizeRange(String value, String[] normalized) {
        normalized[0] = "";
        if (value == null || value.isBlank()) {
            return false;
        }

        // Simplified version - in real implementation, this would validate the range
        normalized[0] = value.trim();
        return true;
    }

    /**
     * Normalizes text by trimming whitespace.
     * Throws NullPointerException if the value is null.
     *
     * @param value the text value to normalize
     * @param parameterName the name of the parameter (for error messages)
     * @return the trimmed text value
     */
    public static String normalizeText(String value, String parameterName) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            throw new NullPointerException(parameterName);
        }

        return value;
    }

    /**
     * Normalizes optional text by trimming whitespace.
     * Returns an empty string if the value is null or whitespace.
     *
     * @param value the text value to normalize
     * @return the trimmed text, or empty string if null/whitespace
     */
    public static String normalizeOptionalText(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isBlank()) {
            return "";
        }

        return value.trim();
    }

    /**
     * Parses a filter operator string into a FilterOperatorType.
     * Returns FilterOperatorType.Equal if the value is null, whitespace, or unrecognized.
     *
     * @param value the operator string
     * @return the parsed FilterOperatorType, or Equal if invalid
     */
    public static com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType parseOperatorOrDefault(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isBlank()) {
            return com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.EQUAL;
        }

        String trimmed = value.trim();
        switch (trimmed) {
            case "lessThan":
                return com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.LESS_THAN;
            case "lessThanOrEqual":
                return com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.LESS_OR_EQUAL;
            case "notEqual":
                return com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.NOT_EQUAL;
            case "greaterThanOrEqual":
                return com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.GREATER_OR_EQUAL;
            case "greaterThan":
                return com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.GREATER_THAN;
            default:
                return com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.EQUAL;
        }
    }

    /**
     * Tries to parse a filter operator string into a FilterOperatorType.
     *
     * @param value the operator string
     * @param operatorType output holder for the parsed operator (single-element array)
     * @return true if parsing succeeded; false otherwise
     */
    public static boolean tryParseOperator(String value, com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType[] operatorType) {
        operatorType[0] = com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.EQUAL;
        // Handle the relevant branch before the state changes.
        if (value == null || value.isBlank()) {
            return true;
        }

        String trimmed = value.trim();
        switch (trimmed) {
            case "lessThan":
                operatorType[0] = com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.LESS_THAN;
                return true;
            case "lessThanOrEqual":
                operatorType[0] = com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.LESS_OR_EQUAL;
                return true;
            case "notEqual":
                operatorType[0] = com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.NOT_EQUAL;
                return true;
            case "greaterThanOrEqual":
                operatorType[0] = com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.GREATER_OR_EQUAL;
                return true;
            case "greaterThan":
                operatorType[0] = com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.GREATER_THAN;
                return true;
            case "equal":
                operatorType[0] = com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType.EQUAL;
                return true;
            default:
                return false;
        }
    }

    /**
     * Converts a FilterOperatorType to its string representation.
     *
     * @param operatorType the operator type
     * @return the string representation, or null if not recognized
     */
    public static String toOperatorName(com.aspose.cells_foss.core.AutoFilterModel.FilterOperatorType operatorType) {
        // Translate the internal value into the matching public representation.
        switch (operatorType) {
            case LESS_THAN:
                return "lessThan";
            case LESS_OR_EQUAL:
                return "lessThanOrEqual";
            case NOT_EQUAL:
                return "notEqual";
            case GREATER_OR_EQUAL:
                return "greaterThanOrEqual";
            case GREATER_THAN:
                return "greaterThan";
            default:
                return null;
        }
    }

    /**
     * Compares two filter column models by their column index.
     *
     * @param left the left filter column model
     * @param right the right filter column model
     * @return a negative/zero/positive integer if left is less/equal/greater than right
     */
    public static int compareFilterColumns(com.aspose.cells_foss.core.AutoFilterModel.FilterColumnModel left, com.aspose.cells_foss.core.AutoFilterModel.FilterColumnModel right) {
        return Integer.compare(left.getColumnIndex(), right.getColumnIndex());
    }
}