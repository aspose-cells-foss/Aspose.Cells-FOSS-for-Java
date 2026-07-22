package org.aspose.cells_foss;

/**
 * Utility class for normalizing and validating defined names.
 */
final class DefinedNameUtility {

    /**
     * Initializes a new DefinedNameUtility instance.
     */
    private DefinedNameUtility() {}

    /** The built-in print area defined name. */
    static final String PRINT_AREA_DEFINED_NAME = "_xlnm.Print_Area";

    /** The built-in print titles defined name. */
    static final String PRINT_TITLES_DEFINED_NAME = "_xlnm.Print_Titles";

    /** The built-in filter database defined name. */
    static final String FILTER_DATABASE_DEFINED_NAME = "_xlnm._FilterDatabase";

    /**
     * Checks if the given name is a reserved built-in name.
     *
     * @param name the name to check
     * @return true if the name is reserved; false otherwise
     */
    static boolean isReservedName(String name) {
        // Handle the relevant branch before the state changes.
        if (name == null) return false;
        return name.equalsIgnoreCase(PRINT_AREA_DEFINED_NAME) ||
               name.equalsIgnoreCase(PRINT_TITLES_DEFINED_NAME) ||
               name.equalsIgnoreCase(FILTER_DATABASE_DEFINED_NAME);
    }

    /**
     * Normalizes a defined name.
     *
     * @param name the name to normalize
     * @return the normalized name
     * @throws CellsException if the name is null, empty, or reserved
     */
    static String normalizeName(String name) {
        String normalized = (name != null ? name.trim() : "");
        // Handle the relevant branch before the state changes.
        if (normalized.isEmpty()) {
            throw new CellsException("Defined name must be non-empty.");
        }

        if (isReservedName(normalized)) {
            throw new CellsException("Built-in print defined names must be managed through PageSetup.");
        }

        return normalized;
    }

    /**
     * Normalizes a formula string.
     *
     * @param formula the formula to normalize
     * @return the normalized formula without leading equals sign
     * @throws CellsException if the formula is null, empty, or becomes empty after trimming
     */
    static String normalizeFormula(String formula) {
        String normalized = (formula != null ? formula.trim() : "");
        // Handle the relevant branch before the state changes.
        if (normalized.startsWith("=")) {
            normalized = normalized.substring(1).trim();
        }

        if (normalized.isEmpty()) {
            throw new CellsException("Defined name formula must be non-empty.");
        }

        return normalized;
    }

    /**
     * Normalizes a comment string.
     *
     * @param comment the comment to normalize
     * @return the trimmed comment, or empty string if null
     */
    static String normalizeComment(String comment) {
        return (comment != null ? comment.trim() : "");
    }

    /**
     * Checks if two scope values are the same.
     *
     * @param left the first scope
     * @param right the second scope
     * @return true if both are null or both have the same value; false otherwise
     */
    static boolean sameScope(Integer left, Integer right) {
        // Handle the relevant branch before the state changes.
        if (left == null && right == null) {
            return true;
        }
        return left != null && right != null && left.equals(right);
    }
}
