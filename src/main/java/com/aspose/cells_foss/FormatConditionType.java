package com.aspose.cells_foss;

/**
 * Represents the type of a format condition.
 */
public enum FormatConditionType {
    /** A condition based on cell value. */
    CELL_VALUE,
    /** A condition based on an expression. */
    EXPRESSION,
    /** A condition based on containing text. */
    CONTAINS_TEXT,
    /** A condition based on not containing text. */
    NOT_CONTAINS_TEXT,
    /** A condition based on beginning with text. */
   begins_with,
    /** A condition based on ending with text. */
    ENDS_WITH,
    /** A condition based on a time period. */
    TIME_PERIOD,
    /** A condition for duplicate values. */
    DUPLICATE_VALUES,
    /** A condition for unique values. */
    UNIQUE_VALUES,
    /** A condition for top 10 values. */
    TOP_10,
    /** A condition for bottom 10 values. */
    BOTTOM_10,
    /** A condition for above average values. */
    ABOVE_AVERAGE,
    /** A condition for below average values. */
    BELOW_AVERAGE,
    /** A color scale condition. */
    COLOR_SCALE,
    /** A data bar condition. */
    DATA_BAR,
    /** An icon set condition. */
    ICON_SET
}