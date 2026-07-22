package org.aspose.cells_foss;

import java.util.HashMap;
import java.util.Map;

/**
 * Provides built-in number format functionality.
 */
public final class NumberFormat {

    /**
     * Prevents instantiation of this utility class.
     */
    private NumberFormat() {}

    private static final Map<Integer, String> BUILT_IN_FORMATS = createBuiltInFormats();

    /**
     * Gets the built-in format string for the specified format ID.
     *
     * @param formatId the format ID
     * @return the format string, or "General" if not found
     */
    public static String getBuiltInFormat(int formatId) {
        String formatCode = BUILT_IN_FORMATS.get(formatId);
        // Handle the relevant branch before the state changes.
        if (formatCode != null) {
            return formatCode;
        }
        return "General";
    }

    /**
     * Checks if the format code is a built-in format.
     *
     * @param formatCode the format code to check
     * @return true if it's a built-in format, false otherwise
     */
    public static boolean isBuiltInFormat(String formatCode) {
        return getBuiltInFormatId(formatCode) != null;
    }

    /**
     * Gets the built-in format ID for the specified format code.
     *
     * @param formatCode the format code
     * @return the format ID, or null if not found
     */
    public static Integer getBuiltInFormatId(String formatCode) {
        // Handle the relevant branch before the state changes.
        if (formatCode == null || formatCode.isBlank()) {
            return 0;
        }

        String trimmed = formatCode.trim();
        for (Map.Entry<Integer, String> entry : BUILT_IN_FORMATS.entrySet()) {
            if (trimmed.equals(entry.getValue())) {
                return entry.getKey();
            }
        }

        return null;
    }

    /**
     * Resolves the format code for the given number and custom format.
     *
     * @param number the number
     * @param custom the custom format, or null
     * @return the format code
     */
    public static String resolveFormatCode(int number, String custom) {
        // Handle the relevant branch before the state changes.
        if (custom != null && !custom.isEmpty()) {
            return custom;
        }
        return getBuiltInFormat(number);
    }

    /**
     * Creates the built in formats.
     * @return the requested result
     */
    private static Map<Integer, String> createBuiltInFormats() {
        Map<Integer, String> formats = new HashMap<>();
        // Walk the current collection so every entry is processed consistently.
        formats.put(0, "General");
        formats.put(1, "0");
        formats.put(2, "0.00");
        formats.put(3, "#,##0");
        formats.put(4, "#,##0.00");
        formats.put(5, "$#,##0_);($#,##0)");
        formats.put(6, "$#,##0_);[Red]($#,##0)");
        formats.put(7, "$#,##0.00_);($#,##0.00)");
        formats.put(8, "$#,##0.00_);[Red]($#,##0.00)");
        formats.put(9, "0%");
        formats.put(10, "0.00%");
        formats.put(11, "0.00E+00");
        formats.put(12, "# ?/?");
        formats.put(13, "# ??/??");
        formats.put(14, "m/d/yy");
        formats.put(15, "d-mmm-yy");
        formats.put(16, "d-mmm");
        formats.put(17, "mmm-yy");
        formats.put(18, "h:mm AM/PM");
        formats.put(19, "h:mm:ss AM/PM");
        formats.put(20, "h:mm");
        formats.put(21, "h:mm:ss");
        formats.put(22, "m/d/yy h:mm");
        formats.put(37, "#,##0_);(#,##0)");
        formats.put(38, "#,##0_);[Red](#,##0)");
        formats.put(39, "#,##0.00_);(#,##0.00)");
        formats.put(40, "#,##0.00_);[Red](#,##0.00)");
        formats.put(41, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)");
        formats.put(42, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)");
        formats.put(43, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)");
        formats.put(44, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)");
        formats.put(45, "mm:ss");
        formats.put(46, "[h]:mm:ss");
        formats.put(47, "mm:ss.0");
        formats.put(48, "##0.0E+0");
        formats.put(49, "@");
        formats.put(50, "General");
        formats.put(51, "0_);(0)");
        formats.put(52, "0_);[Red](0)");
        formats.put(53, "0_);(0)");
        formats.put(54, "0_);[Red](0)");
        formats.put(55, "0_);(0)");
        formats.put(56, "0_);[Red](0)");
        formats.put(57, "0_);(0)");
        formats.put(58, "0_);[Red](0)");
        formats.put(59, "0_);(0)");
        formats.put(60, "0_);[Red](0)");
        formats.put(61, "0_);(0)");
        formats.put(62, "0_);[Red](0)");
        formats.put(63, "0_);(0)");
        formats.put(64, "0_);[Red](0)");
        formats.put(65, "0_);(0)");
        formats.put(66, "0_);[Red](0)");
        formats.put(67, "0_);(0)");
        formats.put(68, "0_);[Red](0)");
        formats.put(69, "0_);(0)");
        formats.put(70, "0_);[Red](0)");
        formats.put(71, "0_);(0)");
        formats.put(72, "0_);[Red](0)");
        formats.put(73, "0_);(0)");
        formats.put(74, "0_);[Red](0)");
        formats.put(75, "0_);(0)");
        formats.put(76, "0_);[Red](0)");
        formats.put(77, "0_);(0)");
        formats.put(78, "0_);[Red](0)");
        formats.put(79, "0_);(0)");
        formats.put(80, "0_);[Red](0)");
        formats.put(81, "0_);(0)");
        formats.put(82, "0_);[Red](0)");
        return formats;
    }
}

