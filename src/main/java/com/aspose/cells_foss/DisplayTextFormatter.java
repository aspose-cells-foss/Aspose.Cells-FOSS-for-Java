package com.aspose.cells_foss;

import com.aspose.cells_foss.core.DateSerialConverter;
import com.aspose.cells_foss.core.DateSystem;
import com.aspose.cells_foss.core.NumberFormatValue;
import com.aspose.cells_foss.core.StyleValue;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.List;
import java.util.Locale;

/**
 * Internal utility class for formatting display text values.
 */
final class DisplayTextFormatter {
    private static final long NANOS_PER_DAY = 86_400_000_000_000L;

    /**
     * Initializes a new DisplayTextFormatter instance.
     */
    private DisplayTextFormatter() {}

    /**
     * Formats a value to its string representation without any formatting applied.
     */
    static String formatStringValue(Object value) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            return "";
        }

        if (value instanceof String text) {
            return text;
        }

        if (value instanceof Boolean booleanValue) {
            return booleanValue ? "TRUE" : "FALSE";
        }

        if (value instanceof LocalDateTime dateTime) {
            return formatRawDateTimeValue(dateTime);
        }

        if (value instanceof Number) {
            return formatNumberValue((Number) value);
        }

        return value.toString();
    }

    /**
     * Formats a value for display with the given style and culture.
     */
    static String formatDisplayValue(Object value, StyleValue style, Locale workbookCulture, DateSystem dateSystem) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            return "";
        }

        StyleValue effectiveStyle = style != null ? style : StyleValue.getDefault();
        Locale effectiveCulture = workbookCulture != null ? workbookCulture : Locale.US;
        DateSystem effectiveDateSystem = dateSystem != null ? dateSystem : DateSystem.WINDOWS_1900;

        if (value instanceof String text) {
            return formatTextValue(text, effectiveStyle);
        }

        if (value instanceof Boolean booleanValue) {
            return booleanValue ? "TRUE" : "FALSE";
        }

        if (value instanceof LocalDateTime dateTime) {
            return formatDateTimeValue(dateTime, effectiveStyle, effectiveCulture, effectiveDateSystem);
        }

        if (isNumericValue(value)) {
            return formatNumericValue(value, effectiveStyle, effectiveCulture);
        }

        return value.toString();
    }

    /**
     * Formats the text value.
     * @param value value to apply
     * @param style style to apply
     * @return the computed result
     */
    private static String formatTextValue(String value, StyleValue style) {
        String formatCode = resolveFormatCode(style);
        if (formatCode == null || formatCode.isBlank() || formatCode.equalsIgnoreCase("General")) {
            return value;
        }

        List<DisplayFormatSectionInfo> sections = DisplayTextFormatterSupport.parseSections(formatCode);
        DisplayFormatSectionInfo textSection = DisplayTextFormatterSupport.selectTextSection(sections);
        if (textSection == null || textSection.getRaw() == null) {
            return value;
        }

        String pattern = DisplayTextFormatterSupport.stripDirectiveBrackets(textSection.getRaw(), false);
        String formatted = DisplayTextFormatterSupport.expandSectionPattern(pattern, value, true);
        if (formatted.isEmpty()) {
            return value;
        }

        return formatted;
    }

    /**
     * Formats the numeric value.
     * @param value value to apply
     * @param style style to apply
     * @param workbookCulture workbook culture
     * @return the computed result
     */
    private static String formatNumericValue(Object value, StyleValue style, Locale workbookCulture) {
        String formatCode = resolveFormatCode(style);
        if (formatCode == null || formatCode.isBlank() || formatCode.equalsIgnoreCase("General")) {
            return formatStringValue(value);
        }

        List<DisplayFormatSectionInfo> sections = DisplayTextFormatterSupport.parseSections(formatCode);
        if (sections.isEmpty()) {
            return formatStringValue(value);
        }

        double numericValue = DisplayTextFormatterSupport.convertToDouble(value);
        boolean[] useAbsoluteValue = {false};
        DisplayFormatSectionInfo selectedSection =
                DisplayTextFormatterSupport.selectNumericSection(sections, numericValue, useAbsoluteValue);
        if (selectedSection == null || selectedSection.getRaw() == null || selectedSection.getRaw().isBlank()) {
            return formatStringValue(value);
        }

        String[] fractionResult = new String[1];
        if (tryFormatFraction(numericValue, selectedSection.getRaw(), useAbsoluteValue[0], fractionResult)) {
            return fractionResult[0];
        }

        Locale[] sectionCulture = {workbookCulture != null ? workbookCulture : Locale.US};
        String localizedSection =
                DisplayTextLocaleSupport.applyLocaleDirectives(selectedSection.getRaw(), sectionCulture[0], sectionCulture);
        String sanitizedSection = DisplayTextFormatterSupport.sanitizeNumericSection(localizedSection);
        if (sanitizedSection == null || sanitizedSection.isBlank()) {
            return formatStringValue(value);
        }

        if (!DisplayTextFormatterSupport.containsNumericPlaceholder(sanitizedSection)) {
            String literal = DisplayTextFormatterSupport.expandSectionPattern(sanitizedSection, "", false);
            if (!literal.isEmpty()) {
                return literal;
            }
            return formatStringValue(value);
        }

        try {
            Number formattedValue = useAbsoluteValue[0] ? getAbsoluteNumber(value) : (Number) value;
            return formatNumericSection(formattedValue, sanitizedSection, sectionCulture[0]);
        } catch (Exception e) {
            return formatStringValue(value);
        }
    }

    /**
     * Formats the date time value.
     * @param value value to apply
     * @param style style to apply
     * @param workbookCulture workbook culture
     * @param dateSystem workbook date system
     * @return the computed result
     */
    private static String formatDateTimeValue(LocalDateTime value,
                                              StyleValue style,
                                              Locale workbookCulture,
                                              DateSystem dateSystem) {
        String formatCode = resolveFormatCode(style);
        if (formatCode == null || formatCode.isBlank() || formatCode.equalsIgnoreCase("General")) {
            return formatRawDateTimeValue(value);
        }

        List<DisplayFormatSectionInfo> sections = DisplayTextFormatterSupport.parseSections(formatCode);
        DisplayFormatSectionInfo section = DisplayTextFormatterSupport.selectDateTimeSection(sections);
        if (section == null || section.getRaw() == null || section.getRaw().isBlank()) {
            return formatRawDateTimeValue(value);
        }

        Locale[] sectionCulture = {workbookCulture != null ? workbookCulture : Locale.US};
        String localizedSection =
                DisplayTextLocaleSupport.applyLocaleDirectives(section.getRaw(), sectionCulture[0], sectionCulture);
        String sectionFormat = DisplayTextFormatterSupport.stripDirectiveBrackets(localizedSection, true);
        if (sectionFormat == null || sectionFormat.isBlank()) {
            return formatRawDateTimeValue(value);
        }

        try {
            if (DisplayTextDateFormatSupport.containsElapsedTimeToken(sectionFormat)) {
                long totalNanos = Math.round(DateSerialConverter.toSerial(value, dateSystem) * NANOS_PER_DAY);
                Duration duration = Duration.ofNanos(totalNanos);
                return DisplayTextDateFormatSupport.formatElapsedTimeValue(duration, sectionFormat, sectionCulture[0]);
            }
            return DisplayTextDateFormatSupport.formatDateTimeValue(value, sectionFormat, sectionCulture[0]);
        } catch (Exception e) {
            return formatRawDateTimeValue(value);
        }
    }

    /**
     * Formats the raw date time value.
     * @param value value to apply
     * @return the computed result
     */
    private static String formatRawDateTimeValue(LocalDateTime value) {
        // Handle the relevant branch before the state changes.
        if (value.toLocalTime().equals(LocalTime.MIDNIGHT)) {
            return value.format(java.time.format.DateTimeFormatter.ofPattern("M/d/yyyy", Locale.ENGLISH));
        }

        return value.format(java.time.format.DateTimeFormatter.ofPattern("M/d/yyyy H:mm", Locale.ENGLISH));
    }

    /**
     * Resolves the effective format code for the style.
     * @param style style
     * @return format code
     */
    private static String resolveFormatCode(StyleValue style) {
        StyleValue effectiveStyle = style != null ? style : StyleValue.getDefault();
        NumberFormatValue numberFormat = effectiveStyle.getNumberFormat();
        if (numberFormat == null) {
            return "General";
        }

        return NumberFormat.resolveFormatCode(numberFormat.getNumber(), numberFormat.getCustom());
    }

    /**
     * Formats a numeric section with locale-aware decimal formatting.
     * @param value value
     * @param section section
     * @param locale locale
     * @return formatted result
     */
    private static String formatNumericSection(Number value, String section, Locale locale) {
        DecimalFormatSymbols symbols = DecimalFormatSymbols.getInstance(locale != null ? locale : Locale.US);
        DecimalFormat format = new DecimalFormat(toJavaDecimalPattern(section), symbols);
        format.setRoundingMode(RoundingMode.HALF_UP);
        String formatted = format.format(value);
        if (containsScientificPlaceholder(section)) {
            return ensurePositiveExponentSign(formatted);
        }
        return formatted;
    }

    /**
     * Converts an Excel-style numeric section into a Java DecimalFormat pattern.
     * @param section section
     * @return Java pattern
     */
    private static String toJavaDecimalPattern(String section) {
        StringBuilder builder = new StringBuilder(section.length());
        boolean inQuote = false;

        for (int index = 0; index < section.length(); index++) {
            char character = section.charAt(index);
            if (character == '"') {
                inQuote = !inQuote;
                builder.append(character);
                continue;
            }

            if (!inQuote && (character == 'E' || character == 'e') && index + 1 < section.length()
                    && section.charAt(index + 1) == '+') {
                builder.append(character);
                continue;
            }

            if (!inQuote && character == '+'
                    && index > 0
                    && (section.charAt(index - 1) == 'E' || section.charAt(index - 1) == 'e')) {
                continue;
            }

            builder.append(character);
        }

        return builder.toString();
    }

    /**
     * Indicates whether the section contains scientific notation placeholders.
     * @param section section
     * @return true when scientific notation is present
     */
    private static boolean containsScientificPlaceholder(String section) {
        boolean inQuote = false;
        for (int index = 0; index < section.length() - 1; index++) {
            char character = section.charAt(index);
            if (character == '"') {
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote && (character == 'E' || character == 'e')) {
                char next = section.charAt(index + 1);
                if (next == '+' || next == '-' || next == '0' || next == '#') {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Ensures positive exponents keep the explicit plus sign used by Excel.
     * @param formatted formatted text
     * @return adjusted text
     */
    private static String ensurePositiveExponentSign(String formatted) {
        int exponentIndex = Math.max(formatted.indexOf('E'), formatted.indexOf('e'));
        if (exponentIndex < 0 || exponentIndex + 1 >= formatted.length()) {
            return formatted;
        }

        char sign = formatted.charAt(exponentIndex + 1);
        if (sign == '+' || sign == '-') {
            return formatted;
        }

        return formatted.substring(0, exponentIndex + 1) + "+" + formatted.substring(exponentIndex + 1);
    }

    /**
     * Converts a numeric value to its absolute value while preserving the concrete wrapper type.
     * @param value value
     * @return absolute value
     */
    private static Number getAbsoluteNumber(Object value) {
        Object absolute = DisplayTextFormatterSupport.getAbsoluteNumericValue(value);
        if (absolute instanceof Number number) {
            return number;
        }
        throw new IllegalStateException("Value is not numeric.");
    }

    /**
     * Attempts to process format fraction.
     * @param numericValue numeric value
     * @param section section
     * @param useAbsoluteValue use absolute value
     * @param result result
     * @return true when the condition is satisfied
     */
    private static boolean tryFormatFraction(double numericValue, String section, boolean useAbsoluteValue, String[] result) {
        result[0] = "";
        String sanitizedSection = DisplayTextFormatterSupport.sanitizeNumericSection(section);
        
        // Handle the relevant branch before the state changes.
        if (sanitizedSection == null || sanitizedSection.indexOf('/') < 0) {
            return false;
        }

        int slashIndex = sanitizedSection.indexOf('/');
        if (slashIndex <= 0) {
            return false;
        }

        int denominatorDigits = 0;
        for (int index = slashIndex + 1; index < sanitizedSection.length(); index++) {
            char character = sanitizedSection.charAt(index);
            if (character == '#' || character == '0') {
                denominatorDigits++;
                continue;
            }

            if (Character.isWhitespace(character)) {
                continue;
            }

            break;
        }

        if (denominatorDigits <= 0) {
            return false;
        }

        double absoluteValue = Math.abs(numericValue);
        long wholePart = (long) Math.floor(absoluteValue);
        double fractionalPart = absoluteValue - wholePart;
        
        if (fractionalPart < 1E-12) {
            result[0] = formatWholeFractionResult(wholePart, useAbsoluteValue, numericValue);
            return true;
        }

        int maxDenominator = 1;
        for (int index = 0; index < denominatorDigits; index++) {
            maxDenominator *= 10;
        }

        maxDenominator -= 1;
        int bestNumerator = 0;
        int bestDenominator = 1;
        double bestError = Double.MAX_VALUE;

        for (int denominator = 1; denominator <= maxDenominator; denominator++) {
            int numerator = (int) Math.round(fractionalPart * denominator);
            if (numerator == 0) {
                continue;
            }

            if (numerator > denominator) {
                numerator = denominator;
            }

            double candidate = (double) numerator / denominator;
            double error = Math.abs(fractionalPart - candidate);
            if (error < bestError) {
                bestError = error;
                bestNumerator = numerator;
                bestDenominator = denominator;
            }
        }

        if (bestNumerator == 0) {
            result[0] = formatWholeFractionResult(wholePart, useAbsoluteValue, numericValue);
            return true;
        }

        int greatestCommonDivisor = greatestCommonDivisor(bestNumerator, bestDenominator);
        bestNumerator /= greatestCommonDivisor;
        bestDenominator /= greatestCommonDivisor;

        if (bestNumerator == bestDenominator) {
            wholePart++;
            bestNumerator = 0;
        }

        String prefix = "";
        if (!useAbsoluteValue && numericValue < 0) {
            prefix = "-";
        }

        if (bestNumerator == 0) {
            result[0] = prefix + Long.toString(wholePart);
        } else if (wholePart == 0) {
            result[0] = prefix + Integer.toString(bestNumerator) + "/" + Integer.toString(bestDenominator);
        } else {
            result[0] = prefix + Long.toString(wholePart) + " " + Integer.toString(bestNumerator) + "/" + Integer.toString(bestDenominator);
        }

        return true;
    }

    /**
     * Formats the whole fraction result.
     * @param wholePart whole part
     * @param useAbsoluteValue use absolute value
     * @param numericValue numeric value
     * @return the computed result
     */
    private static String formatWholeFractionResult(long wholePart, boolean useAbsoluteValue, double numericValue) {
        // Handle the relevant branch before the state changes.
        if (!useAbsoluteValue && numericValue < 0) {
            return "-" + Long.toString(wholePart);
        }

        return Long.toString(wholePart);
    }

    /**
     * Processes greatest common divisor.
     * @param left left
     * @param right right
     * @return the computed result
     */
    private static int greatestCommonDivisor(int left, int right) {
        int first = Math.abs(left);
        int second = Math.abs(right);
        // Walk the current collection so every entry is processed consistently.
        while (second != 0) {
            int remainder = first % second;
            first = second;
            second = remainder;
        }

        if (first == 0) {
            return 1;
        }

        return first;
    }

    /**
     * Formats the number value.
     * @param number number
     * @return the computed result
     */
    private static String formatNumberValue(Number number) {
        return number.toString();
    }

    /**
     * Indicates whether numeric value.
     * @param value value to apply
     * @return true when the condition is satisfied
     */
    private static boolean isNumericValue(Object value) {
        return value instanceof Number;
    }
}
