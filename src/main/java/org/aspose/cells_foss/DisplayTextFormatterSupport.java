package org.aspose.cells_foss;

import java.util.ArrayList;
import java.util.List;

/**
 * Internal utility class for selecting and formatting display text sections.
 */
public final class DisplayTextFormatterSupport {

    /**
     * Prevents instantiation of this utility class.
     */
    private DisplayTextFormatterSupport() {}

    /**
     * Parses format code into a list of section info objects.
     */
    public static List<DisplayFormatSectionInfo> parseSections(String formatCode) {
        List<String> rawSections = splitSections(formatCode);
        List<DisplayFormatSectionInfo> sections = new ArrayList<>(rawSections.size());
        // Walk the current collection so every entry is processed consistently.
        for (String rawSection : rawSections) {
            DisplayFormatSectionInfo section = new DisplayFormatSectionInfo();
            section.setRaw(rawSection);

            String[] conditionOperatorHolder = {""};
            double[] conditionValueHolder = {0.0};
            if (tryParseSectionCondition(rawSection, conditionOperatorHolder, conditionValueHolder)) {
                section.setHasCondition(true);
                section.setConditionOperator(conditionOperatorHolder[0]);
                section.setConditionValue(conditionValueHolder[0]);
            }

            sections.add(section);
        }
        return sections;
    }

    /**
     * Splits sections.
     * @param formatCode format code
     * @return the computed result
     */
    public static List<String> splitSections(String formatCode) {
        List<String> sections = new ArrayList<>();
        StringBuilder builder = new StringBuilder();
        boolean inQuote = false;

        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < formatCode.length(); index++) {
            char character = formatCode.charAt(index);
            if (character == '"') {
                builder.append(character);
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote) {
                if (character == '\\' && index + 1 < formatCode.length()) {
                    builder.append(character);
                    index++;
                    builder.append(formatCode.charAt(index));
                    continue;
                }

                if (character == ';') {
                    sections.add(builder.toString());
                    builder.setLength(0);
                    continue;
                }
            }

            builder.append(character);
        }

        sections.add(builder.toString());
        return sections;
    }

    /**
     * Attempts to process parse section condition.
     * @param section section
     * @param conditionOperator condition operator
     * @param conditionValue condition value
     * @return true when the condition is satisfied
     */
    public static boolean tryParseSectionCondition(String section, String[] conditionOperator, double[] conditionValue) {
        conditionOperator[0] = "";
        conditionValue[0] = 0.0;
        boolean inQuote = false;

        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < section.length(); index++) {
            char character = section.charAt(index);
            if (character == '"') {
                inQuote = !inQuote;
                continue;
            }

            if (inQuote || character != '[') {
                continue;
            }

            int endIndex = section.indexOf(']', index + 1);
            if (endIndex < 0) {
                break;
            }

            String token = section.substring(index + 1, endIndex).trim();
            if (tryParseConditionToken(token, conditionOperator, conditionValue)) {
                return true;
            }

            index = endIndex;
        }

        return false;
    }

    /**
     * Attempts to process parse condition token.
     * @param token token
     * @param conditionOperator condition operator
     * @param conditionValue condition value
     * @return true when the condition is satisfied
     */
    public static boolean tryParseConditionToken(String token, String[] conditionOperator, double[] conditionValue) {
        conditionOperator[0] = "";
        conditionValue[0] = 0.0;
        if (token == null || token.isBlank()) {
            return false;
        }

        String[] operators = {">=", "<=", "<>", ">", "<", "="};
        for (String candidate : operators) {
            if (token.startsWith(candidate)) {
                String numberPart = token.substring(candidate.length()).trim();
                try {
                    double parsedValue = Double.parseDouble(numberPart);
                    conditionOperator[0] = candidate;
                    conditionValue[0] = parsedValue;
                    return true;
                } catch (NumberFormatException e) {
                    // continue
                }
            }
        }

        return false;
    }

    /**
     * Evaluates condition.
     * @param conditionOperator condition operator
     * @param numericValue numeric value
     * @param conditionValue condition value
     * @return true when the condition is satisfied
     */
    public static boolean evaluateCondition(String conditionOperator, double numericValue, double conditionValue) {
        // Translate the internal value into the matching public representation.
        switch (conditionOperator) {
            case ">":
                return numericValue > conditionValue;
            case ">=":
                return numericValue >= conditionValue;
            case "<":
                return numericValue < conditionValue;
            case "<=":
                return numericValue <= conditionValue;
            case "=":
                return Math.abs(numericValue - conditionValue) < 1E-12;
            case "<>":
                return Math.abs(numericValue - conditionValue) >= 1E-12;
            default:
                return false;
        }
    }

    /**
     * Processes should use absolute value.
     * @param section section
     * @param numericValue numeric value
     * @return true when the condition is satisfied
     */
    public static boolean shouldUseAbsoluteValue(String section, double numericValue) {
        // Handle the relevant branch before the state changes.
        if (numericValue >= 0) {
            return false;
        }

        String sanitizedSection = sanitizeNumericSection(section);
        return sanitizedSection.indexOf('-') < 0;
    }

    /**
     * Sanitizes numeric section.
     * @param section section
     * @return the computed result
     */
    public static String sanitizeNumericSection(String section) {
        String withoutDirectives = stripDirectiveBrackets(section, false);
        StringBuilder builder = new StringBuilder(withoutDirectives.length());

        boolean inQuote = false;
        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < withoutDirectives.length(); index++) {
            char character = withoutDirectives.charAt(index);
            if (character == '"') {
                builder.append(character);
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote) {
                if (character == '_') {
                    builder.append(' ');
                    index++;
                    continue;
                }

                if (character == '*') {
                    if (index + 1 < withoutDirectives.length()) {
                        index++;
                        for (int repeat = 0; repeat < 3; repeat++) {
                            builder.append(withoutDirectives.charAt(index));
                        }
                    }
                    continue;
                }

                if (character == '\\') {
                    if (index + 1 < withoutDirectives.length()) {
                        index++;
                        builder.append(withoutDirectives.charAt(index));
                    }
                    continue;
                }

                if (character == '?') {
                    builder.append('#');
                    continue;
                }

                if (character == '[' || character == ']') {
                    continue;
                }
            }

            builder.append(character);
        }

        return builder.toString().trim();
    }

    /**
     * Strips directive brackets.
     * @param section section
     * @param preserveElapsedTokens preserve elapsed tokens
     * @return the computed result
     */
    public static String stripDirectiveBrackets(String section, boolean preserveElapsedTokens) {
        StringBuilder builder = new StringBuilder(section.length());
        boolean inQuote = false;

        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < section.length(); index++) {
            char character = section.charAt(index);
            if (character == '"') {
                builder.append(character);
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote && character == '[') {
                int endIndex = section.indexOf(']', index + 1);
                if (endIndex < 0) {
                    continue;
                }

                String token = section.substring(index + 1, endIndex);
                if (preserveElapsedTokens && DisplayTextDateFormatSupport.isElapsedToken(token)) {
                    builder.append('[');
                    builder.append(token);
                    builder.append(']');
                }

                index = endIndex;
                continue;
            }

            builder.append(character);
        }

        return builder.toString();
    }

    /**
     * Checks whether the content contains numeric placeholder.
     * @param pattern pattern
     * @return true when the condition is satisfied
     */
    public static boolean containsNumericPlaceholder(String pattern) {
        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < pattern.length(); index++) {
            char character = pattern.charAt(index);
            if (character == '0' || character == '#' || character == '?' || character == '.' ||
                character == '%' || character == 'E' || character == 'e' || character == '/') {
                return true;
            }
        }
        return false;
    }

    /**
     * Expands section pattern.
     * @param pattern pattern
     * @param valueText value text
     * @param replaceTextPlaceholder replace text placeholder
     * @return the computed result
     */
    public static String expandSectionPattern(String pattern, String valueText, boolean replaceTextPlaceholder) {
        StringBuilder builder = new StringBuilder(pattern.length() + valueText.length());
        boolean inQuote = false;

        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < pattern.length(); index++) {
            char character = pattern.charAt(index);
            if (character == '"') {
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote) {
                if (character == '@' && replaceTextPlaceholder) {
                    builder.append(valueText);
                    continue;
                }

                if (character == '_') {
                    index++;
                    continue;
                }

                if (character == '*') {
                    index++;
                    continue;
                }

                if (character == '\\') {
                    if (index + 1 < pattern.length()) {
                        index++;
                        builder.append(pattern.charAt(index));
                    }
                    continue;
                }
            }

            builder.append(character);
        }

        return builder.toString();
    }

    /**
     * Indicates whether numeric value.
     * @param value value to apply
     * @return true when the condition is satisfied
     */
    public static boolean isNumericValue(Object value) {
        return value instanceof Byte || value instanceof Short || value instanceof Integer ||
               value instanceof Long || value instanceof Float || value instanceof Double;
    }

    /**
     * Converts to double.
     * @param value value to apply
     * @return the computed result
     */
    public static double convertToDouble(Object value) {
        // Handle the relevant branch before the state changes.
        if (value instanceof Byte byteValue) {
            return byteValue;
        } else if (value instanceof Short shortValue) {
            return shortValue;
        } else if (value instanceof Integer intValue) {
            return intValue;
        } else if (value instanceof Long longValue) {
            return longValue;
        } else if (value instanceof Float floatValue) {
            return floatValue;
        } else if (value instanceof Double doubleValue) {
            return doubleValue;
        } else {
            throw new IllegalStateException("Value is not numeric.");
        }
    }

    /**
     * Processes get absolute numeric value.
     * @param value value to apply
     * @return the requested result
     */
    public static Object getAbsoluteNumericValue(Object value) {
        // Handle the relevant branch before the state changes.
        if (value instanceof Byte byteValue) {
            return byteValue;
        } else if (value instanceof Short shortValue) {
            return Math.abs(shortValue);
        } else if (value instanceof Integer intValue) {
            return Math.abs(intValue);
        } else if (value instanceof Long longValue) {
            return Math.abs(longValue);
        } else if (value instanceof Float floatValue) {
            return Math.abs(floatValue);
        } else if (value instanceof Double doubleValue) {
            return Math.abs(doubleValue);
        } else {
            return value;
        }
    }

    /**
     * Selects the appropriate numeric section based on the value.
     */
    public static DisplayFormatSectionInfo selectNumericSection(List<DisplayFormatSectionInfo> sections, double numericValue) {
        boolean[] useAbsoluteValue = {false};
        return selectNumericSectionImpl(sections, numericValue, useAbsoluteValue);
    }

    /**
     * Selects the appropriate numeric section and reports whether formatting should use the absolute value.
     * @param sections sections
     * @param numericValue numeric value
     * @param useAbsoluteValue use absolute value
     * @return the selected section
     */
    public static DisplayFormatSectionInfo selectNumericSection(List<DisplayFormatSectionInfo> sections,
                                                                double numericValue,
                                                                boolean[] useAbsoluteValue) {
        return selectNumericSectionImpl(sections, numericValue, useAbsoluteValue);
    }

    /**
     * Selects numeric section impl.
     * @param sections sections
     * @param numericValue numeric value
     * @param useAbsoluteValue use absolute value
     * @return the computed result
     */
    private static DisplayFormatSectionInfo selectNumericSectionImpl(List<DisplayFormatSectionInfo> sections,
                                                                       double numericValue,
                                                                       boolean[] useAbsoluteValue) {
        useAbsoluteValue[0] = false;
        int numericSectionCount = sections.size() < 3 ? sections.size() : 3;
        // Handle the relevant branch before the state changes.
        if (numericSectionCount == 0) {
            return null;
        }

        boolean hasCondition = false;
        DisplayFormatSectionInfo fallbackSection = null;
        int fallbackIndex = -1;
        for (int index = 0; index < numericSectionCount; index++) {
            DisplayFormatSectionInfo section = sections.get(index);
            if (section.getHasCondition()) {
                hasCondition = true;
                if (evaluateCondition(section.getConditionOperator(), numericValue, section.getConditionValue())) {
                    useAbsoluteValue[0] = numericValue < 0 && index > 0
                            && shouldUseAbsoluteValue(section.getRaw(), numericValue);
                    return section;
                }
            } else if (fallbackSection == null) {
                fallbackSection = section;
                fallbackIndex = index;
            }
        }

        if (hasCondition && fallbackSection != null) {
            useAbsoluteValue[0] = numericValue < 0 && fallbackIndex > 0
                    && shouldUseAbsoluteValue(fallbackSection.getRaw(), numericValue);
            return fallbackSection;
        }

        if (numericSectionCount == 1) {
            useAbsoluteValue[0] = false;
            return sections.get(0);
        }

        if (numericSectionCount == 2) {
            if (numericValue < 0) {
                useAbsoluteValue[0] = shouldUseAbsoluteValue(sections.get(1).getRaw(), numericValue);
                return sections.get(1);
            }
            useAbsoluteValue[0] = shouldUseAbsoluteValue(sections.get(0).getRaw(), numericValue);
            return sections.get(0);
        }

        if (numericValue > 0) {
            useAbsoluteValue[0] = shouldUseAbsoluteValue(sections.get(0).getRaw(), numericValue);
            return sections.get(0);
        }

        if (numericValue < 0) {
            useAbsoluteValue[0] = shouldUseAbsoluteValue(sections.get(1).getRaw(), numericValue);
            return sections.get(1);
        }

        useAbsoluteValue[0] = false;
        return sections.get(2);
    }

    /**
     * Selects the text section from the format sections.
     */
    public static DisplayFormatSectionInfo selectTextSection(List<DisplayFormatSectionInfo> sections) {
        // Handle the relevant branch before the state changes.
        if (sections.size() >= 4) {
            return sections.get(3);
        }

        for (DisplayFormatSectionInfo section : sections) {
            String sectionStripped = stripDirectiveBrackets(section.getRaw(), false);
            if (sectionStripped.indexOf('@') >= 0) {
                return section;
            }
        }

        return null;
    }

    /**
     * Selects the date/time section from the format sections.
     */
    public static DisplayFormatSectionInfo selectDateTimeSection(List<DisplayFormatSectionInfo> sections) {
        int dateSectionCount = sections.size() < 3 ? sections.size() : 3;
        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < dateSectionCount; index++) {
            String raw = sections.get(index).getRaw();
            if (raw != null && !raw.isBlank()) {
                return sections.get(index);
            }
        }

        return null;
    }
}

