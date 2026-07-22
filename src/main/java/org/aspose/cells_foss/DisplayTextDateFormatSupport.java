package org.aspose.cells_foss;

import java.time.LocalDateTime;
import java.util.Locale;

/**
 * Utility class for formatting date/time values in display text.
 */
class DisplayTextDateFormatSupport {

    /**
     * Initializes a new DisplayTextDateFormatSupport instance.
     */
    private DisplayTextDateFormatSupport() {}

    /**
     * Formats a DateTime value using the specified format code.
     *
     * @param value      the date/time value
     * @param formatCode the format code string
     * @param culture    the culture for locale-specific formatting
     * @return the formatted string
     */
    static String formatDateTimeValue(LocalDateTime value, String formatCode, Locale culture) {
        StringBuilder builder = new StringBuilder(formatCode.length() + 16);
        boolean hasAmPm = formatCode.toLowerCase(culture).contains("am/pm")
                || formatCode.toLowerCase(culture).contains("a/p");
        boolean inQuote = false;

        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < formatCode.length(); index++) {
            char character = formatCode.charAt(index);

            if (matchesToken(formatCode, index, "AM/PM")) {
                builder.append(getAmPmDesignator(value, culture, false));
                index += 4;
                continue;
            }

            if (matchesToken(formatCode, index, "A/P")) {
                builder.append(getAmPmDesignator(value, culture, true));
                index += 2;
                continue;
            }

            if (character == '"') {
                inQuote = !inQuote;
                continue;
            }

            if (inQuote) {
                builder.append(character);
                continue;
            }

            if (character == '\\') {
                if (index + 1 < formatCode.length()) {
                    index++;
                    builder.append(formatCode.charAt(index));
                }
                continue;
            }

            if (character == '_' || character == '*') {
                index++;
                continue;
            }

            if (character == '.' && tryAppendFractionalSeconds(value, formatCode, culture, index, builder)) {
                index += 1; // advance for '.' already consumed
                continue;
            }

            if (character == 'y' || character == 'Y') {
                int count = countRepeated(formatCode, index, character);
                appendYear(value, culture, count, builder);
                index += count - 1;
                continue;
            }

            if (character == 'd' || character == 'D') {
                int count = countRepeated(formatCode, index, character);
                appendDay(value, culture, count, builder);
                index += count - 1;
                continue;
            }

            if (character == 'h' || character == 'H') {
                int count = countRepeated(formatCode, index, character);
                appendHour(value, culture, count, hasAmPm, builder);
                index += count - 1;
                continue;
            }

            if (character == 's' || character == 'S') {
                int count = countRepeated(formatCode, index, character);
                appendSecond(value, culture, count, builder);
                index += count - 1;
                continue;
            }

            if (character == 'm' || character == 'M') {
                int count = countRepeated(formatCode, index, character);
                if (isMinuteContext(formatCode, index, count)) {
                    appendMinute(value, culture, count, builder);
                } else {
                    appendMonth(value, culture, count, builder);
                }
                index += count - 1;
                continue;
            }

            builder.append(character);
        }

        return builder.toString();
    }

    /**
     * Checks if a token matches at the given index.
     */
    static boolean matchesToken(String formatCode, int startIndex, String token) {
        // Handle the relevant branch before the state changes.
        if (startIndex + token.length() > formatCode.length()) {
            return false;
        }

        return formatCode.regionMatches(true, startIndex, token, 0, token.length());
    }

    /**
     * Counts repeated occurrences of a character starting from startIndex.
     */
    static int countRepeated(String formatCode, int startIndex, char token) {
        int count = 1;
        // Walk the current collection so every entry is processed consistently.
        for (int index = startIndex + 1; index < formatCode.length(); index++) {
            if (Character.toLowerCase(formatCode.charAt(index)) != Character.toLowerCase(token)) {
                break;
            }
            count++;
        }
        return count;
    }

    /**
     * Determines if 'm' is in minute context rather than month context.
     */
    static boolean isMinuteContext(String formatCode, int startIndex, int count) {
        char previous = findNeighborToken(formatCode, startIndex - 1, -1);
        char next = findNeighborToken(formatCode, startIndex + count, 1);

        // Handle the relevant branch before the state changes.
        if (previous == ':' || next == ':') {
            return true;
        }

        if (previous == 'h' || previous == 'H' || previous == 's' || previous == 'S') {
            return true;
        }

        if (next == 'h' || next == 'H' || next == 's' || next == 'S') {
            return true;
        }

        return false;
    }

    /**
     * Finds the next/previous non-special character token.
     */
    static char findNeighborToken(String formatCode, int startIndex, int direction) {
        boolean inQuote = false;

        // Handle the relevant branch before the state changes.
        if (direction < 0) {
            for (int index = startIndex; index >= 0; index--) {
                char character = formatCode.charAt(index);
                if (character == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote) {
                    continue;
                }

                if (character == '\\' || character == '_' || character == '*') {
                    continue;
                }

                if (Character.isWhitespace(character)) {
                    continue;
                }

                return character;
            }
        } else {
            for (int index = startIndex; index < formatCode.length(); index++) {
                char character = formatCode.charAt(index);
                if (character == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote) {
                    continue;
                }

                if (character == '\\' || character == '_' || character == '*') {
                    continue;
                }

                if (Character.isWhitespace(character)) {
                    continue;
                }

                return character;
            }
        }

        return '\0';
    }

    /**
     * Checks if a token represents elapsed time.
     */
    static boolean isElapsedToken(String token) {
        String normalized = token.trim().toLowerCase(Locale.ROOT);
        return normalized.equals("h") || normalized.equals("hh")
                || normalized.equals("m") || normalized.equals("mm")
                || normalized.equals("s") || normalized.equals("ss");
    }

    /**
     * Checks if the format code contains elapsed time tokens in brackets.
     */
    static boolean containsElapsedTimeToken(String formatCode) {
        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < formatCode.length(); index++) {
            if (formatCode.charAt(index) != '[') {
                continue;
            }

            int endIndex = formatCode.indexOf(']', index + 1);
            if (endIndex < 0) {
                continue;
            }

            String token = formatCode.substring(index + 1, endIndex);
            if (isElapsedToken(token)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Formats a time span value using the specified format code.
     *
     * @param time       the time span value
     * @param formatCode the format code string
     * @param culture    the culture for locale-specific formatting
     * @return the formatted string
     */
    static String formatElapsedTimeValue(java.time.Duration time, String formatCode, Locale culture) {
        StringBuilder builder = new StringBuilder(formatCode.length() + 8);
        boolean inQuote = false;

        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < formatCode.length(); index++) {
            char character = formatCode.charAt(index);

            if (character == '"') {
                inQuote = !inQuote;
                continue;
            }

            if (inQuote) {
                builder.append(character);
                continue;
            }

            if (character == '[') {
                int endIndex = formatCode.indexOf(']', index + 1);
                if (endIndex > index) {
                    String token = formatCode.substring(index + 1, endIndex).toLowerCase(Locale.ROOT);

                    if (token.equals("h")) {
                        long hours = time.toHours();
                        builder.append(Long.toString(hours));
                    } else if (token.equals("hh")) {
                        long hours = time.toHours();
                        builder.append(String.format("%02d", hours));
                    } else if (token.equals("m")) {
                        long minutes = time.toMinutes() % 60;
                        builder.append(Long.toString(minutes));
                    } else if (token.equals("mm")) {
                        long minutes = time.toMinutes() % 60;
                        builder.append(String.format("%02d", minutes));
                    } else if (token.equals("s")) {
                        long seconds = time.getSeconds() % 60;
                        builder.append(Long.toString(seconds));
                    } else if (token.equals("ss")) {
                        long seconds = time.getSeconds() % 60;
                        builder.append(String.format("%02d", seconds));
                    } else {
                        builder.append('[');
                        builder.append(token);
                        builder.append(']');
                    }

                    index = endIndex;
                    continue;
                }
            }

            if (character == 'h' || character == 'H') {
                int count = countRepeated(formatCode, index, character);
                long hours = time.toHours() % 24;
                if (count == 1) {
                    builder.append(Long.toString(hours));
                } else {
                    builder.append(String.format("%02d", hours));
                }
                index += count - 1;
                continue;
            }

            if (character == 'm' || character == 'M') {
                int count = countRepeated(formatCode, index, character);
                long minutes = time.toMinutes() % 60;
                if (count == 1) {
                    builder.append(Long.toString(minutes));
                } else {
                    builder.append(String.format("%02d", minutes));
                }
                index += count - 1;
                continue;
            }

            if (character == 's' || character == 'S') {
                int count = countRepeated(formatCode, index, character);
                long seconds = time.getSeconds() % 60;
                if (count == 1) {
                    builder.append(Long.toString(seconds));
                } else {
                    builder.append(String.format("%02d", seconds));
                }
                index += count - 1;
                continue;
            }

            if (character == '.' && index + 1 < formatCode.length() && formatCode.charAt(index + 1) == '0') {
                int zeroCount = countRepeated(formatCode, index + 1, '0');
                builder.append('.');
                appendFractionDigits((int) time.toMillisPart(), zeroCount, builder);
                index += zeroCount;
                continue;
            }

            if (character == '\\') {
                if (index + 1 < formatCode.length()) {
                    index++;
                    builder.append(formatCode.charAt(index));
                }
                continue;
            }

            if (character == '_' || character == '*') {
                index++;
                continue;
            }

            builder.append(character);
        }

        return builder.toString();
    }

    /**
     * Appends year.
     * @param value value to apply
     * @param culture culture
     * @param count count
     * @param builder builder
     */
    private static void appendYear(LocalDateTime value, Locale culture, int count, StringBuilder builder) {
        int year = value.getYear();
        // Handle the relevant branch before the state changes.
        if (count <= 1) {
            builder.append(Integer.toString(year % 100));
            return;
        }

        if (count == 2) {
            builder.append(String.format("%02d", year % 100));
            return;
        }

        builder.append(String.format("%0" + count + "d", year));
    }

    /**
     * Appends day.
     * @param value value to apply
     * @param culture culture
     * @param count count
     * @param builder builder
     */
    private static void appendDay(LocalDateTime value, Locale culture, int count, StringBuilder builder) {
        // Handle the relevant branch before the state changes.
        if (count == 1) {
            builder.append(Integer.toString(value.getDayOfMonth()));
            return;
        }

        if (count == 2) {
            builder.append(String.format("%02d", value.getDayOfMonth()));
            return;
        }

        if (count == 3) {
            String[] abbrevDays = {"Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"};
            builder.append(abbrevDays[value.getDayOfWeek().getValue() - 1]);
            return;
        }

        String[] fullDays = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"};
        builder.append(fullDays[value.getDayOfWeek().getValue() - 1]);
    }

    /**
     * Appends month.
     * @param value value to apply
     * @param culture culture
     * @param count count
     * @param builder builder
     */
    private static void appendMonth(LocalDateTime value, Locale culture, int count, StringBuilder builder) {
        // Handle the relevant branch before the state changes.
        if (count == 1) {
            builder.append(Integer.toString(value.getMonthValue()));
            return;
        }

        if (count == 2) {
            builder.append(String.format("%02d", value.getMonthValue()));
            return;
        }

        if (count == 3) {
            String[] abbrevMonths = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
            builder.append(abbrevMonths[value.getMonthValue() - 1]);
            return;
        }

        String[] fullMonths = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
        String monthName = fullMonths[value.getMonthValue() - 1];
        if (count == 4) {
            builder.append(monthName);
            return;
        }

        if (monthName.isEmpty()) {
            builder.append(Integer.toString(value.getMonthValue()));
            return;
        }

        builder.append(monthName.charAt(0));
    }

    /**
     * Appends hour.
     * @param value value to apply
     * @param culture culture
     * @param count count
     * @param hasAmPm has am pm
     * @param builder builder
     */
    private static void appendHour(LocalDateTime value, Locale culture, int count, boolean hasAmPm, StringBuilder builder) {
        int hour = value.getHour();
        // Handle the relevant branch before the state changes.
        if (hasAmPm) {
            hour %= 12;
            if (hour == 0) {
                hour = 12;
            }
        }

        if (count == 1) {
            builder.append(Integer.toString(hour));
            return;
        }

        builder.append(String.format("%02d", hour));
    }

    /**
     * Appends minute.
     * @param value value to apply
     * @param culture culture
     * @param count count
     * @param builder builder
     */
    private static void appendMinute(LocalDateTime value, Locale culture, int count, StringBuilder builder) {
        // Handle the relevant branch before the state changes.
        if (count == 1) {
            builder.append(Integer.toString(value.getMinute()));
            return;
        }

        builder.append(String.format("%02d", value.getMinute()));
    }

    /**
     * Appends second.
     * @param value value to apply
     * @param culture culture
     * @param count count
     * @param builder builder
     */
    private static void appendSecond(LocalDateTime value, Locale culture, int count, StringBuilder builder) {
        // Handle the relevant branch before the state changes.
        if (count == 1) {
            builder.append(Integer.toString(value.getSecond()));
            return;
        }

        builder.append(String.format("%02d", value.getSecond()));
    }

    /**
     * Attempts to process append fractional seconds.
     * @param value value to apply
     * @param formatCode format code
     * @param culture culture
     * @param index index
     * @param builder builder
     * @return true when the condition is satisfied
     */
    private static boolean tryAppendFractionalSeconds(LocalDateTime value, String formatCode, Locale culture, int index, StringBuilder builder) {
        // Handle the relevant branch before the state changes.
        if (index + 1 >= formatCode.length() || formatCode.charAt(index + 1) != '0') {
            return false;
        }

        char previous = findNeighborToken(formatCode, index - 1, -1);
        if (previous != 's' && previous != 'S') {
            return false;
        }

        int zeroCount = countRepeated(formatCode, index + 1, '0');
        builder.append('.');
        appendFractionDigits(value.getNano() / 1000000, zeroCount, builder);
        return true;
    }

    /**
     * Appends fraction digits.
     * @param milliseconds milliseconds
     * @param zeroCount zero count
     * @param builder builder
     */
    private static void appendFractionDigits(int milliseconds, int zeroCount, StringBuilder builder) {
        // Handle the relevant branch before the state changes.
        if (zeroCount <= 0) {
            return;
        }

        if (zeroCount < 3) {
            int scale = 1;
            for (int index = zeroCount; index < 3; index++) {
                scale *= 10;
            }

            int rounded = (int) Math.round(milliseconds / (double) scale);
            int maxValue = 1;
            for (int index = 0; index < zeroCount; index++) {
                maxValue *= 10;
            }
            if (rounded >= maxValue) {
                rounded = maxValue - 1;
            }

            builder.append(String.format("%0" + zeroCount + "d", rounded));
            return;
        }

        String digits = String.format("%03d", milliseconds);
        builder.append(digits);
        for (int i = 3; i < zeroCount; i++) {
            builder.append('0');
        }
    }

    /**
     * Processes get am pm designator.
     * @param value value to apply
     * @param culture culture
     * @param abbreviated abbreviated
     * @return the requested result
     */
    private static String getAmPmDesignator(LocalDateTime value, Locale culture, boolean abbreviated) {
        String designator = value.getHour() < 12 ? "AM" : "PM";
        // Handle the relevant branch before the state changes.
        if (designator == null || designator.isEmpty()) {
            designator = value.getHour() < 12 ? "AM" : "PM";
        }

        if (!abbreviated) {
            return designator;
        }

        return designator.substring(0, 1);
    }
}

