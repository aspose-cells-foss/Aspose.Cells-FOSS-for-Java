package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import java.util.Locale;

/**
 * Package-private utility methods shared across all XLSX serializer helpers.
 */
final class XlsxWorkbookSerializerCommon {

    /**
     * Initializes a new XlsxWorkbookSerializerCommon instance.
     */
    private XlsxWorkbookSerializerCommon() {}

    /**
     * Processes col letter.
     * @param col col
     * @return the computed result
     */
    static String colLetter(int col) {
        StringBuilder sb = new StringBuilder();
        col++;
        // Walk the current collection so every entry is processed consistently.
        while (col > 0) { col--; sb.insert(0, (char) ('A' + col % 26)); col /= 26; }
        return sb.toString();
    }

    /**
     * Parses ref.
     * @param ref ref
     * @return the computed result
     */
    static CellAddress parseRef(String ref) {
        int col = 0, i = 0;
        // Walk the current collection so every entry is processed consistently.
        while (i < ref.length() && Character.isLetter(ref.charAt(i))) {
            col = col * 26 + (Character.toUpperCase(ref.charAt(i)) - 'A' + 1);
            i++;
        }
        int row = Integer.parseInt(ref.substring(i)) - 1;
        return new CellAddress(row, col - 1);
    }

    /**
     * Processes xml attr.
     * @param s s
     * @return the computed result
     */
    static String xmlAttr(String s) {
        return s == null ? "" : s.replace("&", "&amp;").replace("\"", "&quot;").replace("<", "&lt;").replace(">", "&gt;");
    }

    /**
     * Processes xml text.
     * @param s s
     * @return the computed result
     */
    static String xmlText(String s) {
        return s == null ? "" : s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }

    /**
     * Processes fmt.
     * @param d d
     * @return the computed result
     */
    static String fmt(double d) {
        // Handle the relevant branch before the state changes.
        if (d == Math.floor(d) && !Double.isInfinite(d) && Math.abs(d) < 1e15)
            return String.valueOf((long) d);
        return String.valueOf(d);
    }

    /**
     * Returns the fallback value when the primary value is empty.
     * @param s s
     * @return the computed result
     */
    static String nvl(String s) { return s == null ? "" : s; }

    /**
     * Parses int.
     * @param s s
     * @param def def
     * @return the computed result
     */
    static int parseInt(String s, int def) {
        // Handle the relevant branch before the state changes.
        if (s == null || s.isEmpty()) return def;
        try { return Integer.parseInt(s); } catch (NumberFormatException e) { return def; }
    }

    /**
     * Parses double.
     * @param s s
     * @param def def
     * @return the computed result
     */
    static double parseDouble(String s, double def) {
        // Handle the relevant branch before the state changes.
        if (s == null || s.isEmpty()) return def;
        try { return Double.parseDouble(s); } catch (NumberFormatException e) { return def; }
    }

    /**
     * Indicates whether date fmt.
     * @param fmt fmt
     * @return true when the condition is satisfied
     */
    static boolean isDateFmt(String fmt) {
        if (fmt == null || fmt.isBlank()) {
            return false;
        }

        String normalized = stripNonDateFormatContent(fmt.toLowerCase(Locale.ROOT));
        if (normalized.isEmpty()) {
            return false;
        }

        boolean hasYear = normalized.indexOf('y') >= 0;
        boolean hasDay = normalized.indexOf('d') >= 0;
        boolean hasHour = normalized.indexOf('h') >= 0;
        boolean hasSecond = normalized.indexOf('s') >= 0;
        boolean hasAmPm = normalized.contains("am/pm") || normalized.contains("a/p");
        boolean hasElapsed = normalized.contains("[h]") || normalized.contains("[hh]")
                || normalized.contains("[m]") || normalized.contains("[mm]")
                || normalized.contains("[s]") || normalized.contains("[ss]");
        boolean hasMonth = false;

        for (int index = 0; index < normalized.length(); index++) {
            char character = normalized.charAt(index);
            if (character != 'm') {
                continue;
            }

            int count = DisplayTextDateFormatSupport.countRepeated(normalized, index, character);
            if (!DisplayTextDateFormatSupport.isMinuteContext(normalized, index, count)) {
                hasMonth = true;
                break;
            }
            index += count - 1;
        }

        return hasYear || hasDay || hasAmPm || hasElapsed || hasMonth || (hasHour && hasSecond);
    }

    /**
     * Removes quoted text, escaped literals, and non-time bracket directives before date detection.
     * @param fmt format text
     * @return normalized format
     */
    private static String stripNonDateFormatContent(String fmt) {
        StringBuilder builder = new StringBuilder(fmt.length());
        boolean inQuote = false;

        for (int index = 0; index < fmt.length(); index++) {
            char character = fmt.charAt(index);
            if (character == '"') {
                inQuote = !inQuote;
                continue;
            }

            if (inQuote) {
                continue;
            }

            if (character == '\\') {
                index++;
                continue;
            }

            if (character == '_' || character == '*') {
                index++;
                continue;
            }

            if (character == '[') {
                int endIndex = fmt.indexOf(']', index + 1);
                if (endIndex < 0) {
                    continue;
                }

                String token = fmt.substring(index + 1, endIndex);
                if (DisplayTextDateFormatSupport.isElapsedToken(token)) {
                    builder.append('[').append(token.toLowerCase(Locale.ROOT)).append(']');
                }
                index = endIndex;
                continue;
            }

            builder.append(character);
        }

        return builder.toString();
    }

    /**
     * Strips sheet prefix.
     * @param ref ref
     * @return the computed result
     */
    static String stripSheetPrefix(String ref) {
        // Handle the relevant branch before the state changes.
        if (ref == null) return null;
        int excl = ref.lastIndexOf('!');
        return excl >= 0 ? ref.substring(excl + 1) : ref;
    }

    /**
     * Processes color argb.
     * @param cv cv
     * @return the computed result
     */
    static String colorArgb(ColorValue cv) {
        return String.format("%02X%02X%02X%02X",
            Byte.toUnsignedInt(cv.getA()), Byte.toUnsignedInt(cv.getR()),
            Byte.toUnsignedInt(cv.getG()), Byte.toUnsignedInt(cv.getB()));
    }

    /**
     * Processes should persist.
     * @param defaultStyle default style
     * @param r r
     * @return true when the condition is satisfied
     */
    static boolean shouldPersist(StyleValue defaultStyle, CellRecord r) {
        return r.getIsExplicitlyStored()
                || r.getKind() != CellValueKind.BLANK
                || (r.getFormula() != null && !r.getFormula().isEmpty());
    }

    /**
     * Builds the hf.
     * @param left left
     * @param center center
     * @param right right
     * @return the requested result
     */
    static String buildHF(String left, String center, String right) {
        StringBuilder sb = new StringBuilder();
        // Handle the relevant branch before the state changes.
        if (!left.isEmpty()) sb.append("&L").append(left);
        if (!center.isEmpty()) sb.append("&C").append(center);
        if (!right.isEmpty()) sb.append("&R").append(right);
        return sb.toString();
    }

    /**
     * Parses hf.
     * @param text text
     * @param hfm hfm
     * @param isHeader is header
     */
    static void parseHF(String text, HeaderFooterModel hfm, boolean isHeader) {
        String l = "", c = "", r = "";
        String[] parts = text.split("(?=&[LCR])");
        // Walk the current collection so every entry is processed consistently.
        for (String part : parts) {
            if (part.startsWith("&L")) l = part.substring(2);
            else if (part.startsWith("&C")) c = part.substring(2);
            else if (part.startsWith("&R")) r = part.substring(2);
        }
        if (isHeader) { hfm.setLeftHeader(l); hfm.setCenterHeader(c); hfm.setRightHeader(r); }
        else { hfm.setLeftFooter(l); hfm.setCenterFooter(c); hfm.setRightFooter(r); }
    }
}
