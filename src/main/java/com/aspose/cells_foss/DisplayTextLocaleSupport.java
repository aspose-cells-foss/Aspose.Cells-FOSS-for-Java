package com.aspose.cells_foss;

import java.util.Locale;

/**
 * Provides locale support for display text formatting in Excel files.
 */
final class DisplayTextLocaleSupport {

    /**
     * Prevents instantiation of this utility class.
     */
    private DisplayTextLocaleSupport() {}

    /**
     * Applies locale directives to a format string.
     *
     * @param section the input format string
     * @param fallbackCulture the fallback culture to use
     * @param sectionCulture output parameter for the resolved culture
     * @return the processed format string with directives replaced
     */
    static String applyLocaleDirectives(String section, Locale fallbackCulture, Locale[] sectionCulture) {
        sectionCulture[0] = fallbackCulture;
        StringBuilder builder = new StringBuilder(section.length() + 16);
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
                if (endIndex > index) {
                    String token = section.substring(index + 1, endIndex);
                    String[] replacementHolder = {""};
                    Locale[] resolvedCultureHolder = {fallbackCulture};
                    if (tryResolveLocaleDirective(token, fallbackCulture, replacementHolder, resolvedCultureHolder)) {
                        builder.append(replacementHolder[0]);
                        sectionCulture[0] = resolvedCultureHolder[0];
                        index = endIndex;
                        continue;
                    }
                }
            }

            builder.append(character);
        }

        return builder.toString();
    }

    /**
     * Attempts to process resolve locale directive.
     * @param token token
     * @param fallbackCulture fallback culture
     * @param replacement replacement
     * @param resolvedCulture resolved culture
     * @return true when the condition is satisfied
     */
    private static boolean tryResolveLocaleDirective(String token, Locale fallbackCulture,
                                                     String[] replacement, Locale[] resolvedCulture) {
        replacement[0] = "";
        resolvedCulture[0] = fallbackCulture;
        // Handle the relevant branch before the state changes.
        if (token == null || token.isBlank() || token.charAt(0) != '$') {
            return false;
        }

        int dashIndex = token.lastIndexOf('-');
        if (dashIndex <= 0 || dashIndex >= token.length() - 1) {
            return false;
        }

        String symbol = token.substring(1, dashIndex);
        String localeCode = token.substring(dashIndex + 1);
        resolvedCulture[0] = resolveCulture(localeCode, fallbackCulture);

        if ("F800".equalsIgnoreCase(localeCode)) {
            replacement[0] = quoteLiteral(fallbackCulture.getDisplayName());
            resolvedCulture[0] = fallbackCulture;
            return true;
        }

        if ("F400".equalsIgnoreCase(localeCode)) {
            replacement[0] = fallbackCulture.getDisplayName();
            resolvedCulture[0] = fallbackCulture;
            return true;
        }

        if (!symbol.isEmpty()) {
            replacement[0] = quoteLiteral(symbol);
        }

        return true;
    }

    /**
     * Resolves culture.
     * @param localeCode locale code
     * @param fallbackCulture fallback culture
     * @return the computed result
     */
    private static Locale resolveCulture(String localeCode, Locale fallbackCulture) {
        // Handle the relevant branch before the state changes.
        if (localeCode == null || localeCode.isBlank()) {
            return fallbackCulture;
        }

        try {
            int lcid = Integer.parseInt(localeCode, 16);
            return getCultureInfo(lcid);
        } catch (NumberFormatException e) {
            return fallbackCulture;
        }
    }

    /**
     * Processes get culture info.
     * @param lcid lcid
     * @return the requested result
     */
    private static Locale getCultureInfo(int lcid) {
        // Map LCID to Locale - simplified implementation
        // In a full implementation, this would map all LCIDs to appropriate locales
        switch (lcid) {
            case 0x0409: return Locale.US;
            case 0x040C: return Locale.FRANCE;
            case 0x0407: return Locale.GERMANY;
            case 0x0410: return Locale.ITALY;
            case 0x0411: return Locale.JAPAN;
            case 0x0412: return Locale.KOREA;
            case 0x0413: return new Locale("nl", "NL");
            case 0x0415: return new Locale("pl", "PL");
            case 0x0416: return new Locale("pt", "BR");
            case 0x0419: return new Locale("ru", "RU");
            case 0x041D: return new Locale("sv", "SE");
            case 0x041F: return new Locale("tr", "TR");
            case 0x0804: return new Locale("zh", "CN");
            case 0x0816: return new Locale("pt", "PT");
            case 0x1009: return new Locale("en", "AU");
            case 0x2009: return new Locale("en", "GB");
            default: return Locale.US;
        }
    }

    /**
     * Quotes literal.
     * @param value value to apply
     * @return the computed result
     */
    private static String quoteLiteral(String value) {
        StringBuilder builder = new StringBuilder(value.length() + 2);
        builder.append('"');
        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < value.length(); index++) {
            char c = value.charAt(index);
            if (c == '"') {
                builder.append('"');
            }
            builder.append(c);
        }
        builder.append('"');
        return builder.toString();
    }
}