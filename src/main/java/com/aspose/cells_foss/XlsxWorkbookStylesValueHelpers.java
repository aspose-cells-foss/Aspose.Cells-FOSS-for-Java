package com.aspose.cells_foss;

import com.aspose.cells_foss.core.BorderSideValue;
import com.aspose.cells_foss.core.BorderStyle;
import com.aspose.cells_foss.core.ColorValue;
import com.aspose.cells_foss.core.FontValue;
import com.aspose.cells_foss.core.HorizontalAlignment;
import com.aspose.cells_foss.core.VerticalAlignment;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.Locale;
import com.aspose.cells_foss.core.BordersValue;

/**
 * Provides helper methods for parsing and formatting workbook style values.
 * This is a package-internal utility class.
 */
public final class XlsxWorkbookStylesValueHelpers {

    /**
     * Initializes a new XlsxWorkbookStylesValueHelpers instance.
     */
    private XlsxWorkbookStylesValueHelpers() {}

    /**
     * Parses a border style string into a {@link BorderStyle} enum.
     *
     * @param value the string to parse (nullable)
     * @return the corresponding {@link BorderStyle}, or {@link BorderStyle#NONE} if not recognized
     */
    public static BorderStyle parseBorderStyle(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            return BorderStyle.NONE;
        }
        String lower = value.toLowerCase(Locale.ROOT);
        switch (lower) {
            case "thin":         return BorderStyle.THIN;
            case "medium":       return BorderStyle.MEDIUM;
            case "thick":        return BorderStyle.THICK;
            case "dotted":       return BorderStyle.DOTTED;
            case "dashed":       return BorderStyle.DASHED;
            case "double":       return BorderStyle.DOUBLE;
            case "hair":         return BorderStyle.HAIR;
            case "mediumdashed": return BorderStyle.MEDIUM_DASHED;
            case "dashdot":      return BorderStyle.DASH_DOT;
            case "mediumdashdot":return BorderStyle.MEDIUM_DASH_DOT;
            case "dashdotdot":   return BorderStyle.DASH_DOT_DOT;
            case "mediumdashdotdot":return BorderStyle.MEDIUM_DASH_DOT_DOT;
            case "slantdashdot": return BorderStyle.SLANTED_DASH_DOT;
            default:             return BorderStyle.NONE;
        }
    }

    /**
     * Gets the string representation of a {@link BorderStyle}.
     *
     * @param value the border style
     * @return the corresponding string (e.g. "thin", "medium"), or empty string if not recognized
     */
    public static String getBorderStyleName(BorderStyle value) {
        // Translate the internal value into the matching public representation.
        switch (value) {
            case THIN:              return "thin";
            case MEDIUM:            return "medium";
            case THICK:             return "thick";
            case DOTTED:            return "dotted";
            case DASHED:            return "dashed";
            case DOUBLE:            return "double";
            case HAIR:              return "hair";
            case MEDIUM_DASHED:     return "mediumDashed";
            case DASH_DOT:          return "dashDot";
            case MEDIUM_DASH_DOT:   return "mediumDashDot";
            case DASH_DOT_DOT:      return "dashDotDot";
            case MEDIUM_DASH_DOT_DOT: return "mediumDashDotDot";
            case SLANTED_DASH_DOT:  return "slantDashDot";
            default:                return "";
        }
    }

    /**
     * Gets the string representation of a {@link HorizontalAlignment}.
     *
     * @param value the horizontal alignment
     * @return the corresponding string (e.g. "left", "center"), or empty string if not recognized
     */
    public static String getHorizontalAlignmentName(HorizontalAlignment value) {
        // Translate the internal value into the matching public representation.
        switch (value) {
            case LEFT:               return "left";
            case CENTER:             return "center";
            case RIGHT:              return "right";
            case FILL:               return "fill";
            case JUSTIFY:            return "justify";
            case CENTER_CONTINUOUS:  return "centerContinuous";
            case DISTRIBUTED:        return "distributed";
            default:                 return "";
        }
    }

    /**
     * Gets the string representation of a {@link VerticalAlignment}.
     *
     * @param value the vertical alignment
     * @return the corresponding string (e.g. "center", "top"), or empty string if not recognized
     */
    public static String getVerticalAlignmentName(VerticalAlignment value) {
        // Translate the internal value into the matching public representation.
        switch (value) {
            case CENTER:       return "center";
            case TOP:          return "top";
            case JUSTIFY:      return "justify";
            case DISTRIBUTED:  return "distributed";
            default:           return "";
        }
    }

    /**
     * Converts a {@link ColorValue} to an ARGB hex string (e.g. "FFAABBCC").
     *
     * @param color the color value
     * @return the ARGB hex string
     */
    public static String toArgbHex(ColorValue color) {
        // Use DecimalFormat with two-digit padding and no grouping
        DecimalFormat df = new DecimalFormat("00", DecimalFormatSymbols.getInstance(Locale.ROOT));
        return new StringBuilder()
            .append(df.format(color.getA() & 0xFF))
            .append(df.format(color.getR() & 0xFF))
            .append(df.format(color.getG() & 0xFF))
            .append(df.format(color.getB() & 0xFF))
            .toString();
    }

    /**
     * Checks if a {@link ColorValue} is empty (all components are zero).
     *
     * @param color the color value
     * @return true if all ARGB components are zero, false otherwise
     */
    public static boolean isEmptyColor(ColorValue color) {
        return color.getA() == 0 && color.getR() == 0 &&
               color.getG() == 0 && color.getB() == 0;
    }

    /**
     * Compares two {@link FontValue} instances for equality.
     *
     * @param left  the first font value
     * @param right the second font value
     * @return true if all properties are equal, false otherwise
     */
    public static boolean fontEquals(FontValue left, FontValue right) {
        return java.util.Objects.equals(left.getName(), right.getName())
            && Double.compare(left.getSize(), right.getSize()) == 0
            && left.getBold() == right.getBold()
            && left.getItalic() == right.getItalic()
            && left.getUnderline() == right.getUnderline()
            && left.getStrikeThrough() == right.getStrikeThrough()
            && java.util.Objects.equals(left.getColor(), right.getColor());
    }

    /**
     * Compares two {@link BordersValue} instances for equality.
     *
     * @param left  the first borders value
     * @param right the second borders value
     * @return true if all border properties are equal, false otherwise
     */
    public static boolean bordersEqual(BordersValue left, BordersValue right) {
        return borderSideEquals(left.getLeft(), right.getLeft())
            && borderSideEquals(left.getRight(), right.getRight())
            && borderSideEquals(left.getTop(), right.getTop())
            && borderSideEquals(left.getBottom(), right.getBottom())
            && borderSideEquals(left.getDiagonal(), right.getDiagonal())
            && left.getDiagonalUp() == right.getDiagonalUp()
            && left.getDiagonalDown() == right.getDiagonalDown();
    }

    /**
     * Compares two {@link BorderSideValue} instances for equality.
     *
     * @param left  the first border side value
     * @param right the second border side value
     * @return true if style and color are equal, false otherwise
     */
    private static boolean borderSideEquals(BorderSideValue left, BorderSideValue right) {
        return left.getStyle() == right.getStyle() &&
               java.util.Objects.equals(left.getColor(), right.getColor());
    }
}