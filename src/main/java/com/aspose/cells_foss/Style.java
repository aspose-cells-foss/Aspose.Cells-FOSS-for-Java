package com.aspose.cells_foss;

import com.aspose.cells_foss.core.AlignmentValue;
import com.aspose.cells_foss.core.FillPatternKind;
import com.aspose.cells_foss.core.HorizontalAlignment;
import com.aspose.cells_foss.core.NumberFormatValue;
import com.aspose.cells_foss.core.ProtectionValue;
import com.aspose.cells_foss.core.StyleValue;
import com.aspose.cells_foss.core.VerticalAlignment;

/**
 * Represents the full style of a cell: font, borders, alignment, fill, number format, and protection.
 *
 * <p>Obtain a style snapshot via {@link Cell#getStyle()}, modify it, then apply it back via
 * {@link Cell#setStyle(Style)}.
 *
 * <p>Maps to C# {@code Aspose.Cells_FOSS.Style}.
 */
public class Style {

    private Font font;
    private Borders borders;

    // Alignment (flattened from AlignmentValue)
    private HorizontalAlignmentType horizontalAlignment = HorizontalAlignmentType.GENERAL;
    private VerticalAlignmentType verticalAlignment = VerticalAlignmentType.BOTTOM;
    private boolean wrapText;
    private int indentLevel;
    private int textRotation;
    private boolean shrinkToFit;
    private int readingOrder;
    private int relativeIndent;

    // Fill
    private FillPattern pattern = FillPattern.NONE;
    private Color foregroundColor = Color.getEmpty();
    private Color backgroundColor = Color.getEmpty();

    // Number format
    private int number;
    private String custom;

    // Cell protection
    private boolean isLocked = true;
    private boolean isHidden;

    /** Creates a default style (Calibri 11pt, no borders, general alignment, locked). */
    public Style() {
        this.font = new Font();
        this.borders = new Borders();
    }

    // -------------------------------------------------------------------------
    // Font
    // -------------------------------------------------------------------------

    /**
     * Returns the font.
     * @return the requested result
     */
    public Font getFont() { return font; }
    /**
     * Sets the font.
     * @param font font
     */
    public void setFont(Font font) {
        java.util.Objects.requireNonNull(font, "font");
        this.font = font;
    }

    // -------------------------------------------------------------------------
    // Borders
    // -------------------------------------------------------------------------

    /**
     * Returns the borders.
     * @return the requested result
     */
    public Borders getBorders() { return borders; }
    /**
     * Sets the borders.
     * @param borders borders
     */
    public void setBorders(Borders borders) {
        java.util.Objects.requireNonNull(borders, "borders");
        this.borders = borders;
    }

    // -------------------------------------------------------------------------
    // Alignment
    // -------------------------------------------------------------------------

    /**
     * Returns the horizontal alignment.
     * @return the requested result
     */
    public HorizontalAlignmentType getHorizontalAlignment() { return horizontalAlignment; }
    /**
     * Sets the horizontal alignment.
     * @param value value to apply
     */
    public void setHorizontalAlignment(HorizontalAlignmentType value) {
        this.horizontalAlignment = value != null ? value : HorizontalAlignmentType.GENERAL;
    }

    /**
     * Returns the vertical alignment.
     * @return the requested result
     */
    public VerticalAlignmentType getVerticalAlignment() { return verticalAlignment; }
    /**
     * Sets the vertical alignment.
     * @param value value to apply
     */
    public void setVerticalAlignment(VerticalAlignmentType value) {
        this.verticalAlignment = value != null ? value : VerticalAlignmentType.BOTTOM;
    }

    /**
     * Indicates whether wrap text.
     * @return true when the condition is satisfied
     */
    public boolean isTextWrapped() { return wrapText; }
    /**
     * Sets the text wrapped.
     * @param value value to apply
     */
    public void setTextWrapped(boolean value) { this.wrapText = value; }

    /**
     * Returns the indent level.
     * @return the requested result
     */
    public int getIndentLevel() { return indentLevel; }
    /**
     * Sets the indent level.
     * @param value value to apply
     */
    public void setIndentLevel(int value) {
        if (value < 0) throw new CellsException("IndentLevel must be non-negative.");
        this.indentLevel = value;
    }

    /**
     * Returns the text rotation.
     * @return the requested result
     */
    public int getRotationAngle() { return textRotation; }
    /**
     * Sets the rotation angle.
     * @param value value to apply
     */
    public void setRotationAngle(int value) {
        if (value < 0 || value > 180) throw new CellsException("RotationAngle must be between 0 and 180.");
        this.textRotation = value;
    }

    /**
     * Indicates whether shrink to fit.
     * @return true when the condition is satisfied
     */
    public boolean getShrinkToFit() { return shrinkToFit; }
    /**
     * Sets the shrink to fit.
     * @param value value to apply
     */
    public void setShrinkToFit(boolean value) { this.shrinkToFit = value; }

    /**
     * Returns the reading order.
     * @return the requested result
     */
    public int getReadingOrder() { return readingOrder; }
    /**
     * Sets the reading order.
     * @param value value to apply
     */
    public void setReadingOrder(int value) {
        if (value < 0 || value > 2) throw new CellsException("ReadingOrder must be 0, 1, or 2.");
        this.readingOrder = value;
    }

    /**
     * Returns the relative indent.
     * @return the requested result
     */
    public int getRelativeIndent() { return relativeIndent; }
    /**
     * Sets the relative indent.
     * @param value value to apply
     */
    public void setRelativeIndent(int value) { this.relativeIndent = value; }

    // -------------------------------------------------------------------------
    // Fill
    // -------------------------------------------------------------------------

    /**
     * Returns the pattern.
     * @return the requested result
     */
    public FillPattern getPattern() { return pattern; }
    /**
     * Sets the pattern.
     * @param value value to apply
     */
    public void setPattern(FillPattern value) { this.pattern = value != null ? value : FillPattern.NONE; }

    /**
     * Returns the foreground color.
     * @return the requested result
     */
    public Color getForegroundColor() { return foregroundColor; }
    /**
     * Sets the foreground color.
     * @param value value to apply
     */
    public void setForegroundColor(Color value) { this.foregroundColor = value != null ? value : Color.getEmpty(); }

    /**
     * Returns the background color.
     * @return the requested result
     */
    public Color getBackgroundColor() { return backgroundColor; }
    /**
     * Sets the background color.
     * @param value value to apply
     */
    public void setBackgroundColor(Color value) { this.backgroundColor = value != null ? value : Color.getEmpty(); }

    // -------------------------------------------------------------------------
    // Number format
    // -------------------------------------------------------------------------

    /** Built-in number format index (0 = General). Maps to ECMA-376 numFmtId for built-in formats. */
    public int getNumber() { return number; }
    /**
     * Sets the number.
     * @param value value to apply
     */
    public void setNumber(int value) { this.number = value; }

    /** Custom format string (e.g. {@code "#,##0.00"}, {@code "yyyy-mm-dd"}). */
    public String getCustom() { return custom; }
    /**
     * Sets the custom.
     * @param value value to apply
     */
    public void setCustom(String value) { this.custom = value; }

    /**
     * Combined number format getter: returns the custom format string if set,
     * otherwise returns the built-in format string for the current {@code number} index.
     */
    public String getNumberFormat() {
        return NumberFormat.resolveFormatCode(number, custom);
    }

    /**
     * Combined number format setter: if {@code formatCode} matches a built-in format,
     * sets {@link #number} to its ID and clears {@link #custom};
     * otherwise sets {@link #number} to 0 and stores the code in {@link #custom}.
     */
    public void setNumberFormat(String formatCode) {
        if (formatCode == null || formatCode.isBlank()) {
            number = 0;
            custom = null;
            return;
        }
        Integer id = NumberFormat.getBuiltInFormatId(formatCode.trim());
        if (id != null) {
            number = id;
            custom = null;
        } else {
            number = 0;
            custom = formatCode.trim();
        }
    }

    // -------------------------------------------------------------------------
    // Cell protection
    // -------------------------------------------------------------------------

    /** Whether the cell is locked when the sheet is protected. Default: {@code true}. */
    public boolean isLocked() { return isLocked; }
    /**
     * Sets the locked.
     * @param value value to apply
     */
    public void setLocked(boolean value) { this.isLocked = value; }

    /** Whether the cell formula is hidden when the sheet is protected. */
    public boolean isFormulaHidden() { return isHidden; }
    /**
     * Sets whether the cell formula is hidden.
     * @param value value to apply
     */
    public void setFormulaHidden(boolean value) { this.isHidden = value; }

    // -------------------------------------------------------------------------
    // Clone
    // -------------------------------------------------------------------------

    /**
     * Creates a copy of this instance.
     * @return the computed result
     */
    public Style clone() {
        Style c = new Style();
        c.font = this.font.clone();
        c.borders = this.borders.clone();
        c.horizontalAlignment = this.horizontalAlignment;
        c.verticalAlignment = this.verticalAlignment;
        c.wrapText = this.wrapText;
        c.indentLevel = this.indentLevel;
        c.textRotation = this.textRotation;
        c.shrinkToFit = this.shrinkToFit;
        c.readingOrder = this.readingOrder;
        c.relativeIndent = this.relativeIndent;
        c.pattern = this.pattern;
        c.foregroundColor = this.foregroundColor;
        c.backgroundColor = this.backgroundColor;
        c.number = this.number;
        c.custom = this.custom;
        c.isLocked = this.isLocked;
        c.isHidden = this.isHidden;
        return c;
    }

    // -------------------------------------------------------------------------
    // Model conversion (package-internal)
    // -------------------------------------------------------------------------

    /** Creates a {@code Style} from an internal {@code StyleValue} model. */
    static Style fromModel(StyleValue sv) {
        Style s = new Style();
        s.font = Font.fromModel(sv.getFont());
        s.borders = Borders.fromModel(sv.getBorders());

        // Alignment
        AlignmentValue a = sv.getAlignment();
        s.horizontalAlignment = mapHAlign(a.getHorizontal());
        s.verticalAlignment = mapVAlign(a.getVertical());
        s.wrapText = a.getWrapText();
        s.indentLevel = a.getIndentLevel();
        s.textRotation = a.getTextRotation();
        s.shrinkToFit = a.getShrinkToFit();
        s.readingOrder = a.getReadingOrder();
        s.relativeIndent = a.getRelativeIndent();

        // Fill
        s.pattern = sv.getPattern() != null ? mapPattern(sv.getPattern()) : FillPattern.NONE;
        s.foregroundColor = sv.getForegroundColor() != null
                ? Color.fromCore(sv.getForegroundColor()) : Color.getEmpty();
        s.backgroundColor = sv.getBackgroundColor() != null
                ? Color.fromCore(sv.getBackgroundColor()) : Color.getEmpty();

        // Number format — if a custom code is set, the public number stays 0
        NumberFormatValue nf = sv.getNumberFormat();
        s.custom = nf.getCustom();
        s.number = (s.custom != null && !s.custom.isEmpty()) ? 0 : nf.getNumber();

        // Protection
        ProtectionValue pv = sv.getProtection();
        s.isLocked = pv.getIsLocked();
        s.isHidden = pv.getIsHidden();

        return s;
    }

    /** Converts this {@code Style} to an internal {@code StyleValue} model. */
    StyleValue toModel() {
        StyleValue sv = new StyleValue();

        sv.setFont(font.toModel());
        sv.setBorders(borders.toModel());

        // Alignment
        AlignmentValue a = new AlignmentValue();
        a.setHorizontal(mapHAlignToCore(horizontalAlignment));
        a.setVertical(mapVAlignToCore(verticalAlignment));
        a.setWrapText(wrapText);
        a.setIndentLevel(indentLevel);
        a.setTextRotation(textRotation);
        a.setShrinkToFit(shrinkToFit);
        a.setReadingOrder(readingOrder);
        a.setRelativeIndent(relativeIndent);
        sv.setAlignment(a);

        // Fill
        sv.setPattern(pattern != FillPattern.NONE ? mapPatternToCore(pattern) : null);
        sv.setForegroundColor(foregroundColor.equals(Color.getEmpty()) ? null : foregroundColor.toCore());
        sv.setBackgroundColor(backgroundColor.equals(Color.getEmpty()) ? null : backgroundColor.toCore());

        // Number format
        NumberFormatValue nf = new NumberFormatValue();
        nf.setNumber(number);
        nf.setCustom(custom);
        sv.setNumberFormat(nf);

        // Protection
        ProtectionValue pv = new ProtectionValue();
        pv.setIsLocked(isLocked);
        pv.setIsHidden(isHidden);
        sv.setProtection(pv);

        return sv;
    }

    // -------------------------------------------------------------------------
    // Enum mapping helpers
    // -------------------------------------------------------------------------

    /**
     * Maps h align.
     * @param h h
     * @return the computed result
     */
    private static HorizontalAlignmentType mapHAlign(HorizontalAlignment h) {
        // Handle the relevant branch before the state changes.
        if (h == null) return HorizontalAlignmentType.GENERAL;
        switch (h) {
            case LEFT:              return HorizontalAlignmentType.LEFT;
            case CENTER:            return HorizontalAlignmentType.CENTER;
            case RIGHT:             return HorizontalAlignmentType.RIGHT;
            case FILL:              return HorizontalAlignmentType.FILL;
            case JUSTIFY:           return HorizontalAlignmentType.JUSTIFY;
            case CENTER_CONTINUOUS: return HorizontalAlignmentType.CENTER_CONTINUOUS;
            case DISTRIBUTED:       return HorizontalAlignmentType.DISTRIBUTED;
            default:                return HorizontalAlignmentType.GENERAL;
        }
    }

    /**
     * Maps h align to core.
     * @param h h
     * @return the computed result
     */
    private static HorizontalAlignment mapHAlignToCore(HorizontalAlignmentType h) {
        // Handle the relevant branch before the state changes.
        if (h == null) return HorizontalAlignment.GENERAL;
        switch (h) {
            case LEFT:              return HorizontalAlignment.LEFT;
            case CENTER:            return HorizontalAlignment.CENTER;
            case RIGHT:             return HorizontalAlignment.RIGHT;
            case FILL:              return HorizontalAlignment.FILL;
            case JUSTIFY:           return HorizontalAlignment.JUSTIFY;
            case CENTER_CONTINUOUS: return HorizontalAlignment.CENTER_CONTINUOUS;
            case DISTRIBUTED:       return HorizontalAlignment.DISTRIBUTED;
            default:                return HorizontalAlignment.GENERAL;
        }
    }

    /**
     * Maps v align.
     * @param v v
     * @return the computed result
     */
    private static VerticalAlignmentType mapVAlign(VerticalAlignment v) {
        // Handle the relevant branch before the state changes.
        if (v == null) return VerticalAlignmentType.BOTTOM;
        switch (v) {
            case CENTER:      return VerticalAlignmentType.CENTER;
            case TOP:         return VerticalAlignmentType.TOP;
            case JUSTIFY:     return VerticalAlignmentType.JUSTIFY;
            case DISTRIBUTED: return VerticalAlignmentType.DISTRIBUTED;
            default:          return VerticalAlignmentType.BOTTOM;
        }
    }

    /**
     * Maps v align to core.
     * @param v v
     * @return the computed result
     */
    private static VerticalAlignment mapVAlignToCore(VerticalAlignmentType v) {
        // Handle the relevant branch before the state changes.
        if (v == null) return VerticalAlignment.BOTTOM;
        switch (v) {
            case CENTER:      return VerticalAlignment.CENTER;
            case TOP:         return VerticalAlignment.TOP;
            case JUSTIFY:     return VerticalAlignment.JUSTIFY;
            case DISTRIBUTED: return VerticalAlignment.DISTRIBUTED;
            default:          return VerticalAlignment.BOTTOM;
        }
    }

    /**
     * Maps pattern.
     * @param k k
     * @return the computed result
     */
    private static FillPattern mapPattern(FillPatternKind k) {
        // Translate the internal value into the matching public representation.
        switch (k) {
            case SOLID:            return FillPattern.SOLID;
            case MEDIUM_GRAY:      return FillPattern.MEDIUM_GRAY;
            case DARK_GRAY:        return FillPattern.DARK_GRAY;
            case GRAY_125:         return FillPattern.GRAY_125;
            case GRAY_0625:        return FillPattern.GRAY_0625;
            case DARK_HORIZONTAL:  return FillPattern.DARK_HORIZONTAL;
            case DARK_VERTICAL:    return FillPattern.DARK_VERTICAL;
            case DARK_DOWN:        return FillPattern.DARK_DOWN;
            case DARK_UP:          return FillPattern.DARK_UP;
            case DARK_GRID:        return FillPattern.DARK_GRID;
            case DARK_TRELLIS:     return FillPattern.DARK_TRELLIS;
            case LIGHT_HORIZONTAL: return FillPattern.LIGHT_HORIZONTAL;
            case LIGHT_VERTICAL:   return FillPattern.LIGHT_VERTICAL;
            case LIGHT_DOWN:       return FillPattern.LIGHT_DOWN;
            case LIGHT_UP:         return FillPattern.LIGHT_UP;
            case LIGHT_GRID:       return FillPattern.LIGHT_GRID;
            case LIGHT_TRELLIS:    return FillPattern.LIGHT_TRELLIS;
            default:               return FillPattern.NONE;
        }
    }

    /**
     * Maps pattern to core.
     * @param p p
     * @return the computed result
     */
    private static FillPatternKind mapPatternToCore(FillPattern p) {
        // Translate the internal value into the matching public representation.
        switch (p) {
            case SOLID:            return FillPatternKind.SOLID;
            case MEDIUM_GRAY:      return FillPatternKind.MEDIUM_GRAY;
            case DARK_GRAY:        return FillPatternKind.DARK_GRAY;
            case GRAY_125:         return FillPatternKind.GRAY_125;
            case GRAY_0625:        return FillPatternKind.GRAY_0625;
            case DARK_HORIZONTAL:  return FillPatternKind.DARK_HORIZONTAL;
            case DARK_VERTICAL:    return FillPatternKind.DARK_VERTICAL;
            case DARK_DOWN:        return FillPatternKind.DARK_DOWN;
            case DARK_UP:          return FillPatternKind.DARK_UP;
            case DARK_GRID:        return FillPatternKind.DARK_GRID;
            case DARK_TRELLIS:     return FillPatternKind.DARK_TRELLIS;
            case LIGHT_HORIZONTAL: return FillPatternKind.LIGHT_HORIZONTAL;
            case LIGHT_VERTICAL:   return FillPatternKind.LIGHT_VERTICAL;
            case LIGHT_DOWN:       return FillPatternKind.LIGHT_DOWN;
            case LIGHT_UP:         return FillPatternKind.LIGHT_UP;
            case LIGHT_GRID:       return FillPatternKind.LIGHT_GRID;
            case LIGHT_TRELLIS:    return FillPatternKind.LIGHT_TRELLIS;
            default:               return FillPatternKind.NONE;
        }
    }
}
