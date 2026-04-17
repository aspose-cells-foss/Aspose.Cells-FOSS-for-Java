package com.aspose.cells_foss.core;

/**
 * Represents a style value with various formatting properties.
 */
public final class StyleValue {
    private FontValue font = new FontValue();
    private FillPatternKind pattern;
    private ColorValue foregroundColor;
    private ColorValue backgroundColor;
    private BordersValue borders = new BordersValue();
    private AlignmentValue alignment = new AlignmentValue();
    private ProtectionValue protection = new ProtectionValue();
    private NumberFormatValue numberFormat = new NumberFormatValue();

    /**
     * Gets the default style value.
     */
    public static StyleValue getDefault() {
        return new StyleValue();
    }

    /**
     * Returns the font.
     * @return the requested result
     */
    public FontValue getFont() { return font; }
    /**
     * Sets the font.
     * @param font font
     */
    public void setFont(FontValue font) { this.font = font; }

    /**
     * Returns the pattern.
     * @return the requested result
     */
    public FillPatternKind getPattern() { return pattern; }
    /**
     * Sets the pattern.
     * @param pattern pattern
     */
    public void setPattern(FillPatternKind pattern) { this.pattern = pattern; }

    /**
     * Returns the foreground color.
     * @return the requested result
     */
    public ColorValue getForegroundColor() { return foregroundColor; }
    /**
     * Sets the foreground color.
     * @param foregroundColor foreground color
     */
    public void setForegroundColor(ColorValue foregroundColor) { this.foregroundColor = foregroundColor; }

    /**
     * Returns the background color.
     * @return the requested result
     */
    public ColorValue getBackgroundColor() { return backgroundColor; }
    /**
     * Sets the background color.
     * @param backgroundColor background color
     */
    public void setBackgroundColor(ColorValue backgroundColor) { this.backgroundColor = backgroundColor; }

    /**
     * Returns the borders.
     * @return the requested result
     */
    public BordersValue getBorders() { return borders; }
    /**
     * Sets the borders.
     * @param borders borders
     */
    public void setBorders(BordersValue borders) { this.borders = borders; }

    /**
     * Returns the alignment.
     * @return the requested result
     */
    public AlignmentValue getAlignment() { return alignment; }
    /**
     * Sets the alignment.
     * @param alignment alignment
     */
    public void setAlignment(AlignmentValue alignment) { this.alignment = alignment; }

    /**
     * Returns the protection.
     * @return the requested result
     */
    public ProtectionValue getProtection() { return protection; }
    /**
     * Sets the protection.
     * @param protection protection
     */
    public void setProtection(ProtectionValue protection) { this.protection = protection; }

    /**
     * Returns the number format.
     * @return the requested result
     */
    public NumberFormatValue getNumberFormat() { return numberFormat; }
    /**
     * Sets the number format.
     * @param numberFormat number format
     */
    public void setNumberFormat(NumberFormatValue numberFormat) { this.numberFormat = numberFormat; }

    /**
     * Creates a copy of this instance.
     * @return the computed result
     */
    public StyleValue clone() {
        StyleValue cloned = new StyleValue();
        cloned.font = this.font.clone();
        cloned.pattern = this.pattern;
        cloned.foregroundColor = this.foregroundColor;
        cloned.backgroundColor = this.backgroundColor;
        cloned.borders = this.borders.clone();
        cloned.alignment = this.alignment.clone();
        cloned.protection = this.protection.clone();
        cloned.numberFormat = this.numberFormat.clone();
        return cloned;
    }
}