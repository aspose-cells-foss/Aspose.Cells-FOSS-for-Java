package com.aspose.cells_foss.core;

/**
 * Represents a border side value with style and color.
 */
public final class BorderSideValue {
    private BorderStyle style;
    private ColorValue color;

    /**
     * Initializes a new BorderSideValue instance.
     */
    public BorderSideValue() {}

    /**
     * Returns the style.
     * @return the requested result
     */
    public BorderStyle getStyle() { return style; }
    /**
     * Sets the style.
     * @param style style to apply
     */
    public void setStyle(BorderStyle style) { this.style = style; }

    /**
     * Returns the color.
     * @return the requested result
     */
    public ColorValue getColor() { return color; }
    /**
     * Sets the color.
     * @param color color
     */
    public void setColor(ColorValue color) { this.color = color; }

    /**
     * Creates a copy of this instance.
     * @return the computed result
     */
    public BorderSideValue clone() {
        BorderSideValue cloned = new BorderSideValue();
        cloned.style = this.style;
        cloned.color = this.color;
        return cloned;
    }
}