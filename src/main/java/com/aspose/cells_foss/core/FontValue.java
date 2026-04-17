package com.aspose.cells_foss.core;

/**
 * Represents a font value with its properties.
 */
public final class FontValue {
    private String name = "Calibri";
    private double size = 11.0;
    private boolean bold;
    private boolean italic;
    private boolean underline;
    private boolean strikeThrough;
    private ColorValue color;

    /**
     * Returns the name.
     * @return the requested result
     */
    public String getName() { return name; }
    /**
     * Sets the name.
     * @param name name
     */
    public void setName(String name) { this.name = name; }

    /**
     * Returns the size.
     * @return the requested result
     */
    public double getSize() { return size; }
    /**
     * Sets the size.
     * @param size size
     */
    public void setSize(double size) { this.size = size; }

    /**
     * Returns the bold.
     * @return the requested result
     */
    public boolean getBold() { return bold; }
    /**
     * Sets the bold.
     * @param bold bold
     */
    public void setBold(boolean bold) { this.bold = bold; }

    /**
     * Returns the italic.
     * @return the requested result
     */
    public boolean getItalic() { return italic; }
    /**
     * Sets the italic.
     * @param italic italic
     */
    public void setItalic(boolean italic) { this.italic = italic; }

    /**
     * Returns the underline.
     * @return the requested result
     */
    public boolean getUnderline() { return underline; }
    /**
     * Sets the underline.
     * @param underline underline
     */
    public void setUnderline(boolean underline) { this.underline = underline; }

    /**
     * Returns the strike through.
     * @return the requested result
     */
    public boolean getStrikeThrough() { return strikeThrough; }
    /**
     * Sets the strike through.
     * @param strikeThrough strike through
     */
    public void setStrikeThrough(boolean strikeThrough) { this.strikeThrough = strikeThrough; }

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
    public FontValue clone() {
        FontValue cloned = new FontValue();
        cloned.name = this.name;
        cloned.size = this.size;
        cloned.bold = this.bold;
        cloned.italic = this.italic;
        cloned.underline = this.underline;
        cloned.strikeThrough = this.strikeThrough;
        cloned.color = this.color;
        return cloned;
    }
}