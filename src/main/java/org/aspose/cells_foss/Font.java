package org.aspose.cells_foss;

import org.aspose.cells_foss.core.FontValue;

/**
 * Represents a font with its properties.
 */
public class Font {
    private String name = "Calibri";
    private double size = 11.0;
    private boolean bold;
    private boolean italic;
    private boolean underline;
    private boolean strikeThrough;
    private Color color = Color.getEmpty();

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
    public Color getColor() { return color; }
    /**
     * Sets the color.
     * @param color color
     */
    public void setColor(Color color) { this.color = color; }

    /**
     * Creates a copy of this instance.
     * @return the computed result
     */
    public Font clone() {
        Font cloned = new Font();
        cloned.name = this.name;
        cloned.size = this.size;
        cloned.bold = this.bold;
        cloned.italic = this.italic;
        cloned.underline = this.underline;
        cloned.strikeThrough = this.strikeThrough;
        cloned.color = this.color;
        return cloned;
    }

    // Package-internal model conversion
    /**
     * Creates an API object from the backing model.
     * @param fv fv
     * @return the computed result
     */
    static Font fromModel(FontValue fv) {
        Font f = new Font();
        f.name = fv.getName() != null ? fv.getName() : "Calibri";
        f.size = fv.getSize();
        f.bold = fv.getBold();
        f.italic = fv.getItalic();
        f.underline = fv.getUnderline();
        f.strikeThrough = fv.getStrikeThrough();
        f.color = fv.getColor() != null ? Color.fromCore(fv.getColor()) : Color.getEmpty();
        return f;
    }

    /**
     * Converts this instance to the backing model representation.
     * @return the computed result
     */
    FontValue toModel() {
        FontValue fv = new FontValue();
        fv.setName(name);
        fv.setSize(size);
        fv.setBold(bold);
        fv.setItalic(italic);
        fv.setUnderline(underline);
        fv.setStrikeThrough(strikeThrough);
        fv.setColor(color.equals(Color.getEmpty()) ? null : color.toCore());
        return fv;
    }
}
