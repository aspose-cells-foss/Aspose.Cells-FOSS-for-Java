package org.aspose.cells_foss;



/**
 * Represents a border with line style and color.
 */
public class Border {
    private BorderStyleType lineStyle;
    private Color color = Color.getEmpty();

    /**
     * Returns the line style.
     * @return the requested result
     */
    public BorderStyleType getLineStyle() { return lineStyle; }
    /**
     * Sets the line style.
     * @param lineStyle line style
     */
    public void setLineStyle(BorderStyleType lineStyle) { this.lineStyle = lineStyle; }

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
    public Border clone() {
        Border cloned = new Border();
        cloned.lineStyle = this.lineStyle;
        cloned.color = this.color;
        return cloned;
    }

    // Package-internal model conversion
    /**
     * Creates an API object from the backing model.
     * @param bsv bsv
     * @return the computed result
     */
    static Border fromModel(org.aspose.cells_foss.core.BorderSideValue bsv) {
        Border b = new Border();
        b.lineStyle = bsv.getStyle() != null ? mapStyle(bsv.getStyle()) : null;
        b.color = bsv.getColor() != null ? Color.fromCore(bsv.getColor()) : Color.getEmpty();
        return b;
    }

    /**
     * Converts this instance to the backing model representation.
     * @return the computed result
     */
    org.aspose.cells_foss.core.BorderSideValue toModel() {
        org.aspose.cells_foss.core.BorderSideValue bsv = new org.aspose.cells_foss.core.BorderSideValue();
        bsv.setStyle(lineStyle != null ? mapStyleToCore(lineStyle) : null);
        bsv.setColor(color.equals(Color.getEmpty()) ? null : color.toCore());
        return bsv;
    }

    /**
     * Maps style.
     * @param s s
     * @return the computed result
     */
    private static BorderStyleType mapStyle(org.aspose.cells_foss.core.BorderStyle s) {
        // Translate the internal value into the matching public representation.
        switch (s) {
            case THIN:               return BorderStyleType.THIN;
            case MEDIUM:             return BorderStyleType.MEDIUM;
            case THICK:              return BorderStyleType.THICK;
            case DOTTED:             return BorderStyleType.DOTTED;
            case DASHED:             return BorderStyleType.DASHED;
            case DOUBLE:             return BorderStyleType.DOUBLE;
            case HAIR:               return BorderStyleType.HAIR;
            case MEDIUM_DASHED:      return BorderStyleType.MEDIUM_DASHED;
            case DASH_DOT:           return BorderStyleType.DASH_DOT;
            case MEDIUM_DASH_DOT:    return BorderStyleType.MEDIUM_DASH_DOT;
            case DASH_DOT_DOT:       return BorderStyleType.DASH_DOT_DOT;
            case MEDIUM_DASH_DOT_DOT:return BorderStyleType.MEDIUM_DASH_DOT_DOT;
            case SLANTED_DASH_DOT:   return BorderStyleType.SLANTED_DASH_DOT;
            default:                 return BorderStyleType.NONE;
        }
    }

    /**
     * Maps style to core.
     * @param s s
     * @return the computed result
     */
    private static org.aspose.cells_foss.core.BorderStyle mapStyleToCore(BorderStyleType s) {
        // Handle the relevant branch before the state changes.
        if (s == null) return null;
        switch (s) {
            case THIN:               return org.aspose.cells_foss.core.BorderStyle.THIN;
            case MEDIUM:             return org.aspose.cells_foss.core.BorderStyle.MEDIUM;
            case THICK:              return org.aspose.cells_foss.core.BorderStyle.THICK;
            case DOTTED:             return org.aspose.cells_foss.core.BorderStyle.DOTTED;
            case DASHED:             return org.aspose.cells_foss.core.BorderStyle.DASHED;
            case DOUBLE:             return org.aspose.cells_foss.core.BorderStyle.DOUBLE;
            case HAIR:               return org.aspose.cells_foss.core.BorderStyle.HAIR;
            case MEDIUM_DASHED:      return org.aspose.cells_foss.core.BorderStyle.MEDIUM_DASHED;
            case DASH_DOT:           return org.aspose.cells_foss.core.BorderStyle.DASH_DOT;
            case MEDIUM_DASH_DOT:    return org.aspose.cells_foss.core.BorderStyle.MEDIUM_DASH_DOT;
            case DASH_DOT_DOT:       return org.aspose.cells_foss.core.BorderStyle.DASH_DOT_DOT;
            case MEDIUM_DASH_DOT_DOT:return org.aspose.cells_foss.core.BorderStyle.MEDIUM_DASH_DOT_DOT;
            case SLANTED_DASH_DOT:   return org.aspose.cells_foss.core.BorderStyle.SLANTED_DASH_DOT;
            default:                 return org.aspose.cells_foss.core.BorderStyle.NONE;
        }
    }
}
