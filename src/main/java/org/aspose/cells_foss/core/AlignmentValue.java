package org.aspose.cells_foss.core;

/**
 * Represents alignment settings for a cell style.
 */
public final class AlignmentValue {
    private HorizontalAlignment horizontal;
    private VerticalAlignment vertical = VerticalAlignment.BOTTOM;
    private boolean wrapText;
    private int indentLevel;
    private int textRotation;
    private boolean shrinkToFit;
    private int readingOrder;
    private int relativeIndent;

    /**
     * Returns the horizontal.
     * @return the requested result
     */
    public HorizontalAlignment getHorizontal() { return horizontal; }
    /**
     * Sets the horizontal.
     * @param horizontal horizontal
     */
    public void setHorizontal(HorizontalAlignment horizontal) { this.horizontal = horizontal; }

    /**
     * Returns the vertical.
     * @return the requested result
     */
    public VerticalAlignment getVertical() { return vertical; }
    /**
     * Sets the vertical.
     * @param vertical vertical
     */
    public void setVertical(VerticalAlignment vertical) { this.vertical = vertical; }

    /**
     * Returns the wrap text.
     * @return the requested result
     */
    public boolean getWrapText() { return wrapText; }
    /**
     * Sets the wrap text.
     * @param wrapText wrap text
     */
    public void setWrapText(boolean wrapText) { this.wrapText = wrapText; }

    /**
     * Returns the indent level.
     * @return the requested result
     */
    public int getIndentLevel() { return indentLevel; }
    /**
     * Sets the indent level.
     * @param indentLevel indent level
     */
    public void setIndentLevel(int indentLevel) { this.indentLevel = indentLevel; }

    /**
     * Returns the text rotation.
     * @return the requested result
     */
    public int getTextRotation() { return textRotation; }
    /**
     * Sets the text rotation.
     * @param textRotation text rotation
     */
    public void setTextRotation(int textRotation) { this.textRotation = textRotation; }

    /**
     * Returns the shrink to fit.
     * @return the requested result
     */
    public boolean getShrinkToFit() { return shrinkToFit; }
    /**
     * Sets the shrink to fit.
     * @param shrinkToFit shrink to fit
     */
    public void setShrinkToFit(boolean shrinkToFit) { this.shrinkToFit = shrinkToFit; }

    /**
     * Returns the reading order.
     * @return the requested result
     */
    public int getReadingOrder() { return readingOrder; }
    /**
     * Sets the reading order.
     * @param readingOrder reading order
     */
    public void setReadingOrder(int readingOrder) { this.readingOrder = readingOrder; }

    /**
     * Returns the relative indent.
     * @return the requested result
     */
    public int getRelativeIndent() { return relativeIndent; }
    /**
     * Sets the relative indent.
     * @param relativeIndent relative indent
     */
    public void setRelativeIndent(int relativeIndent) { this.relativeIndent = relativeIndent; }

    /**
     * Creates a copy of this instance.
     * @return the computed result
     */
    public AlignmentValue clone() {
        AlignmentValue cloned = new AlignmentValue();
        cloned.horizontal = this.horizontal;
        cloned.vertical = this.vertical;
        cloned.wrapText = this.wrapText;
        cloned.indentLevel = this.indentLevel;
        cloned.textRotation = this.textRotation;
        cloned.shrinkToFit = this.shrinkToFit;
        cloned.readingOrder = this.readingOrder;
        cloned.relativeIndent = this.relativeIndent;
        return cloned;
    }
}
