package org.aspose.cells_foss.core;

/**
 * Internal model for a cell comment (note).
 */
public final class CommentModel {
    private int row;
    private int column;
    private String author = "";
    private String note = "";
    private boolean isVisible = false;
    private int width = 129;
    private int height = 75;
    private String rawVmlShapeXml;

    public int getRow() { return row; }
    public void setRow(int row) { this.row = row; }

    public int getColumn() { return column; }
    public void setColumn(int column) { this.column = column; }

    public String getAuthor() { return author; }
    public void setAuthor(String author) { this.author = author != null ? author : ""; }

    public String getNote() { return note; }
    public void setNote(String note) { this.note = note != null ? note : ""; }

    public boolean isVisible() { return isVisible; }
    public void setVisible(boolean visible) { this.isVisible = visible; }

    public int getWidth() { return width; }
    public void setWidth(int width) { this.width = Math.max(1, width); }

    public int getHeight() { return height; }
    public void setHeight(int height) { this.height = Math.max(1, height); }

    /** Preserved VML shape XML for round-trip fidelity (may be null for programmatic comments). */
    public String getRawVmlShapeXml() { return rawVmlShapeXml; }
    public void setRawVmlShapeXml(String rawVmlShapeXml) { this.rawVmlShapeXml = rawVmlShapeXml; }
}

