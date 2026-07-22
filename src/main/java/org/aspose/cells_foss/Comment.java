package org.aspose.cells_foss;

import org.aspose.cells_foss.core.CommentModel;

/** Represents a cell comment (note). */
public final class Comment {
    private final CommentModel model;

    Comment(CommentModel model) {
        this.model = model;
    }

    CommentModel getModel() { return model; }

    /** Zero-based row index of the cell this comment is anchored to. */
    public int getRow() { return model.getRow(); }

    /** Zero-based column index of the cell this comment is anchored to. */
    public int getColumn() { return model.getColumn(); }

    public String getAuthor() { return model.getAuthor(); }
    public void setAuthor(String author) { model.setAuthor(author != null ? author : ""); }

    public String getNote() { return model.getNote(); }
    public void setNote(String note) { model.setNote(note != null ? note : ""); }

    public boolean isVisible() { return model.isVisible(); }
    public void setVisible(boolean visible) { model.setVisible(visible); }

    /** Comment box width in pixels (minimum 1). */
    public int getWidth() { return model.getWidth(); }
    public void setWidth(int width) {
        if (width < 1) throw new CellsException("Width must be at least 1.");
        model.setWidth(width);
    }

    /** Comment box height in pixels (minimum 1). */
    public int getHeight() { return model.getHeight(); }
    public void setHeight(int height) {
        if (height < 1) throw new CellsException("Height must be at least 1.");
        model.setHeight(height);
    }
}

