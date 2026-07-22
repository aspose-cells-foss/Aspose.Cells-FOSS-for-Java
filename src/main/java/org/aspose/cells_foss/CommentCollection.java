package org.aspose.cells_foss;

import org.aspose.cells_foss.core.CellAddress;
import org.aspose.cells_foss.core.CommentModel;
import java.util.List;

/** Collection of cell comments on a worksheet. */
public final class CommentCollection {
    private final List<CommentModel> models;

    CommentCollection(List<CommentModel> models) {
        this.models = models;
    }

    public int getCount() { return models.size(); }

    public Comment get(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("Comment index out of range.");
        return new Comment(models.get(index));
    }

    /** Returns the comment at the given A1-style cell name, or null if none. */
    public Comment get(String cellName) {
        if (cellName == null || cellName.isBlank()) return null;
        CellAddress addr = CellAddress.parse(cellName.trim());
        for (CommentModel m : models) {
            if (m.getRow() == addr.getRowIndex() && m.getColumn() == addr.getColumnIndex())
                return new Comment(m);
        }
        return null;
    }

    /** Adds a comment at the given zero-based row/column. */
    public Comment add(int row, int column) {
        if (row < 0) throw new CellsException("Row must be zero or greater.");
        if (column < 0) throw new CellsException("Column must be zero or greater.");
        for (CommentModel m : models) {
            if (m.getRow() == row && m.getColumn() == column)
                throw new CellsException("A comment already exists at row " + row + ", column " + column + ".");
        }
        CommentModel model = new CommentModel();
        model.setRow(row);
        model.setColumn(column);
        models.add(model);
        return new Comment(model);
    }

    /** Adds a comment at the given A1-style cell reference. */
    public Comment add(String cellName) {
        if (cellName == null || cellName.isBlank())
            throw new CellsException("Cell name must be non-empty.");
        CellAddress addr = CellAddress.parse(cellName.trim());
        return add(addr.getRowIndex(), addr.getColumnIndex());
    }

    public void removeAt(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("Comment index out of range.");
        models.remove(index);
    }

    /** Removes the comment at the given cell reference (no-op if none). */
    public void removeAt(String cellName) {
        if (cellName == null || cellName.isBlank()) return;
        CellAddress addr = CellAddress.parse(cellName.trim());
        models.removeIf(m -> m.getRow() == addr.getRowIndex() && m.getColumn() == addr.getColumnIndex());
    }
}

