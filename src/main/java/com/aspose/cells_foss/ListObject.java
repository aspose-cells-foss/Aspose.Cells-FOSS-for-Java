package com.aspose.cells_foss;

import com.aspose.cells_foss.core.CellAddress;
import com.aspose.cells_foss.core.ListColumnModel;
import com.aspose.cells_foss.core.ListObjectModel;
import java.util.List;

/** Represents an Excel table (structured reference / ListObject). */
public final class ListObject {
    private final ListObjectModel model;
    private final List<ListObjectModel> ownerList;
    private final ListColumnCollection listColumns;

    ListObject(ListObjectModel model, List<ListObjectModel> ownerList) {
        this.model = model;
        this.ownerList = ownerList;
        this.listColumns = new ListColumnCollection(model.getColumns());
    }

    ListObjectModel getModel() { return model; }

    public String getDisplayName() { return model.getDisplayName(); }
    public void setDisplayName(String name) {
        if (name == null || name.isBlank() || name.contains(" "))
            throw new CellsException("DisplayName must be non-empty and contain no spaces.");
        model.setDisplayName(name.trim());
    }

    public String getComment() { return model.getComment(); }
    public void setComment(String comment) { model.setComment(comment != null ? comment : ""); }

    public int getStartRow() { return model.getStartRow(); }
    public int getStartColumn() { return model.getStartColumn(); }
    public int getEndRow() { return model.getEndRow(); }
    public int getEndColumn() { return model.getEndColumn(); }

    public boolean isShowHeaderRow() { return model.isShowHeaderRow(); }
    public void setShowHeaderRow(boolean show) {
        model.setShowHeaderRow(show);
        model.setAutoFilterEnabled(show && model.isAutoFilterEnabled());
    }

    public boolean isShowTotals() { return model.isShowTotals(); }
    public void setShowTotals(boolean show) { model.setShowTotals(show); }

    public TableStyleType getTableStyleType() {
        return ListObjectCollection.parseTableStyleType(model.getTableStyleName());
    }
    public void setTableStyleType(TableStyleType type) {
        if (type == TableStyleType.CUSTOM) return;
        model.setTableStyleName(ListObjectCollection.tableStyleTypeName(type));
    }

    public String getTableStyleName() { return model.getTableStyleName(); }
    public void setTableStyleName(String name) { model.setTableStyleName(name != null ? name : ""); }

    public boolean isShowTableStyleFirstColumn() { return model.isShowTableStyleFirstColumn(); }
    public void setShowTableStyleFirstColumn(boolean v) { model.setShowTableStyleFirstColumn(v); }

    public boolean isShowTableStyleLastColumn() { return model.isShowTableStyleLastColumn(); }
    public void setShowTableStyleLastColumn(boolean v) { model.setShowTableStyleLastColumn(v); }

    public boolean isShowTableStyleRowStripes() { return model.isShowTableStyleRowStripes(); }
    public void setShowTableStyleRowStripes(boolean v) { model.setShowTableStyleRowStripes(v); }

    public boolean isShowTableStyleColumnStripes() { return model.isShowTableStyleColumnStripes(); }
    public void setShowTableStyleColumnStripes(boolean v) { model.setShowTableStyleColumnStripes(v); }

    public ListColumnCollection getListColumns() { return listColumns; }

    /** Resizes the table to a new range, optionally preserving/removing header row. */
    public void resize(int startRow, int startColumn, int endRow, int endColumn, boolean hasHeaders) {
        if (startRow < 0 || startColumn < 0) throw new CellsException("Start row and column must be zero or greater.");
        if (endRow < startRow) throw new CellsException("EndRow must be >= StartRow.");
        if (endColumn < startColumn) throw new CellsException("EndColumn must be >= StartColumn.");
        model.setStartRow(startRow);
        model.setStartColumn(startColumn);
        model.setEndRow(endRow);
        model.setEndColumn(endColumn);
        model.setShowHeaderRow(hasHeaders);
        rebuildColumns();
    }

    public void showAutoFilter() { model.setAutoFilterEnabled(true); }
    public void removeAutoFilter() { model.setAutoFilterEnabled(false); }

    /** Removes the table structure, leaving cell data in place, and removes it from the collection. */
    public void convertToRange() {
        ownerList.remove(model);
    }

    private void rebuildColumns() {
        List<ListColumnModel> cols = model.getColumns();
        int colCount = model.getEndColumn() - model.getStartColumn() + 1;
        while (cols.size() < colCount) {
            ListColumnModel col = new ListColumnModel();
            col.setId(cols.size() + 1);
            col.setName("Column" + (cols.size() + 1));
            cols.add(col);
        }
        while (cols.size() > colCount) {
            cols.remove(cols.size() - 1);
        }
    }
}
