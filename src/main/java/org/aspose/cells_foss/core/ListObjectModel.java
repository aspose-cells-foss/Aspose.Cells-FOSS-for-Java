package org.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.List;

/**
 * Internal model for a table (ListObject / structured reference).
 */
public final class ListObjectModel {
    private int id;
    private String displayName = "";
    private String comment = "";
    private int startRow;
    private int startColumn;
    private int endRow;
    private int endColumn;
    private boolean showHeaderRow = true;
    private boolean showTotals = false;
    private String tableStyleName = "TableStyleMedium2";
    private boolean showTableStyleFirstColumn = false;
    private boolean showTableStyleLastColumn = false;
    private boolean showTableStyleRowStripes = true;
    private boolean showTableStyleColumnStripes = false;
    private boolean autoFilterEnabled = true;
    private final List<ListColumnModel> columns = new ArrayList<>();

    public int getId() { return id; }
    public void setId(int id) { this.id = id; }

    public String getDisplayName() { return displayName; }
    public void setDisplayName(String displayName) { this.displayName = displayName != null ? displayName : ""; }

    public String getComment() { return comment; }
    public void setComment(String comment) { this.comment = comment != null ? comment : ""; }

    public int getStartRow() { return startRow; }
    public void setStartRow(int startRow) { this.startRow = startRow; }

    public int getStartColumn() { return startColumn; }
    public void setStartColumn(int startColumn) { this.startColumn = startColumn; }

    public int getEndRow() { return endRow; }
    public void setEndRow(int endRow) { this.endRow = endRow; }

    public int getEndColumn() { return endColumn; }
    public void setEndColumn(int endColumn) { this.endColumn = endColumn; }

    public boolean isShowHeaderRow() { return showHeaderRow; }
    public void setShowHeaderRow(boolean showHeaderRow) { this.showHeaderRow = showHeaderRow; }

    public boolean isShowTotals() { return showTotals; }
    public void setShowTotals(boolean showTotals) { this.showTotals = showTotals; }

    public String getTableStyleName() { return tableStyleName; }
    public void setTableStyleName(String tableStyleName) { this.tableStyleName = tableStyleName != null ? tableStyleName : ""; }

    public boolean isShowTableStyleFirstColumn() { return showTableStyleFirstColumn; }
    public void setShowTableStyleFirstColumn(boolean v) { this.showTableStyleFirstColumn = v; }

    public boolean isShowTableStyleLastColumn() { return showTableStyleLastColumn; }
    public void setShowTableStyleLastColumn(boolean v) { this.showTableStyleLastColumn = v; }

    public boolean isShowTableStyleRowStripes() { return showTableStyleRowStripes; }
    public void setShowTableStyleRowStripes(boolean v) { this.showTableStyleRowStripes = v; }

    public boolean isShowTableStyleColumnStripes() { return showTableStyleColumnStripes; }
    public void setShowTableStyleColumnStripes(boolean v) { this.showTableStyleColumnStripes = v; }

    public boolean isAutoFilterEnabled() { return autoFilterEnabled; }
    public void setAutoFilterEnabled(boolean autoFilterEnabled) { this.autoFilterEnabled = autoFilterEnabled; }

    public List<ListColumnModel> getColumns() { return columns; }
}

