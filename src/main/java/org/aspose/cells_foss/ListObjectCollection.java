package org.aspose.cells_foss;

import org.aspose.cells_foss.core.CellAddress;
import org.aspose.cells_foss.core.ListColumnModel;
import org.aspose.cells_foss.core.ListObjectModel;
import java.util.List;

/** Collection of Excel tables (ListObjects) on a worksheet. */
public final class ListObjectCollection {
    private final List<ListObjectModel> models;
    private Cells cells;
    private int nextId = 1;

    ListObjectCollection(List<ListObjectModel> models) {
        this.models = models;
        for (ListObjectModel m : models) {
            if (m.getId() >= nextId) nextId = m.getId() + 1;
        }
    }

    void setCells(Cells cells) { this.cells = cells; }

    public int getCount() { return models.size(); }

    public ListObject get(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("ListObject index out of range.");
        return new ListObject(models.get(index), models);
    }

    /** Returns the table with the given display name (case-insensitive). Throws CellsException if not found. */
    public ListObject get(String name) {
        if (name == null) throw new CellsException("Table name must not be null.");
        for (ListObjectModel m : models) {
            if (m.getDisplayName().equalsIgnoreCase(name)) return new ListObject(m, models);
        }
        throw new CellsException("Table '" + name + "' was not found.");
    }

    /** Adds a table covering the given zero-based row/column range. Returns the new table's index. */
    public int add(int startRow, int startColumn, int endRow, int endColumn, boolean hasHeaders) {
        if (startRow < 0 || startColumn < 0) throw new CellsException("Start coordinates must be zero or greater.");
        if (endRow < startRow) throw new CellsException("EndRow must be >= StartRow.");
        if (endColumn < startColumn) throw new CellsException("EndColumn must be >= StartColumn.");
        checkNoOverlap(startRow, startColumn, endRow, endColumn, null);
        ListObjectModel model = buildModel(startRow, startColumn, endRow, endColumn, hasHeaders);
        models.add(model);
        return models.size() - 1;
    }

    /** Adds a table from A1-style cell names (e.g. "A1", "D5"). Returns the new table's index. */
    public int add(String startCellName, String endCellName, boolean hasHeaders) {
        if (startCellName == null || startCellName.isBlank()) throw new CellsException("Start cell name must be non-empty.");
        if (endCellName == null || endCellName.isBlank()) throw new CellsException("End cell name must be non-empty.");
        CellAddress a1 = CellAddress.parse(startCellName.trim());
        CellAddress a2 = CellAddress.parse(endCellName.trim());
        return add(a1.getRowIndex(), a1.getColumnIndex(), a2.getRowIndex(), a2.getColumnIndex(), hasHeaders);
    }

    public void removeAt(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("ListObject index out of range.");
        models.remove(index);
    }

    // -------------------------------------------------------------------------
    // Package-private helpers used by ListObject.java
    // -------------------------------------------------------------------------

    public static TableStyleType parseTableStyleType(String name) {
        if (name == null || name.isBlank()) return TableStyleType.NONE;
        for (TableStyleType t : TableStyleType.values()) {
            if (tableStyleTypeName(t).equalsIgnoreCase(name)) return t;
        }
        return TableStyleType.CUSTOM;
    }

    public static String tableStyleTypeName(TableStyleType type) {
        return switch (type) {
            case NONE                    -> "";
            case CUSTOM                  -> "";
            case TABLE_STYLE_LIGHT_1     -> "TableStyleLight1";
            case TABLE_STYLE_LIGHT_2     -> "TableStyleLight2";
            case TABLE_STYLE_LIGHT_3     -> "TableStyleLight3";
            case TABLE_STYLE_LIGHT_4     -> "TableStyleLight4";
            case TABLE_STYLE_LIGHT_5     -> "TableStyleLight5";
            case TABLE_STYLE_LIGHT_6     -> "TableStyleLight6";
            case TABLE_STYLE_LIGHT_7     -> "TableStyleLight7";
            case TABLE_STYLE_LIGHT_8     -> "TableStyleLight8";
            case TABLE_STYLE_LIGHT_9     -> "TableStyleLight9";
            case TABLE_STYLE_LIGHT_10    -> "TableStyleLight10";
            case TABLE_STYLE_LIGHT_11    -> "TableStyleLight11";
            case TABLE_STYLE_LIGHT_12    -> "TableStyleLight12";
            case TABLE_STYLE_LIGHT_13    -> "TableStyleLight13";
            case TABLE_STYLE_LIGHT_14    -> "TableStyleLight14";
            case TABLE_STYLE_LIGHT_15    -> "TableStyleLight15";
            case TABLE_STYLE_LIGHT_16    -> "TableStyleLight16";
            case TABLE_STYLE_LIGHT_17    -> "TableStyleLight17";
            case TABLE_STYLE_LIGHT_18    -> "TableStyleLight18";
            case TABLE_STYLE_LIGHT_19    -> "TableStyleLight19";
            case TABLE_STYLE_LIGHT_20    -> "TableStyleLight20";
            case TABLE_STYLE_LIGHT_21    -> "TableStyleLight21";
            case TABLE_STYLE_MEDIUM_1    -> "TableStyleMedium1";
            case TABLE_STYLE_MEDIUM_2    -> "TableStyleMedium2";
            case TABLE_STYLE_MEDIUM_3    -> "TableStyleMedium3";
            case TABLE_STYLE_MEDIUM_4    -> "TableStyleMedium4";
            case TABLE_STYLE_MEDIUM_5    -> "TableStyleMedium5";
            case TABLE_STYLE_MEDIUM_6    -> "TableStyleMedium6";
            case TABLE_STYLE_MEDIUM_7    -> "TableStyleMedium7";
            case TABLE_STYLE_MEDIUM_8    -> "TableStyleMedium8";
            case TABLE_STYLE_MEDIUM_9    -> "TableStyleMedium9";
            case TABLE_STYLE_MEDIUM_10   -> "TableStyleMedium10";
            case TABLE_STYLE_MEDIUM_11   -> "TableStyleMedium11";
            case TABLE_STYLE_MEDIUM_12   -> "TableStyleMedium12";
            case TABLE_STYLE_MEDIUM_13   -> "TableStyleMedium13";
            case TABLE_STYLE_MEDIUM_14   -> "TableStyleMedium14";
            case TABLE_STYLE_MEDIUM_15   -> "TableStyleMedium15";
            case TABLE_STYLE_MEDIUM_16   -> "TableStyleMedium16";
            case TABLE_STYLE_MEDIUM_17   -> "TableStyleMedium17";
            case TABLE_STYLE_MEDIUM_18   -> "TableStyleMedium18";
            case TABLE_STYLE_MEDIUM_19   -> "TableStyleMedium19";
            case TABLE_STYLE_MEDIUM_20   -> "TableStyleMedium20";
            case TABLE_STYLE_MEDIUM_21   -> "TableStyleMedium21";
            case TABLE_STYLE_MEDIUM_22   -> "TableStyleMedium22";
            case TABLE_STYLE_MEDIUM_23   -> "TableStyleMedium23";
            case TABLE_STYLE_MEDIUM_24   -> "TableStyleMedium24";
            case TABLE_STYLE_MEDIUM_25   -> "TableStyleMedium25";
            case TABLE_STYLE_MEDIUM_26   -> "TableStyleMedium26";
            case TABLE_STYLE_MEDIUM_27   -> "TableStyleMedium27";
            case TABLE_STYLE_MEDIUM_28   -> "TableStyleMedium28";
            case TABLE_STYLE_DARK_1      -> "TableStyleDark1";
            case TABLE_STYLE_DARK_2      -> "TableStyleDark2";
            case TABLE_STYLE_DARK_3      -> "TableStyleDark3";
            case TABLE_STYLE_DARK_4      -> "TableStyleDark4";
            case TABLE_STYLE_DARK_5      -> "TableStyleDark5";
            case TABLE_STYLE_DARK_6      -> "TableStyleDark6";
            case TABLE_STYLE_DARK_7      -> "TableStyleDark7";
            case TABLE_STYLE_DARK_8      -> "TableStyleDark8";
            case TABLE_STYLE_DARK_9      -> "TableStyleDark9";
            case TABLE_STYLE_DARK_10     -> "TableStyleDark10";
            case TABLE_STYLE_DARK_11     -> "TableStyleDark11";
        };
    }

    // -------------------------------------------------------------------------

    private void checkNoOverlap(int sr, int sc, int er, int ec, ListObjectModel skip) {
        for (ListObjectModel m : models) {
            if (m == skip) continue;
            if (sr <= m.getEndRow() && er >= m.getStartRow() &&
                sc <= m.getEndColumn() && ec >= m.getStartColumn()) {
                throw new CellsException("Table range overlaps with existing table '" + m.getDisplayName() + "'.");
            }
        }
    }

    private ListObjectModel buildModel(int sr, int sc, int er, int ec, boolean hasHeaders) {
        ListObjectModel model = new ListObjectModel();
        model.setId(nextId++);
        model.setDisplayName("Table" + model.getId());
        model.setStartRow(sr);
        model.setStartColumn(sc);
        model.setEndRow(er);
        model.setEndColumn(ec);
        model.setShowHeaderRow(hasHeaders);
        int colCount = ec - sc + 1;
        for (int i = 0; i < colCount; i++) {
            ListColumnModel col = new ListColumnModel();
            col.setId(i + 1);
            String name = null;
            if (hasHeaders && cells != null) {
                String cellValue = cells.get(sr, sc + i).getStringValue();
                if (cellValue != null && !cellValue.isBlank()) name = cellValue;
            }
            col.setName(name != null ? name : "Column" + (i + 1));
            model.getColumns().add(col);
        }
        return model;
    }
}

