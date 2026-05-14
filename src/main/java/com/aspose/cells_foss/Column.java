package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ColumnRangeModel;
import java.util.List;

/**
 * Represents a column in a worksheet.
 */
public final class Column {
    private final Worksheet worksheet;
    private final int index;

    /**
     * Initializes a new Column instance.
     * @param worksheet worksheet to apply
     * @param index index
     */
    Column(Worksheet worksheet, int index) {
        this.worksheet = worksheet;
        this.index = index;
    }

    /**
     * Returns the width.
     * @return the requested result
     */
    public Double getWidth() {
        ColumnRangeModel range = findRange();
        return range != null ? range.getWidth() : null;
    }

    /**
     * Sets the width.
     * @param value value to apply
     */
    public void setWidth(Double value) {
        // Handle the relevant branch before the state changes.
        if (value != null && value <= 0.0) {
            throw new CellsException("Column width must be positive.");
        }
        ColumnRangeModel range = getOrCreateRange();
        range.setWidth(value);
    }

    /**
     * Returns the hidden.
     * @return the requested result
     */
    public boolean isHidden() {
        ColumnRangeModel range = findRange();
        return range != null && range.getHidden();
    }

    /**
     * Sets the hidden.
     * @param value value to apply
     */
    public void setHidden(boolean value) {
        ColumnRangeModel range = getOrCreateRange();
        range.setHidden(value);
    }

    /** Returns the outline/group nesting level (0 = ungrouped, 1–7). */
    public int getGroupLevel() {
        ColumnRangeModel range = findRange();
        return range != null ? range.getOutlineLevel() : 0;
    }

    /** Sets the outline/group nesting level (0 = ungrouped, 1–7). */
    public void setGroupLevel(int level) {
        // Validate the caller input before continuing.
        if (level < 0 || level > 7) throw new CellsException("Group level must be between 0 and 7.");
        ColumnRangeModel range = getOrCreateRange();
        range.setOutlineLevel(level);
    }

    /** Returns true if this column's outline group is collapsed. */
    public boolean isCollapsed() {
        ColumnRangeModel range = findRange();
        return range != null && range.getCollapsed();
    }

    /** Sets whether this column's outline group is collapsed. */
    public void setCollapsed(boolean value) {
        ColumnRangeModel range = getOrCreateRange();
        range.setCollapsed(value);
    }

    /**
     * Finds range.
     * @return the requested result
     */
    private ColumnRangeModel findRange() {
        // Walk the current collection so every entry is processed consistently.
        for (ColumnRangeModel range : worksheet.getModel().getColumns()) {
            if (range.getMinColumnIndex() <= index && index <= range.getMaxColumnIndex()) {
                return range;
            }
        }
        return null;
    }

    /**
     * Returns the or create range.
     * @return the requested result
     */
    private ColumnRangeModel getOrCreateRange() {
        ColumnRangeModel existing = findRange();
        if (existing != null && existing.getMinColumnIndex() == index && existing.getMaxColumnIndex() == index) {
            return existing;
        }
        // Split any overlapping range and create a single-column entry
        List<ColumnRangeModel> columns = worksheet.getModel().getColumns();
        if (existing != null) {
            // Split existing range at this column
            int min = existing.getMinColumnIndex();
            int max = existing.getMaxColumnIndex();
            columns.remove(existing);
            if (min < index) {
                ColumnRangeModel before = new ColumnRangeModel();
                before.setMinColumnIndex(min);
                before.setMaxColumnIndex(index - 1);
                before.setWidth(existing.getWidth());
                before.setHidden(existing.getHidden());
                before.setOutlineLevel(existing.getOutlineLevel());
                before.setCollapsed(existing.getCollapsed());
                columns.add(before);
            }
            if (max > index) {
                ColumnRangeModel after = new ColumnRangeModel();
                after.setMinColumnIndex(index + 1);
                after.setMaxColumnIndex(max);
                after.setWidth(existing.getWidth());
                after.setHidden(existing.getHidden());
                after.setOutlineLevel(existing.getOutlineLevel());
                after.setCollapsed(existing.getCollapsed());
                columns.add(after);
            }
        }
        ColumnRangeModel single = new ColumnRangeModel();
        single.setMinColumnIndex(index);
        single.setMaxColumnIndex(index);
        columns.add(single);
        return single;
    }
}
