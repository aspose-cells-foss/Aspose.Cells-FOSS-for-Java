package com.aspose.cells_foss;

import com.aspose.cells_foss.core.RowModel;

/**
 * Represents a row in a worksheet.
 */
public final class Row {
    private final Worksheet worksheet;
    private final int index;

    /**
     * Initializes a new Row instance.
     * @param worksheet worksheet to apply
     * @param index index
     */
    Row(Worksheet worksheet, int index) {
        this.worksheet = worksheet;
        this.index = index;
    }

    /**
     * Returns the height.
     * @return the requested result
     */
    public Double getHeight() {
        RowModel model = worksheet.getModel().getRows().get(index);
        return model != null ? model.getHeight() : null;
    }

    /**
     * Sets the height.
     * @param value value to apply
     */
    public void setHeight(Double value) {
        // Handle the relevant branch before the state changes.
        if (value != null && value <= 0.0) {
            throw new CellsException("Row height must be positive.");
        }
        RowModel model = worksheet.getModel().getRows().computeIfAbsent(index, k -> new RowModel());
        model.setHeight(value);
    }

    /**
     * Returns the hidden.
     * @return the requested result
     */
    public boolean getIsHidden() {
        RowModel model = worksheet.getModel().getRows().get(index);
        return model != null && model.getHidden();
    }

    /**
     * Sets the hidden.
     * @param value value to apply
     */
    public void setIsHidden(boolean value) {
        // Lazily create the backing record the first time this path is used.
        RowModel model = worksheet.getModel().getRows().computeIfAbsent(index, k -> new RowModel());
        model.setHidden(value);
    }

    /** Returns the outline/group nesting level (0 = ungrouped, 1–7). */
    public int getGroupLevel() {
        RowModel model = worksheet.getModel().getRows().get(index);
        return model != null ? model.getOutlineLevel() : 0;
    }

    /** Sets the outline/group nesting level (0 = ungrouped, 1–7). */
    public void setGroupLevel(int level) {
        // Validate the caller input before continuing.
        if (level < 0 || level > 7) throw new CellsException("Group level must be between 0 and 7.");
        RowModel model = worksheet.getModel().getRows().computeIfAbsent(index, k -> new RowModel());
        model.setOutlineLevel(level);
    }

    /** Returns true if this row's outline group is collapsed. */
    public boolean isCollapsed() {
        RowModel model = worksheet.getModel().getRows().get(index);
        return model != null && model.getCollapsed();
    }

    /** Sets whether this row's outline group is collapsed. */
    public void setIsCollapsed(boolean value) {
        // Lazily create the backing record the first time this path is used.
        RowModel model = worksheet.getModel().getRows().computeIfAbsent(index, k -> new RowModel());
        model.setCollapsed(value);
    }
}
