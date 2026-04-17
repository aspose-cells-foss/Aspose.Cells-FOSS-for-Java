package com.aspose.cells_foss;

import com.aspose.cells_foss.core.WorksheetProtectionModel;

/**
 * Represents protection settings for a worksheet.
 */
public final class WorksheetProtection {
    private final WorksheetProtectionModel model;

    /**
     * Initializes a new WorksheetProtection instance.
     * @param model model
     */
    WorksheetProtection(WorksheetProtectionModel model) {
        this.model = model;
    }

    /**
     * Returns the protected.
     * @return the requested result
     */
    public boolean getIsProtected() {
        return model.getIsProtected();
    }

    /**
     * Sets the protected.
     * @param isProtected is protected
     */
    public void setIsProtected(boolean isProtected) {
        model.setIsProtected(isProtected);
    }

    /**
     * Returns the objects.
     * @return the requested result
     */
    public boolean getObjects() {
        return model.getObjects();
    }

    /**
     * Sets the objects.
     * @param objects objects
     */
    public void setObjects(boolean objects) {
        model.setObjects(objects);
        markProtectedWhenEnabled(objects);
    }

    /**
     * Returns the scenarios.
     * @return the requested result
     */
    public boolean getScenarios() {
        return model.getScenarios();
    }

    /**
     * Sets the scenarios.
     * @param scenarios scenarios
     */
    public void setScenarios(boolean scenarios) {
        model.setScenarios(scenarios);
        markProtectedWhenEnabled(scenarios);
    }

    /**
     * Returns the format cells.
     * @return the requested result
     */
    public boolean getFormatCells() {
        return model.getFormatCells();
    }

    /**
     * Sets the format cells.
     * @param formatCells format cells
     */
    public void setFormatCells(boolean formatCells) {
        model.setFormatCells(formatCells);
        markProtectedWhenEnabled(formatCells);
    }

    /**
     * Returns the format columns.
     * @return the requested result
     */
    public boolean getFormatColumns() {
        return model.getFormatColumns();
    }

    /**
     * Sets the format columns.
     * @param formatColumns format columns
     */
    public void setFormatColumns(boolean formatColumns) {
        model.setFormatColumns(formatColumns);
        markProtectedWhenEnabled(formatColumns);
    }

    /**
     * Returns the format rows.
     * @return the requested result
     */
    public boolean getFormatRows() {
        return model.getFormatRows();
    }

    /**
     * Sets the format rows.
     * @param formatRows format rows
     */
    public void setFormatRows(boolean formatRows) {
        model.setFormatRows(formatRows);
        markProtectedWhenEnabled(formatRows);
    }

    /**
     * Returns the insert columns.
     * @return the requested result
     */
    public boolean getInsertColumns() {
        return model.getInsertColumns();
    }

    /**
     * Sets the insert columns.
     * @param insertColumns insert columns
     */
    public void setInsertColumns(boolean insertColumns) {
        model.setInsertColumns(insertColumns);
        markProtectedWhenEnabled(insertColumns);
    }

    /**
     * Returns the insert rows.
     * @return the requested result
     */
    public boolean getInsertRows() {
        return model.getInsertRows();
    }

    /**
     * Sets the insert rows.
     * @param insertRows insert rows
     */
    public void setInsertRows(boolean insertRows) {
        model.setInsertRows(insertRows);
        markProtectedWhenEnabled(insertRows);
    }

    /**
     * Returns the insert hyperlinks.
     * @return the requested result
     */
    public boolean getInsertHyperlinks() {
        return model.getInsertHyperlinks();
    }

    /**
     * Sets the insert hyperlinks.
     * @param insertHyperlinks insert hyperlinks
     */
    public void setInsertHyperlinks(boolean insertHyperlinks) {
        model.setInsertHyperlinks(insertHyperlinks);
        markProtectedWhenEnabled(insertHyperlinks);
    }

    /**
     * Returns the delete columns.
     * @return the requested result
     */
    public boolean getDeleteColumns() {
        return model.getDeleteColumns();
    }

    /**
     * Sets the delete columns.
     * @param deleteColumns delete columns
     */
    public void setDeleteColumns(boolean deleteColumns) {
        model.setDeleteColumns(deleteColumns);
        markProtectedWhenEnabled(deleteColumns);
    }

    /**
     * Returns the delete rows.
     * @return the requested result
     */
    public boolean getDeleteRows() {
        return model.getDeleteRows();
    }

    /**
     * Sets the delete rows.
     * @param deleteRows delete rows
     */
    public void setDeleteRows(boolean deleteRows) {
        model.setDeleteRows(deleteRows);
        markProtectedWhenEnabled(deleteRows);
    }

    /**
     * Returns the select locked cells.
     * @return the requested result
     */
    public boolean getSelectLockedCells() {
        return model.getSelectLockedCells();
    }

    /**
     * Sets the select locked cells.
     * @param selectLockedCells select locked cells
     */
    public void setSelectLockedCells(boolean selectLockedCells) {
        model.setSelectLockedCells(selectLockedCells);
        markProtectedWhenEnabled(selectLockedCells);
    }

    /**
     * Returns the sort.
     * @return the requested result
     */
    public boolean getSort() {
        return model.getSort();
    }

    /**
     * Sets the sort.
     * @param sort sort
     */
    public void setSort(boolean sort) {
        model.setSort(sort);
        markProtectedWhenEnabled(sort);
    }

    /**
     * Returns the auto filter.
     * @return the requested result
     */
    public boolean getAutoFilter() {
        return model.getAutoFilter();
    }

    /**
     * Sets the auto filter.
     * @param autoFilter auto filter
     */
    public void setAutoFilter(boolean autoFilter) {
        model.setAutoFilter(autoFilter);
        markProtectedWhenEnabled(autoFilter);
    }

    /**
     * Returns the pivot tables.
     * @return the requested result
     */
    public boolean getPivotTables() {
        return model.getPivotTables();
    }

    /**
     * Sets the pivot tables.
     * @param pivotTables pivot tables
     */
    public void setPivotTables(boolean pivotTables) {
        model.setPivotTables(pivotTables);
        markProtectedWhenEnabled(pivotTables);
    }

    /**
     * Returns the select unlocked cells.
     * @return the requested result
     */
    public boolean getSelectUnlockedCells() {
        return model.getSelectUnlockedCells();
    }

    /**
     * Sets the select unlocked cells.
     * @param selectUnlockedCells select unlocked cells
     */
    public void setSelectUnlockedCells(boolean selectUnlockedCells) {
        model.setSelectUnlockedCells(selectUnlockedCells);
        markProtectedWhenEnabled(selectUnlockedCells);
    }

    /**
     * Resets the current state to its defaults.
     */
    void reset() {
        model.clear();
    }

    /**
     * Processes mark protected when enabled.
     * @param value value to apply
     */
    private void markProtectedWhenEnabled(boolean value) {
        // Handle the relevant branch before the state changes.
        if (value) {
            model.setIsProtected(true);
        }
    }
}