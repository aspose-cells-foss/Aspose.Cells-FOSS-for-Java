package org.aspose.cells_foss;

import org.aspose.cells_foss.core.WorksheetProtectionModel;

/**
 * Represents protection settings for a worksheet.
 * Boolean allow-properties follow the Aspose.Cells convention: {@code true} means
 * the operation IS allowed (i.e. NOT restricted), {@code false} means it is restricted.
 * This is the inverse of the raw OOXML {@code sheetProtection} attribute values.
 */
public final class WorksheetProtection {
    private final WorksheetProtectionModel model;

    WorksheetProtection(WorksheetProtectionModel model) {
        this.model = model;
    }

    /** Returns {@code true} if the worksheet is protected. */
    public boolean isProtected() {
        return model.getIsProtected();
    }

    /** Returns {@code true} if editing objects (shapes, charts) is allowed. */
    public boolean getAllowEditingObject() { return !model.getObjects(); }
    /** Sets whether editing objects is allowed. Setting to {@code false} also marks the sheet protected. */
    public void setAllowEditingObject(boolean allow) {
        model.setObjects(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if editing scenarios is allowed. */
    public boolean getAllowEditingScenario() { return !model.getScenarios(); }
    /** Sets whether editing scenarios is allowed. */
    public void setAllowEditingScenario(boolean allow) {
        model.setScenarios(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if formatting cells is allowed. */
    public boolean getAllowFormattingCell() { return !model.getFormatCells(); }
    /** Sets whether formatting cells is allowed. */
    public void setAllowFormattingCell(boolean allow) {
        model.setFormatCells(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if formatting columns is allowed. */
    public boolean getAllowFormattingColumn() { return !model.getFormatColumns(); }
    /** Sets whether formatting columns is allowed. */
    public void setAllowFormattingColumn(boolean allow) {
        model.setFormatColumns(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if formatting rows is allowed. */
    public boolean getAllowFormattingRow() { return !model.getFormatRows(); }
    /** Sets whether formatting rows is allowed. */
    public void setAllowFormattingRow(boolean allow) {
        model.setFormatRows(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if inserting columns is allowed. */
    public boolean getAllowInsertingColumn() { return !model.getInsertColumns(); }
    /** Sets whether inserting columns is allowed. */
    public void setAllowInsertingColumn(boolean allow) {
        model.setInsertColumns(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if inserting rows is allowed. */
    public boolean getAllowInsertingRow() { return !model.getInsertRows(); }
    /** Sets whether inserting rows is allowed. */
    public void setAllowInsertingRow(boolean allow) {
        model.setInsertRows(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if inserting hyperlinks is allowed. */
    public boolean getAllowInsertingHyperlink() { return !model.getInsertHyperlinks(); }
    /** Sets whether inserting hyperlinks is allowed. */
    public void setAllowInsertingHyperlink(boolean allow) {
        model.setInsertHyperlinks(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if deleting columns is allowed. */
    public boolean getAllowDeletingColumn() { return !model.getDeleteColumns(); }
    /** Sets whether deleting columns is allowed. */
    public void setAllowDeletingColumn(boolean allow) {
        model.setDeleteColumns(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if deleting rows is allowed. */
    public boolean getAllowDeletingRow() { return !model.getDeleteRows(); }
    /** Sets whether deleting rows is allowed. */
    public void setAllowDeletingRow(boolean allow) {
        model.setDeleteRows(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if selecting locked cells is allowed. */
    public boolean getAllowSelectingLockedCell() { return !model.getSelectLockedCells(); }
    /** Sets whether selecting locked cells is allowed. */
    public void setAllowSelectingLockedCell(boolean allow) {
        model.setSelectLockedCells(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if sorting is allowed. */
    public boolean getAllowSorting() { return !model.getSort(); }
    /** Sets whether sorting is allowed. */
    public void setAllowSorting(boolean allow) {
        model.setSort(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if using AutoFilter is allowed. */
    public boolean getAllowFiltering() { return !model.getAutoFilter(); }
    /** Sets whether using AutoFilter is allowed. */
    public void setAllowFiltering(boolean allow) {
        model.setAutoFilter(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if using pivot tables is allowed. */
    public boolean getAllowUsingPivotTable() { return !model.getPivotTables(); }
    /** Sets whether using pivot tables is allowed. */
    public void setAllowUsingPivotTable(boolean allow) {
        model.setPivotTables(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Returns {@code true} if selecting unlocked cells is allowed. */
    public boolean getAllowSelectingUnlockedCell() { return !model.getSelectUnlockedCells(); }
    /** Sets whether selecting unlocked cells is allowed. */
    public void setAllowSelectingUnlockedCell(boolean allow) {
        model.setSelectUnlockedCells(!allow);
        if (!allow) model.setIsProtected(true);
    }

    /** Resets all protection settings to defaults. */
    void reset() {
        model.clear();
    }
}

