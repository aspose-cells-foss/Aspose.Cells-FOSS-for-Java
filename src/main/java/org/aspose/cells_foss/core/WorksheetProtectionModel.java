package org.aspose.cells_foss.core;

/**
 * Represents the protection model for a worksheet in an Excel file.
 */
public final class WorksheetProtectionModel {
    private boolean isProtected;
    private boolean objects;
    private boolean scenarios;
    private boolean formatCells;
    private boolean formatColumns;
    private boolean formatRows;
    private boolean insertColumns;
    private boolean insertRows;
    private boolean insertHyperlinks;
    private boolean deleteColumns;
    private boolean deleteRows;
    private boolean selectLockedCells;
    private boolean sort;
    private boolean autoFilter;
    private boolean pivotTables;
    private boolean selectUnlockedCells;
    private String passwordHash;
    private String algorithmName;
    private String hashValue;
    private String saltValue;
    private String spinCount;

    /**
     * Returns the protected.
     * @return the requested result
     */
    public boolean getIsProtected() { return isProtected; }
    /**
     * Sets the protected.
     * @param isProtected is protected
     */
    public void setIsProtected(boolean isProtected) { this.isProtected = isProtected; }

    /**
     * Returns the objects.
     * @return the requested result
     */
    public boolean getObjects() { return objects; }
    /**
     * Sets the objects.
     * @param objects objects
     */
    public void setObjects(boolean objects) { this.objects = objects; }

    /**
     * Returns the scenarios.
     * @return the requested result
     */
    public boolean getScenarios() { return scenarios; }
    /**
     * Sets the scenarios.
     * @param scenarios scenarios
     */
    public void setScenarios(boolean scenarios) { this.scenarios = scenarios; }

    /**
     * Returns the format cells.
     * @return the requested result
     */
    public boolean getFormatCells() { return formatCells; }
    /**
     * Sets the format cells.
     * @param formatCells format cells
     */
    public void setFormatCells(boolean formatCells) { this.formatCells = formatCells; }

    /**
     * Returns the format columns.
     * @return the requested result
     */
    public boolean getFormatColumns() { return formatColumns; }
    /**
     * Sets the format columns.
     * @param formatColumns format columns
     */
    public void setFormatColumns(boolean formatColumns) { this.formatColumns = formatColumns; }

    /**
     * Returns the format rows.
     * @return the requested result
     */
    public boolean getFormatRows() { return formatRows; }
    /**
     * Sets the format rows.
     * @param formatRows format rows
     */
    public void setFormatRows(boolean formatRows) { this.formatRows = formatRows; }

    /**
     * Returns the insert columns.
     * @return the requested result
     */
    public boolean getInsertColumns() { return insertColumns; }
    /**
     * Sets the insert columns.
     * @param insertColumns insert columns
     */
    public void setInsertColumns(boolean insertColumns) { this.insertColumns = insertColumns; }

    /**
     * Returns the insert rows.
     * @return the requested result
     */
    public boolean getInsertRows() { return insertRows; }
    /**
     * Sets the insert rows.
     * @param insertRows insert rows
     */
    public void setInsertRows(boolean insertRows) { this.insertRows = insertRows; }

    /**
     * Returns the insert hyperlinks.
     * @return the requested result
     */
    public boolean getInsertHyperlinks() { return insertHyperlinks; }
    /**
     * Sets the insert hyperlinks.
     * @param insertHyperlinks insert hyperlinks
     */
    public void setInsertHyperlinks(boolean insertHyperlinks) { this.insertHyperlinks = insertHyperlinks; }

    /**
     * Returns the delete columns.
     * @return the requested result
     */
    public boolean getDeleteColumns() { return deleteColumns; }
    /**
     * Sets the delete columns.
     * @param deleteColumns delete columns
     */
    public void setDeleteColumns(boolean deleteColumns) { this.deleteColumns = deleteColumns; }

    /**
     * Returns the delete rows.
     * @return the requested result
     */
    public boolean getDeleteRows() { return deleteRows; }
    /**
     * Sets the delete rows.
     * @param deleteRows delete rows
     */
    public void setDeleteRows(boolean deleteRows) { this.deleteRows = deleteRows; }

    /**
     * Returns the select locked cells.
     * @return the requested result
     */
    public boolean getSelectLockedCells() { return selectLockedCells; }
    /**
     * Sets the select locked cells.
     * @param selectLockedCells select locked cells
     */
    public void setSelectLockedCells(boolean selectLockedCells) { this.selectLockedCells = selectLockedCells; }

    /**
     * Returns the sort.
     * @return the requested result
     */
    public boolean getSort() { return sort; }
    /**
     * Sets the sort.
     * @param sort sort
     */
    public void setSort(boolean sort) { this.sort = sort; }

    /**
     * Returns the auto filter.
     * @return the requested result
     */
    public boolean getAutoFilter() { return autoFilter; }
    /**
     * Sets the auto filter.
     * @param autoFilter auto filter
     */
    public void setAutoFilter(boolean autoFilter) { this.autoFilter = autoFilter; }

    /**
     * Returns the pivot tables.
     * @return the requested result
     */
    public boolean getPivotTables() { return pivotTables; }
    /**
     * Sets the pivot tables.
     * @param pivotTables pivot tables
     */
    public void setPivotTables(boolean pivotTables) { this.pivotTables = pivotTables; }

    /**
     * Returns the select unlocked cells.
     * @return the requested result
     */
    public boolean getSelectUnlockedCells() { return selectUnlockedCells; }
    /**
     * Sets the select unlocked cells.
     * @param selectUnlockedCells select unlocked cells
     */
    public void setSelectUnlockedCells(boolean selectUnlockedCells) { this.selectUnlockedCells = selectUnlockedCells; }

    /**
     * Returns the password hash.
     * @return the requested result
     */
    public String getPasswordHash() { return passwordHash; }
    /**
     * Sets the password hash.
     * @param passwordHash password hash
     */
    public void setPasswordHash(String passwordHash) { this.passwordHash = passwordHash; }

    /**
     * Returns the algorithm name.
     * @return the requested result
     */
    public String getAlgorithmName() { return algorithmName; }
    /**
     * Sets the algorithm name.
     * @param algorithmName name to use
     */
    public void setAlgorithmName(String algorithmName) { this.algorithmName = algorithmName; }

    /**
     * Returns the hash value.
     * @return the requested result
     */
    public String getHashValue() { return hashValue; }
    /**
     * Sets the hash value.
     * @param hashValue hash value
     */
    public void setHashValue(String hashValue) { this.hashValue = hashValue; }

    /**
     * Returns the salt value.
     * @return the requested result
     */
    public String getSaltValue() { return saltValue; }
    /**
     * Sets the salt value.
     * @param saltValue salt value
     */
    public void setSaltValue(String saltValue) { this.saltValue = saltValue; }

    /**
     * Returns the spin count.
     * @return the requested result
     */
    public String getSpinCount() { return spinCount; }
    /**
     * Sets the spin count.
     * @param spinCount spin count
     */
    public void setSpinCount(String spinCount) { this.spinCount = spinCount; }

    /**
     * Clears all protection settings to their default values.
     */
    public void clear() {
        isProtected = false;
        objects = false;
        scenarios = false;
        // Walk the current collection so every entry is processed consistently.
        formatCells = false;
        formatColumns = false;
        formatRows = false;
        insertColumns = false;
        insertRows = false;
        insertHyperlinks = false;
        deleteColumns = false;
        deleteRows = false;
        selectLockedCells = false;
        sort = false;
        autoFilter = false;
        pivotTables = false;
        selectUnlockedCells = false;
        passwordHash = null;
        algorithmName = null;
        hashValue = null;
        saltValue = null;
        spinCount = null;
    }

    /**
     * Determines whether any protection state has been stored.
     *
     * @return true if any protection setting is enabled or any string field is non-empty; otherwise false.
     */
    public boolean hasStoredState() {
        return isProtected
            || objects
            || scenarios
            || formatCells
            || formatColumns
            || formatRows
            || insertColumns
            || insertRows
            || insertHyperlinks
            || deleteColumns
            || deleteRows
            || selectLockedCells
            || sort
            || autoFilter
            || pivotTables
            || selectUnlockedCells
            || !isNullOrBlank(passwordHash)
            || !isNullOrBlank(algorithmName)
            || !isNullOrBlank(hashValue)
            || !isNullOrBlank(saltValue)
            || !isNullOrBlank(spinCount);
    }

    /**
     * Indicates whether null or blank.
     * @param s s
     * @return true when the condition is satisfied
     */
    private static boolean isNullOrBlank(String s) {
        return s == null || s.isBlank();
    }
}
