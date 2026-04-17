package com.aspose.cells_foss;

import com.aspose.cells_foss.core.WorkbookPropertiesModel;

/**
 * Represents the properties of a workbook (workbookPr attributes).
 */
public final class WorkbookProperties {

    private final WorkbookPropertiesModel model;

    /**
     * Initializes a new WorkbookProperties instance.
     * @param model model
     */
    WorkbookProperties(WorkbookPropertiesModel model) {
        this.model = model;
    }

    /**
     * Returns the code name.
     * @return the requested result
     */
    public String getCodeName() { return model.getCodeName(); }
    /**
     * Sets the code name.
     * @param value value to apply
     */
    public void setCodeName(String value) { model.setCodeName(value != null ? value : ""); }

    /**
     * Returns the show objects.
     * @return the requested result
     */
    public String getShowObjects() { return model.getShowObjects(); }
    /**
     * Sets the show objects.
     * @param value value to apply
     */
    public void setShowObjects(String value) { model.setShowObjects(value != null ? value : ""); }

    /**
     * Returns the filter privacy.
     * @return the requested result
     */
    public boolean getFilterPrivacy() { return model.getFilterPrivacy(); }
    /**
     * Sets the filter privacy.
     * @param value value to apply
     */
    public void setFilterPrivacy(boolean value) { model.setFilterPrivacy(value); }

    /**
     * Returns the show border unselected tables.
     * @return the requested result
     */
    public boolean getShowBorderUnselectedTables() { return model.getShowBorderUnselectedTables(); }
    /**
     * Sets the show border unselected tables.
     * @param value value to apply
     */
    public void setShowBorderUnselectedTables(boolean value) { model.setShowBorderUnselectedTables(value); }

    /**
     * Returns the show ink annotation.
     * @return the requested result
     */
    public boolean getShowInkAnnotation() { return model.getShowInkAnnotation(); }
    /**
     * Sets the show ink annotation.
     * @param value value to apply
     */
    public void setShowInkAnnotation(boolean value) { model.setShowInkAnnotation(value); }

    /**
     * Returns the backup file.
     * @return the requested result
     */
    public boolean getBackupFile() { return model.getBackupFile(); }
    /**
     * Sets the backup file.
     * @param value value to apply
     */
    public void setBackupFile(boolean value) { model.setBackupFile(value); }

    /**
     * Returns the save external link values.
     * @return the requested result
     */
    public boolean getSaveExternalLinkValues() { return model.getSaveExternalLinkValues(); }
    /**
     * Sets the save external link values.
     * @param value value to apply
     */
    public void setSaveExternalLinkValues(boolean value) { model.setSaveExternalLinkValues(value); }

    /**
     * Returns the update links.
     * @return the requested result
     */
    public String getUpdateLinks() { return model.getUpdateLinks(); }
    /**
     * Sets the update links.
     * @param value value to apply
     */
    public void setUpdateLinks(String value) { model.setUpdateLinks(value != null ? value : ""); }

    /**
     * Returns the hide pivot field list.
     * @return the requested result
     */
    public boolean getHidePivotFieldList() { return model.getHidePivotFieldList(); }
    /**
     * Sets the hide pivot field list.
     * @param value value to apply
     */
    public void setHidePivotFieldList(boolean value) { model.setHidePivotFieldList(value); }

    /**
     * Returns the default theme version.
     * @return the requested result
     */
    public Integer getDefaultThemeVersion() { return model.getDefaultThemeVersion(); }
    /**
     * Sets the default theme version.
     * @param value value to apply
     */
    public void setDefaultThemeVersion(Integer value) { model.setDefaultThemeVersion(value); }

    /** Not yet implemented — returns null. */
    public Object getProtection() { return null; }
    /** Not yet implemented — returns null. */
    public Object getView() { return null; }
    /** Not yet implemented — returns null. */
    public Object getCalculation() { return null; }
}
