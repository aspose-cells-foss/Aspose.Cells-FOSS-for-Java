package com.aspose.cells_foss;

import com.aspose.cells_foss.core.WorkbookModel;
import com.aspose.cells_foss.core.WorkbookPropertiesModel;

/**
 * Represents the properties of a workbook (workbookPr attributes).
 */
public final class WorkbookProperties {

    private final WorkbookPropertiesModel model;
    private final WorkbookProtection protection;
    private final WorkbookView view;
    private final CalculationProperties calculation;

    /**
     * Initializes a new WorkbookProperties instance.
     * @param model properties model
     * @param workbookModel workbook model (needed by WorkbookView for activeTab)
     */
    WorkbookProperties(WorkbookPropertiesModel model, WorkbookModel workbookModel) {
        this.model = model;
        this.protection = new WorkbookProtection(model.getProtection());
        this.view = new WorkbookView(model.getView(), workbookModel);
        this.calculation = new CalculationProperties(model.getCalculation());
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

    /** Returns the workbook-level protection settings (structure, windows, revision). */
    public WorkbookProtection getProtection() { return protection; }

    /** Returns the workbook view/window settings. */
    public WorkbookView getView() { return view; }

    /** Returns the workbook calculation properties. */
    public CalculationProperties getCalculation() { return calculation; }
}
