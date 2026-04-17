package com.aspose.cells_foss;

import com.aspose.cells_foss.core.WorkbookSettingsModel;

/**
 * Represents workbook settings for an Excel file.
 */
public final class WorkbookSettings {
    private final WorkbookSettingsModel model;

    /**
     * Initializes a new WorkbookSettings instance.
     * @param model model
     */
    WorkbookSettings(WorkbookSettingsModel model) {
        this.model = model;
    }

    /**
     * Gets or sets a value indicating whether the workbook uses the 1904 date system.
     */
    public boolean getDate1904() {
        return model.getDateSystem() == com.aspose.cells_foss.core.DateSystem.MAC_1904;
    }

    /**
     * Sets the date 1904.
     * @param value value to apply
     */
    public void setDate1904(boolean value) {
        model.setDateSystem(value ? com.aspose.cells_foss.core.DateSystem.MAC_1904 : com.aspose.cells_foss.core.DateSystem.WINDOWS_1900);
    }

    /**
     * Gets or sets the culture used for display.
     */
    public java.util.Locale getCulture() {
        return (java.util.Locale) model.getDisplayCulture().clone();
    }

    /**
     * Sets the culture.
     * @param value value to apply
     */
    public void setCulture(java.util.Locale value) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            throw new IllegalArgumentException("value");
        }
        model.setDisplayCulture((java.util.Locale) value.clone());
    }
}