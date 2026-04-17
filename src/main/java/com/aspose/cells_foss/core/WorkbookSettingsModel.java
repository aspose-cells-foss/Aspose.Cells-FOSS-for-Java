package com.aspose.cells_foss.core;

import java.util.Locale;

/**
 * Represents workbook settings model.
 */
public final class WorkbookSettingsModel {
    private DateSystem dateSystem = DateSystem.WINDOWS_1900;
    private Locale displayCulture = Locale.ROOT;

    /**
     * Returns the date system.
     * @return the requested result
     */
    public DateSystem getDateSystem() { return dateSystem; }
    /**
     * Sets the date system.
     * @param dateSystem date system
     */
    public void setDateSystem(DateSystem dateSystem) { this.dateSystem = dateSystem; }

    /**
     * Returns the display culture.
     * @return the requested result
     */
    public Locale getDisplayCulture() { return displayCulture; }
    /**
     * Sets the display culture.
     * @param displayCulture display culture
     */
    public void setDisplayCulture(Locale displayCulture) { this.displayCulture = displayCulture; }
}