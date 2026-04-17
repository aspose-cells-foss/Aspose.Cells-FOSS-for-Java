package com.aspose.cells_foss.core;

/**
 * Represents the workbook properties model.
 */
public final class WorkbookPropertiesModel {
    private String codeName = "";
    private String showObjects = "";
    private boolean filterPrivacy;
    private boolean showBorderUnselectedTables = true;
    private boolean showInkAnnotation = true;
    private boolean backupFile;
    private boolean saveExternalLinkValues = true;
    private String updateLinks = "";
    private boolean hidePivotFieldList;
    private Integer defaultThemeVersion;
    private final WorkbookProtectionModel protection;
    private final WorkbookViewModel view;
    private final CalculationPropertiesModel calculation;

    /**
     * Initializes a new WorkbookPropertiesModel instance.
     */
    public WorkbookPropertiesModel() {
        this.protection = new WorkbookProtectionModel();
        this.view = new WorkbookViewModel();
        this.calculation = new CalculationPropertiesModel();
    }

    /**
     * Returns the code name.
     * @return the requested result
     */
    public String getCodeName() { return codeName; }
    /**
     * Sets the code name.
     * @param codeName name to use
     */
    public void setCodeName(String codeName) { this.codeName = codeName; }

    /**
     * Returns the show objects.
     * @return the requested result
     */
    public String getShowObjects() { return showObjects; }
    /**
     * Sets the show objects.
     * @param showObjects show objects
     */
    public void setShowObjects(String showObjects) { this.showObjects = showObjects; }

    /**
     * Returns the filter privacy.
     * @return the requested result
     */
    public boolean getFilterPrivacy() { return filterPrivacy; }
    /**
     * Sets the filter privacy.
     * @param filterPrivacy filter privacy
     */
    public void setFilterPrivacy(boolean filterPrivacy) { this.filterPrivacy = filterPrivacy; }

    /**
     * Returns the show border unselected tables.
     * @return the requested result
     */
    public boolean getShowBorderUnselectedTables() { return showBorderUnselectedTables; }
    /**
     * Sets the show border unselected tables.
     * @param showBorderUnselectedTables show border unselected tables
     */
    public void setShowBorderUnselectedTables(boolean showBorderUnselectedTables) { this.showBorderUnselectedTables = showBorderUnselectedTables; }

    /**
     * Returns the show ink annotation.
     * @return the requested result
     */
    public boolean getShowInkAnnotation() { return showInkAnnotation; }
    /**
     * Sets the show ink annotation.
     * @param showInkAnnotation show ink annotation
     */
    public void setShowInkAnnotation(boolean showInkAnnotation) { this.showInkAnnotation = showInkAnnotation; }

    /**
     * Returns the backup file.
     * @return the requested result
     */
    public boolean getBackupFile() { return backupFile; }
    /**
     * Sets the backup file.
     * @param backupFile backup file
     */
    public void setBackupFile(boolean backupFile) { this.backupFile = backupFile; }

    /**
     * Returns the save external link values.
     * @return the requested result
     */
    public boolean getSaveExternalLinkValues() { return saveExternalLinkValues; }
    /**
     * Sets the save external link values.
     * @param saveExternalLinkValues save external link values
     */
    public void setSaveExternalLinkValues(boolean saveExternalLinkValues) { this.saveExternalLinkValues = saveExternalLinkValues; }

    /**
     * Returns the update links.
     * @return the requested result
     */
    public String getUpdateLinks() { return updateLinks; }
    /**
     * Sets the update links.
     * @param updateLinks update links
     */
    public void setUpdateLinks(String updateLinks) { this.updateLinks = updateLinks; }

    /**
     * Returns the hide pivot field list.
     * @return the requested result
     */
    public boolean getHidePivotFieldList() { return hidePivotFieldList; }
    /**
     * Sets the hide pivot field list.
     * @param hidePivotFieldList hide pivot field list
     */
    public void setHidePivotFieldList(boolean hidePivotFieldList) { this.hidePivotFieldList = hidePivotFieldList; }

    /**
     * Returns the default theme version.
     * @return the requested result
     */
    public Integer getDefaultThemeVersion() { return defaultThemeVersion; }
    /**
     * Sets the default theme version.
     * @param defaultThemeVersion default theme version
     */
    public void setDefaultThemeVersion(Integer defaultThemeVersion) { this.defaultThemeVersion = defaultThemeVersion; }

    /**
     * Returns the protection.
     * @return the requested result
     */
    public WorkbookProtectionModel getProtection() { return protection; }

    /**
     * Returns the view.
     * @return the requested result
     */
    public WorkbookViewModel getView() { return view; }

    /**
     * Returns the calculation.
     * @return the requested result
     */
    public CalculationPropertiesModel getCalculation() { return calculation; }

    /**
     * Copies the from.
     * @param source source
     */
    public void copyFrom(WorkbookPropertiesModel source) {
        codeName = source.codeName;
        showObjects = source.showObjects;
        filterPrivacy = source.filterPrivacy;
        showBorderUnselectedTables = source.showBorderUnselectedTables;
        showInkAnnotation = source.showInkAnnotation;
        backupFile = source.backupFile;
        saveExternalLinkValues = source.saveExternalLinkValues;
        updateLinks = source.updateLinks;
        hidePivotFieldList = source.hidePivotFieldList;
        defaultThemeVersion = source.defaultThemeVersion;
        protection.copyFrom(source.protection);
        view.copyFrom(source.view);
        calculation.copyFrom(source.calculation);
    }

    /**
     * Indicates whether this instance has workbook properties state.
     * @return true when the condition is satisfied
     */
    public boolean hasWorkbookPropertiesState() {
        return !codeName.isBlank()
            || !showObjects.isBlank()
            || filterPrivacy
            || !showBorderUnselectedTables
            || !showInkAnnotation
            || backupFile
            || !saveExternalLinkValues
            || !updateLinks.isBlank()
            || hidePivotFieldList
            || defaultThemeVersion != null;
    }

    // Nested classes to match C# file-level definitions
    /**
     * Represents the WorkbookProtectionModel component.
     */
    public static final class WorkbookProtectionModel {
        private boolean lockStructure;
        private boolean lockWindows;
        private boolean lockRevision;
        private String workbookPassword = "";
        private String revisionsPassword = "";

        /**
         * Returns the lock structure.
         * @return the requested result
         */
        public boolean getLockStructure() { return lockStructure; }
        /**
         * Sets the lock structure.
         * @param lockStructure lock structure
         */
        public void setLockStructure(boolean lockStructure) { this.lockStructure = lockStructure; }

        /**
         * Returns the lock windows.
         * @return the requested result
         */
        public boolean getLockWindows() { return lockWindows; }
        /**
         * Sets the lock windows.
         * @param lockWindows lock windows
         */
        public void setLockWindows(boolean lockWindows) { this.lockWindows = lockWindows; }

        /**
         * Returns the lock revision.
         * @return the requested result
         */
        public boolean getLockRevision() { return lockRevision; }
        /**
         * Sets the lock revision.
         * @param lockRevision lock revision
         */
        public void setLockRevision(boolean lockRevision) { this.lockRevision = lockRevision; }

        /**
         * Returns the workbook password.
         * @return the requested result
         */
        public String getWorkbookPassword() { return workbookPassword; }
        /**
         * Sets the workbook password.
         * @param workbookPassword workbook password
         */
        public void setWorkbookPassword(String workbookPassword) { this.workbookPassword = workbookPassword; }

        /**
         * Returns the revisions password.
         * @return the requested result
         */
        public String getRevisionsPassword() { return revisionsPassword; }
        /**
         * Sets the revisions password.
         * @param revisionsPassword revisions password
         */
        public void setRevisionsPassword(String revisionsPassword) { this.revisionsPassword = revisionsPassword; }

        /**
         * Copies the from.
         * @param source source
         */
        public void copyFrom(WorkbookProtectionModel source) {
            lockStructure = source.lockStructure;
            lockWindows = source.lockWindows;
            lockRevision = source.lockRevision;
            workbookPassword = source.workbookPassword;
            revisionsPassword = source.revisionsPassword;
        }

        /**
         * Indicates whether this instance has stored state.
         * @return true when the condition is satisfied
         */
        public boolean hasStoredState() {
            return lockStructure
                || lockWindows
                || lockRevision
                || !workbookPassword.isBlank()
                || !revisionsPassword.isBlank();
        }
    }

    /**
     * Represents the WorkbookViewModel component.
     */
    public static final class WorkbookViewModel {
        private Integer xWindow;
        private Integer yWindow;
        private Integer windowWidth;
        private Integer windowHeight;
        private Integer firstSheet;
        private Boolean showHorizontalScroll;
        private Boolean showVerticalScroll;
        private Boolean showSheetTabs;
        private Integer tabRatio;
        private String visibility = "";
        private boolean minimized;
        private boolean autoFilterDateGrouping = true;

        /**
         * Returns the x window.
         * @return the requested result
         */
        public Integer getXWindow() { return xWindow; }
        /**
         * Sets the x window.
         * @param xWindow x window
         */
        public void setXWindow(Integer xWindow) { this.xWindow = xWindow; }

        /**
         * Returns the y window.
         * @return the requested result
         */
        public Integer getYWindow() { return yWindow; }
        /**
         * Sets the y window.
         * @param yWindow y window
         */
        public void setYWindow(Integer yWindow) { this.yWindow = yWindow; }

        /**
         * Returns the window width.
         * @return the requested result
         */
        public Integer getWindowWidth() { return windowWidth; }
        /**
         * Sets the window width.
         * @param windowWidth window width
         */
        public void setWindowWidth(Integer windowWidth) { this.windowWidth = windowWidth; }

        /**
         * Returns the window height.
         * @return the requested result
         */
        public Integer getWindowHeight() { return windowHeight; }
        /**
         * Sets the window height.
         * @param windowHeight window height
         */
        public void setWindowHeight(Integer windowHeight) { this.windowHeight = windowHeight; }

        /**
         * Returns the first sheet.
         * @return the requested result
         */
        public Integer getFirstSheet() { return firstSheet; }
        /**
         * Sets the first sheet.
         * @param firstSheet first sheet
         */
        public void setFirstSheet(Integer firstSheet) { this.firstSheet = firstSheet; }

        /**
         * Returns the show horizontal scroll.
         * @return the requested result
         */
        public Boolean getShowHorizontalScroll() { return showHorizontalScroll; }
        /**
         * Sets the show horizontal scroll.
         * @param showHorizontalScroll show horizontal scroll
         */
        public void setShowHorizontalScroll(Boolean showHorizontalScroll) { this.showHorizontalScroll = showHorizontalScroll; }

        /**
         * Returns the show vertical scroll.
         * @return the requested result
         */
        public Boolean getShowVerticalScroll() { return showVerticalScroll; }
        /**
         * Sets the show vertical scroll.
         * @param showVerticalScroll show vertical scroll
         */
        public void setShowVerticalScroll(Boolean showVerticalScroll) { this.showVerticalScroll = showVerticalScroll; }

        /**
         * Returns the show sheet tabs.
         * @return the requested result
         */
        public Boolean getShowSheetTabs() { return showSheetTabs; }
        /**
         * Sets the show sheet tabs.
         * @param showSheetTabs show sheet tabs
         */
        public void setShowSheetTabs(Boolean showSheetTabs) { this.showSheetTabs = showSheetTabs; }

        /**
         * Returns the tab ratio.
         * @return the requested result
         */
        public Integer getTabRatio() { return tabRatio; }
        /**
         * Sets the tab ratio.
         * @param tabRatio tab ratio
         */
        public void setTabRatio(Integer tabRatio) { this.tabRatio = tabRatio; }

        /**
         * Returns the visibility.
         * @return the requested result
         */
        public String getVisibility() { return visibility; }
        /**
         * Sets the visibility.
         * @param visibility visibility
         */
        public void setVisibility(String visibility) { this.visibility = visibility; }

        /**
         * Returns the minimized.
         * @return the requested result
         */
        public boolean getMinimized() { return minimized; }
        /**
         * Sets the minimized.
         * @param minimized minimized
         */
        public void setMinimized(boolean minimized) { this.minimized = minimized; }

        /**
         * Returns the auto filter date grouping.
         * @return the requested result
         */
        public boolean getAutoFilterDateGrouping() { return autoFilterDateGrouping; }
        /**
         * Sets the auto filter date grouping.
         * @param autoFilterDateGrouping auto filter date grouping
         */
        public void setAutoFilterDateGrouping(boolean autoFilterDateGrouping) { this.autoFilterDateGrouping = autoFilterDateGrouping; }

        /**
         * Copies the from.
         * @param source source
         */
        public void copyFrom(WorkbookViewModel source) {
            xWindow = source.xWindow;
            yWindow = source.yWindow;
            windowWidth = source.windowWidth;
            windowHeight = source.windowHeight;
            firstSheet = source.firstSheet;
            showHorizontalScroll = source.showHorizontalScroll;
            showVerticalScroll = source.showVerticalScroll;
            showSheetTabs = source.showSheetTabs;
            tabRatio = source.tabRatio;
            visibility = source.visibility;
            minimized = source.minimized;
            autoFilterDateGrouping = source.autoFilterDateGrouping;
        }

        /**
         * Indicates whether has stored state.
         * @param activeSheetIndex zero-based active sheet index
         * @return true when the condition is satisfied
         */
        public boolean hasStoredState(int activeSheetIndex) {
            return activeSheetIndex > 0
                || xWindow != null
                || yWindow != null
                || windowWidth != null
                || windowHeight != null
                || firstSheet != null
                || showHorizontalScroll != null
                || showVerticalScroll != null
                || showSheetTabs != null
                || tabRatio != null
                || !visibility.isBlank()
                || minimized
                || !autoFilterDateGrouping;
        }
    }

    /**
     * Represents the CalculationPropertiesModel component.
     */
    public static final class CalculationPropertiesModel {
        private Integer calculationId;
        private String calculationMode = "";
        private boolean fullCalculationOnLoad;
        private String referenceMode = "";
        private boolean iterate;
        private Integer iterateCount;
        private Double iterateDelta;
        private Boolean fullPrecision;
        private Boolean calculationCompleted;
        private Boolean calculationOnSave;
        private Boolean concurrentCalculation;
        private boolean forceFullCalculation;

        /**
         * Returns the calculation id.
         * @return the requested result
         */
        public Integer getCalculationId() { return calculationId; }
        /**
         * Sets the calculation id.
         * @param calculationId calculation id
         */
        public void setCalculationId(Integer calculationId) { this.calculationId = calculationId; }

        /**
         * Returns the calculation mode.
         * @return the requested result
         */
        public String getCalculationMode() { return calculationMode; }
        /**
         * Sets the calculation mode.
         * @param calculationMode calculation mode
         */
        public void setCalculationMode(String calculationMode) { this.calculationMode = calculationMode; }

        /**
         * Returns the full calculation on load.
         * @return the requested result
         */
        public boolean getFullCalculationOnLoad() { return fullCalculationOnLoad; }
        /**
         * Sets the full calculation on load.
         * @param fullCalculationOnLoad full calculation on load
         */
        public void setFullCalculationOnLoad(boolean fullCalculationOnLoad) { this.fullCalculationOnLoad = fullCalculationOnLoad; }

        /**
         * Returns the reference mode.
         * @return the requested result
         */
        public String getReferenceMode() { return referenceMode; }
        /**
         * Sets the reference mode.
         * @param referenceMode reference mode
         */
        public void setReferenceMode(String referenceMode) { this.referenceMode = referenceMode; }

        /**
         * Returns the iterate.
         * @return the requested result
         */
        public boolean getIterate() { return iterate; }
        /**
         * Sets the iterate.
         * @param iterate iterate
         */
        public void setIterate(boolean iterate) { this.iterate = iterate; }

        /**
         * Returns the iterate count.
         * @return the requested result
         */
        public Integer getIterateCount() { return iterateCount; }
        /**
         * Sets the iterate count.
         * @param iterateCount iterate count
         */
        public void setIterateCount(Integer iterateCount) { this.iterateCount = iterateCount; }

        /**
         * Returns the iterate delta.
         * @return the requested result
         */
        public Double getIterateDelta() { return iterateDelta; }
        /**
         * Sets the iterate delta.
         * @param iterateDelta iterate delta
         */
        public void setIterateDelta(Double iterateDelta) { this.iterateDelta = iterateDelta; }

        /**
         * Returns the full precision.
         * @return the requested result
         */
        public Boolean getFullPrecision() { return fullPrecision; }
        /**
         * Sets the full precision.
         * @param fullPrecision full precision
         */
        public void setFullPrecision(Boolean fullPrecision) { this.fullPrecision = fullPrecision; }

        /**
         * Returns the calculation completed.
         * @return the requested result
         */
        public Boolean getCalculationCompleted() { return calculationCompleted; }
        /**
         * Sets the calculation completed.
         * @param calculationCompleted calculation completed
         */
        public void setCalculationCompleted(Boolean calculationCompleted) { this.calculationCompleted = calculationCompleted; }

        /**
         * Returns the calculation on save.
         * @return the requested result
         */
        public Boolean getCalculationOnSave() { return calculationOnSave; }
        /**
         * Sets the calculation on save.
         * @param calculationOnSave calculation on save
         */
        public void setCalculationOnSave(Boolean calculationOnSave) { this.calculationOnSave = calculationOnSave; }

        /**
         * Returns the concurrent calculation.
         * @return the requested result
         */
        public Boolean getConcurrentCalculation() { return concurrentCalculation; }
        /**
         * Sets the concurrent calculation.
         * @param concurrentCalculation concurrent calculation
         */
        public void setConcurrentCalculation(Boolean concurrentCalculation) { this.concurrentCalculation = concurrentCalculation; }

        /**
         * Returns the force full calculation.
         * @return the requested result
         */
        public boolean getForceFullCalculation() { return forceFullCalculation; }
        /**
         * Sets the force full calculation.
         * @param forceFullCalculation force full calculation
         */
        public void setForceFullCalculation(boolean forceFullCalculation) { this.forceFullCalculation = forceFullCalculation; }

        /**
         * Copies the from.
         * @param source source
         */
        public void copyFrom(CalculationPropertiesModel source) {
            calculationId = source.calculationId;
            calculationMode = source.calculationMode;
            fullCalculationOnLoad = source.fullCalculationOnLoad;
            referenceMode = source.referenceMode;
            iterate = source.iterate;
            iterateCount = source.iterateCount;
            iterateDelta = source.iterateDelta;
            fullPrecision = source.fullPrecision;
            calculationCompleted = source.calculationCompleted;
            calculationOnSave = source.calculationOnSave;
            concurrentCalculation = source.concurrentCalculation;
            // Walk the current collection so every entry is processed consistently.
            forceFullCalculation = source.forceFullCalculation;
        }

        /**
         * Indicates whether this instance has stored state.
         * @return true when the condition is satisfied
         */
        public boolean hasStoredState() {
            return calculationId != null
                || !calculationMode.isBlank()
                || fullCalculationOnLoad
                || !referenceMode.isBlank()
                || iterate
                || iterateCount != null
                || iterateDelta != null
                || fullPrecision != null
                || calculationCompleted != null
                || calculationOnSave != null
                || concurrentCalculation != null
                || forceFullCalculation;
        }
    }
}