package com.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Represents the model of a worksheet in the Excel file.
 */
public final class WorksheetModel {
    private String name;
    private SheetVisibility visibility = SheetVisibility.VISIBLE;
    private final Map<CellAddress, CellRecord> cells = new HashMap<>();
    private final Map<Integer, RowModel> rows = new HashMap<>();
    private final List<ColumnRangeModel> columns = new ArrayList<>();
    private final List<MergeRegion> mergeRegions = new ArrayList<>();
    private final List<HyperlinkModel> hyperlinks = new ArrayList<>();
    private final List<ValidationModel> validations = new ArrayList<>();
    private final List<ConditionalFormattingModel> conditionalFormattings = new ArrayList<>();
    private final PageSetupModel pageSetup = new PageSetupModel();
    private final WorksheetViewModel view = new WorksheetViewModel();
    private final WorksheetProtectionModel protection = new WorksheetProtectionModel();
    private final AutoFilterModel autoFilter = new AutoFilterModel();
    private ColorValue tabColor;

    /**
     * Initializes a new WorksheetModel instance.
     * @param name name
     */
    public WorksheetModel(String name) {
        this.name = name;
    }

    /**
     * Returns the name.
     * @return the requested result
     */
    public String getName() { return name; }
    /**
     * Sets the name.
     * @param name name
     */
    public void setName(String name) { this.name = name; }

    /**
     * Returns the visibility.
     * @return the requested result
     */
    public SheetVisibility getVisibility() { return visibility; }
    /**
     * Sets the visibility.
     * @param visibility visibility
     */
    public void setVisibility(SheetVisibility visibility) { this.visibility = visibility; }

    /**
     * Returns the cells.
     * @return the requested result
     */
    public Map<CellAddress, CellRecord> getCells() { return cells; }

    /**
     * Returns the rows.
     * @return the requested result
     */
    public Map<Integer, RowModel> getRows() { return rows; }

    /**
     * Returns the columns.
     * @return the requested result
     */
    public List<ColumnRangeModel> getColumns() { return columns; }

    /**
     * Returns the merge regions.
     * @return the requested result
     */
    public List<MergeRegion> getMergeRegions() { return mergeRegions; }

    /**
     * Returns the hyperlinks.
     * @return the requested result
     */
    public List<HyperlinkModel> getHyperlinks() { return hyperlinks; }

    /**
     * Returns the validations.
     * @return the requested result
     */
    public List<ValidationModel> getValidations() { return validations; }

    /**
     * Returns the conditional formattings.
     * @return the requested result
     */
    public List<ConditionalFormattingModel> getConditionalFormattings() { return conditionalFormattings; }

    /**
     * Returns the page setup.
     * @return the requested result
     */
    public PageSetupModel getPageSetup() { return pageSetup; }

    /**
     * Returns the view.
     * @return the requested result
     */
    public WorksheetViewModel getView() { return view; }

    /**
     * Returns the protection.
     * @return the requested result
     */
    public WorksheetProtectionModel getProtection() { return protection; }

    /**
     * Returns the auto filter.
     * @return the requested result
     */
    public AutoFilterModel getAutoFilter() { return autoFilter; }

    /**
     * Returns the tab color.
     * @return the requested result
     */
    public ColorValue getTabColor() { return tabColor; }
    /**
     * Sets the tab color.
     * @param tabColor tab color
     */
    public void setTabColor(ColorValue tabColor) { this.tabColor = tabColor; }
}