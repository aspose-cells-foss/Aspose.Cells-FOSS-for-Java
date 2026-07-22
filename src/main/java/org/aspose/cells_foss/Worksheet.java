package org.aspose.cells_foss;

import org.aspose.cells_foss.core.CellAddress;
import org.aspose.cells_foss.core.WorksheetModel;

/**
 * Represents a worksheet in a workbook.
 */
public class Worksheet {
    private final Workbook workbook;
    private final WorksheetModel model;
    private final Cells cells;
    private final WorksheetProtection protection;
    private final PageSetup pageSetup;
    private final ValidationCollection validations;
    private final HyperlinkCollection hyperlinks;
    private final AutoFilter autoFilter;
    private final ConditionalFormattingCollection conditionalFormattings;
    private final CommentCollection comments;
    private final PictureCollection pictures;
    private final ShapeCollection shapes;
    private final ChartCollection charts;
    private final ListObjectCollection listObjects;

    /**
     * Initializes a new Worksheet instance.
     * @param workbook workbook to apply
     * @param model model
     */
    Worksheet(Workbook workbook, WorksheetModel model) {
        this.workbook = workbook;
        this.model = model;
        this.cells = new Cells(this);
        this.protection = new WorksheetProtection(model.getProtection());
        this.pageSetup = new PageSetup(model.getPageSetup());
        this.validations = new ValidationCollection(model.getValidations());
        this.hyperlinks = new HyperlinkCollection(model.getHyperlinks());
        this.autoFilter = new AutoFilter(model.getAutoFilter());
        this.conditionalFormattings = new ConditionalFormattingCollection(model.getConditionalFormattings());
        this.comments = new CommentCollection(model.getComments());
        this.pictures = new PictureCollection(model.getPictures());
        this.shapes = new ShapeCollection(model.getShapes());
        this.charts = new ChartCollection(model.getCharts());
        this.listObjects = new ListObjectCollection(model.getListObjects());
        this.listObjects.setCells(this.cells);
    }

    /**
     * Returns the workbook.
     * @return the requested result
     */
    Workbook getWorkbook() { return workbook; }

    /**
     * Returns the model.
     * @return the requested result
     */
    WorksheetModel getModel() { return model; }

    /**
     * Returns the name.
     * @return the requested result
     */
    public String getName() { return model.getName(); }

    /**
     * Sets the name.
     * @param name name
     */
    public void setName(String name) {
        // Handle the relevant branch before the state changes.
        if (name == null || name.isBlank()) {
            throw new CellsException("Worksheet name must be non-empty.");
        }
        workbook.ensureUniqueSheetName(name, model);
        model.setName(name);
    }

    /**
     * Returns the visibility type.
     * @return the requested result
     */
    public VisibilityType getVisibilityType() {
        // Translate the internal value into the matching public representation.
        switch (model.getVisibility()) {
            case HIDDEN: return VisibilityType.HIDDEN;
            case VERY_HIDDEN: return VisibilityType.VERY_HIDDEN;
            default: return VisibilityType.VISIBLE;
        }
    }

    /**
     * Sets the visibility type.
     * @param visibilityType visibility type
     */
    public void setVisibilityType(VisibilityType visibilityType) {
        // Translate the internal value into the matching public representation.
        switch (visibilityType) {
            case HIDDEN:
                model.setVisibility(org.aspose.cells_foss.core.SheetVisibility.HIDDEN); break;
            case VERY_HIDDEN:
                model.setVisibility(org.aspose.cells_foss.core.SheetVisibility.VERY_HIDDEN); break;
            default:
                model.setVisibility(org.aspose.cells_foss.core.SheetVisibility.VISIBLE); break;
        }
    }

    /**
     * Returns the tab color.
     * @return the requested result
     */
    public Color getTabColor() {
        return model.getTabColor() != null ? Color.fromCore(model.getTabColor()) : Color.getEmpty();
    }

    /**
     * Sets the tab color.
     * @param value value to apply
     */
    public void setTabColor(Color value) {
        model.setTabColor(value.equals(Color.getEmpty()) ? null : value.toCore());
    }

    /**
     * Returns the show gridlines.
     * @return the requested result
     */
    public boolean getShowGridlines() { return model.getView().getShowGridLines(); }
    /**
     * Sets the show gridlines.
     * @param value value to apply
     */
    public void setShowGridlines(boolean value) { model.getView().setShowGridLines(value); }

    /**
     * Returns the show row column headers.
     * @return the requested result
     */
    public boolean getShowRowColumnHeaders() { return model.getView().getShowRowColumnHeaders(); }
    /**
     * Sets the show row column headers.
     * @param value value to apply
     */
    public void setShowRowColumnHeaders(boolean value) { model.getView().setShowRowColumnHeaders(value); }

    /**
     * Returns the show zeros.
     * @return the requested result
     */
    public boolean getShowZeros() { return model.getView().getShowZeros(); }
    /**
     * Sets the show zeros.
     * @param value value to apply
     */
    public void setShowZeros(boolean value) { model.getView().setShowZeros(value); }

    /**
     * Returns the right to left.
     * @return the requested result
     */
    public boolean getRightToLeft() { return model.getView().getRightToLeft(); }
    /**
     * Sets the right to left.
     * @param value value to apply
     */
    public void setRightToLeft(boolean value) { model.getView().setRightToLeft(value); }

    /**
     * Returns the zoom.
     * @return the requested result
     */
    public int getZoom() { return model.getView().getZoomScale(); }
    /**
     * Sets the zoom.
     * @param value value to apply
     */
    public void setZoom(int value) {
        // Validate the caller input before continuing.
        if (value < 10 || value > 400) throw new CellsException("Zoom must be between 10 and 400.");
        model.getView().setZoomScale(value);
    }

    /**
     * Returns the cells.
     * @return the requested result
     */
    public Cells getCells() { return cells; }

    /**
     * Returns the protection.
     * @return the requested result
     */
    public WorksheetProtection getProtection() { return protection; }

    /**
     * Returns the page setup.
     * @return the requested result
     */
    public PageSetup getPageSetup() { return pageSetup; }

    /**
     * Returns the validations.
     * @return the requested result
     */
    public ValidationCollection getValidations() { return validations; }

    /**
     * Returns the hyperlinks.
     * @return the requested result
     */
    public HyperlinkCollection getHyperlinks() { return hyperlinks; }

    /**
     * Returns the auto filter.
     * @return the requested result
     */
    public AutoFilter getAutoFilter() { return autoFilter; }

    /**
     * Returns the conditional formattings.
     * @return the requested result
     */
    public ConditionalFormattingCollection getConditionalFormattings() { return conditionalFormattings; }

    /** Returns the cell comments collection. */
    public CommentCollection getComments() { return comments; }

    /** Returns the embedded pictures collection. */
    public PictureCollection getPictures() { return pictures; }
    /** Returns the collection of drawing objects (shapes) on this worksheet. */
    public ShapeCollection getShapes() { return shapes; }

    /** Returns the embedded charts collection. */
    public ChartCollection getCharts() { return charts; }

    /** Returns the Excel tables (ListObjects) collection. */
    public ListObjectCollection getListObjects() { return listObjects; }

    // =========================================================================
    // Freeze panes
    // =========================================================================

    /**
     * Freezes panes at the specified cell in the worksheet.
     * @param row zero-based row index of the first unfrozen row
     * @param column zero-based column index of the first unfrozen column
     * @param freezedRows number of rows to freeze from the top
     * @param freezedCols number of columns to freeze from the left
     */
    public void freezePanes(int row, int column, int freezedRows, int freezedCols) {
        if (freezedRows < 0) throw new CellsException("freezedRows must be non-negative.");
        if (freezedCols < 0) throw new CellsException("freezedCols must be non-negative.");
        model.getView().setFreezeRow(freezedRows);
        model.getView().setFreezeColumn(freezedCols);
    }

    /**
     * Freezes panes at the specified cell in the worksheet.
     * @param cellName name of the top-left cell of the unfrozen area (e.g. "B2")
     * @param freezedRows number of rows to freeze from the top
     * @param freezedCols number of columns to freeze from the left
     */
    public void freezePanes(String cellName, int freezedRows, int freezedCols) {
        if (cellName == null || cellName.isBlank()) throw new CellsException("cellName must be non-empty.");
        CellAddress addr = CellAddress.parse(cellName);
        freezePanes(addr.getRowIndex(), addr.getColumnIndex(), freezedRows, freezedCols);
    }

    /**
     * Unfreezes all frozen panes in the worksheet.
     */
    public void unFreezePanes() {
        model.getView().setFreezeRow(0);
        model.getView().setFreezeColumn(0);
    }

    /**
     * Returns whether any panes are currently frozen.
     * @return the requested result
     */
    public boolean isFrozen() {
        return model.getView().getFreezeRow() > 0 || model.getView().getFreezeColumn() > 0;
    }

    /**
     * Returns the number of rows frozen from the top. Zero means no row freeze.
     * @return the requested result
     */
    public int getFreezedRows() { return model.getView().getFreezeRow(); }

    /**
     * Returns the number of columns frozen from the left. Zero means no column freeze.
     * @return the requested result
     */
    public int getFreezedColumns() { return model.getView().getFreezeColumn(); }

    /**
     * Enables protection for the current object.
     */
    public void protect() {
        model.getProtection().setIsProtected(true);
    }

    /**
     * Disables protection for the current object.
     */
    public void unprotect() {
        model.getProtection().clear();
    }
}

