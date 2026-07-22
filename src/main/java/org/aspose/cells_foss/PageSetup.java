package org.aspose.cells_foss;

import org.aspose.cells_foss.core.PageSetupModel;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * Represents page setup options for a worksheet.
 */
public final class PageSetup {
    private static final double CENTIMETERS_PER_INCH = 2.54;
    private final PageSetupModel model;

    /**
     * Initializes a new PageSetup instance.
     * @param model model
     */
    PageSetup(PageSetupModel model) {
        this.model = model;
    }

    /**
     * Returns the paper size.
     * @return the requested result
     */
    public PaperSizeType getPaperSize() {
        return PaperSizeType.values()[model.getPaperSize()];
    }

    /**
     * Sets the paper size.
     * @param value value to apply
     */
    public void setPaperSize(PaperSizeType value) {
        model.setPaperSize(value.ordinal());
    }

    /**
     * Returns the orientation.
     * @return the requested result
     */
    public PageOrientationType getOrientation() {
        // Translate the internal value into the matching public representation.
        switch (model.getOrientation()) {
            case PORTRAIT:
                return PageOrientationType.PORTRAIT;
            case LANDSCAPE:
                return PageOrientationType.LANDSCAPE;
            default:
                return PageOrientationType.DEFAULT;
        }
    }

    /**
     * Sets the orientation.
     * @param value value to apply
     */
    public void setOrientation(PageOrientationType value) {
        // Translate the internal value into the matching public representation.
        switch (value) {
            case PORTRAIT:
                model.setOrientation(org.aspose.cells_foss.core.PageOrientation.PORTRAIT);
                break;
            case LANDSCAPE:
                model.setOrientation(org.aspose.cells_foss.core.PageOrientation.LANDSCAPE);
                break;
            default:
                model.setOrientation(org.aspose.cells_foss.core.PageOrientation.DEFAULT);
                break;
        }
    }

    /**
     * Returns the first page number.
     * @return the requested result
     */
    public Integer getFirstPageNumber() {
        return model.getFirstPageNumber();
    }

    /**
     * Sets the first page number.
     * @param value value to apply
     */
    public void setFirstPageNumber(Integer value) {
        // Handle the relevant branch before the state changes.
        if (value != null && value <= 0) {
            throw new CellsException("FirstPageNumber must be positive.");
        }
        model.setFirstPageNumber(value);
    }

    /**
     * Returns the scale.
     * @return the requested result
     */
    public Integer getScale() {
        return model.getScale();
    }

    /**
     * Sets the scale.
     * @param value value to apply
     */
    public void setScale(Integer value) {
        // Handle the relevant branch before the state changes.
        if (value != null && (value < 10 || value > 400)) {
            throw new CellsException("Scale must be between 10 and 400.");
        }
        model.setScale(value);
    }

    /**
     * Returns the fit to pages wide.
     * @return the requested result
     */
    public Integer getFitToPagesWide() {
        return model.getFitToWidth();
    }

    /**
     * Sets the fit to pages wide.
     * @param value value to apply
     */
    public void setFitToPagesWide(Integer value) {
        // Handle the relevant branch before the state changes.
        if (value != null && value < 0) {
            throw new CellsException("FitToPagesWide must be zero or greater.");
        }
        model.setFitToWidth(value);
    }

    /**
     * Returns the fit to pages tall.
     * @return the requested result
     */
    public Integer getFitToPagesTall() {
        return model.getFitToHeight();
    }

    /**
     * Sets the fit to pages tall.
     * @param value value to apply
     */
    public void setFitToPagesTall(Integer value) {
        // Handle the relevant branch before the state changes.
        if (value != null && value < 0) {
            throw new CellsException("FitToPagesTall must be zero or greater.");
        }
        model.setFitToHeight(value);
    }

    /**
     * Returns the print area.
     * @return the requested result
     */
    public String getPrintArea() {
        String value = model.getPrintArea();
        return value != null ? value : "";
    }

    /**
     * Sets the print area.
     * @param value value to apply
     */
    public void setPrintArea(String value) {
        model.setPrintArea(normalizeText(value));
    }

    /**
     * Returns the print title rows.
     * @return the requested result
     */
    public String getPrintTitleRows() {
        String value = model.getPrintTitleRows();
        return value != null ? value : "";
    }

    /**
     * Sets the print title rows.
     * @param value value to apply
     */
    public void setPrintTitleRows(String value) {
        model.setPrintTitleRows(normalizeText(value));
    }

    /**
     * Returns the print title columns.
     * @return the requested result
     */
    public String getPrintTitleColumns() {
        String value = model.getPrintTitleColumns();
        return value != null ? value : "";
    }

    /**
     * Sets the print title columns.
     * @param value value to apply
     */
    public void setPrintTitleColumns(String value) {
        model.setPrintTitleColumns(normalizeText(value));
    }

    /**
     * Returns the left margin.
     * @return the requested result
     */
    public double getLeftMargin() {
        return toCentimeters(model.getMargins().getLeft());
    }

    /**
     * Sets the left margin.
     * @param value value to apply
     */
    public void setLeftMargin(double value) {
        model.getMargins().setLeft(validateMargin(toInches(value), "LeftMargin"));
    }

    /**
     * Returns the right margin.
     * @return the requested result
     */
    public double getRightMargin() {
        return toCentimeters(model.getMargins().getRight());
    }

    /**
     * Sets the right margin.
     * @param value value to apply
     */
    public void setRightMargin(double value) {
        model.getMargins().setRight(validateMargin(toInches(value), "RightMargin"));
    }

    /**
     * Returns the top margin.
     * @return the requested result
     */
    public double getTopMargin() {
        return toCentimeters(model.getMargins().getTop());
    }

    /**
     * Sets the top margin.
     * @param value value to apply
     */
    public void setTopMargin(double value) {
        model.getMargins().setTop(validateMargin(toInches(value), "TopMargin"));
    }

    /**
     * Returns the bottom margin.
     * @return the requested result
     */
    public double getBottomMargin() {
        return toCentimeters(model.getMargins().getBottom());
    }

    /**
     * Sets the bottom margin.
     * @param value value to apply
     */
    public void setBottomMargin(double value) {
        model.getMargins().setBottom(validateMargin(toInches(value), "BottomMargin"));
    }

    /**
     * Returns the header margin.
     * @return the requested result
     */
    public double getHeaderMargin() {
        return toCentimeters(model.getMargins().getHeader());
    }

    /**
     * Sets the header margin.
     * @param value value to apply
     */
    public void setHeaderMargin(double value) {
        model.getMargins().setHeader(validateMargin(toInches(value), "HeaderMargin"));
    }

    /**
     * Returns the footer margin.
     * @return the requested result
     */
    public double getFooterMargin() {
        return toCentimeters(model.getMargins().getFooter());
    }

    /**
     * Sets the footer margin.
     * @param value value to apply
     */
    public void setFooterMargin(double value) {
        model.getMargins().setFooter(validateMargin(toInches(value), "FooterMargin"));
    }

    /**
     * Returns the left margin inch.
     * @return the requested result
     */
    public double getLeftMarginInch() {
        return model.getMargins().getLeft();
    }

    /**
     * Sets the left margin inch.
     * @param value value to apply
     */
    public void setLeftMarginInch(double value) {
        model.getMargins().setLeft(validateMargin(value, "LeftMarginInch"));
    }

    /**
     * Returns the right margin inch.
     * @return the requested result
     */
    public double getRightMarginInch() {
        return model.getMargins().getRight();
    }

    /**
     * Sets the right margin inch.
     * @param value value to apply
     */
    public void setRightMarginInch(double value) {
        model.getMargins().setRight(validateMargin(value, "RightMarginInch"));
    }

    /**
     * Returns the top margin inch.
     * @return the requested result
     */
    public double getTopMarginInch() {
        return model.getMargins().getTop();
    }

    /**
     * Sets the top margin inch.
     * @param value value to apply
     */
    public void setTopMarginInch(double value) {
        model.getMargins().setTop(validateMargin(value, "TopMarginInch"));
    }

    /**
     * Returns the bottom margin inch.
     * @return the requested result
     */
    public double getBottomMarginInch() {
        return model.getMargins().getBottom();
    }

    /**
     * Sets the bottom margin inch.
     * @param value value to apply
     */
    public void setBottomMarginInch(double value) {
        model.getMargins().setBottom(validateMargin(value, "BottomMarginInch"));
    }

    /**
     * Returns the header margin inch.
     * @return the requested result
     */
    public double getHeaderMarginInch() {
        return model.getMargins().getHeader();
    }

    /**
     * Sets the header margin inch.
     * @param value value to apply
     */
    public void setHeaderMarginInch(double value) {
        model.getMargins().setHeader(validateMargin(value, "HeaderMarginInch"));
    }

    /**
     * Returns the footer margin inch.
     * @return the requested result
     */
    public double getFooterMarginInch() {
        return model.getMargins().getFooter();
    }

    /**
     * Sets the footer margin inch.
     * @param value value to apply
     */
    public void setFooterMarginInch(double value) {
        model.getMargins().setFooter(validateMargin(value, "FooterMarginInch"));
    }

    /**
     * Returns the left header.
     * @return the requested result
     */
    public String getLeftHeader() {
        String value = model.getHeaderFooter().getLeftHeader();
        return value != null ? value : "";
    }

    /**
     * Sets the left header.
     * @param value value to apply
     */
    public void setLeftHeader(String value) {
        model.getHeaderFooter().setLeftHeader(normalizeText(value));
    }

    /**
     * Returns the center header.
     * @return the requested result
     */
    public String getCenterHeader() {
        String value = model.getHeaderFooter().getCenterHeader();
        return value != null ? value : "";
    }

    /**
     * Sets the center header.
     * @param value value to apply
     */
    public void setCenterHeader(String value) {
        model.getHeaderFooter().setCenterHeader(normalizeText(value));
    }

    /**
     * Returns the right header.
     * @return the requested result
     */
    public String getRightHeader() {
        String value = model.getHeaderFooter().getRightHeader();
        return value != null ? value : "";
    }

    /**
     * Sets the right header.
     * @param value value to apply
     */
    public void setRightHeader(String value) {
        model.getHeaderFooter().setRightHeader(normalizeText(value));
    }

    /**
     * Returns the left footer.
     * @return the requested result
     */
    public String getLeftFooter() {
        String value = model.getHeaderFooter().getLeftFooter();
        return value != null ? value : "";
    }

    /**
     * Sets the left footer.
     * @param value value to apply
     */
    public void setLeftFooter(String value) {
        model.getHeaderFooter().setLeftFooter(normalizeText(value));
    }

    /**
     * Returns the center footer.
     * @return the requested result
     */
    public String getCenterFooter() {
        String value = model.getHeaderFooter().getCenterFooter();
        return value != null ? value : "";
    }

    /**
     * Sets the center footer.
     * @param value value to apply
     */
    public void setCenterFooter(String value) {
        model.getHeaderFooter().setCenterFooter(normalizeText(value));
    }

    /**
     * Returns the right footer.
     * @return the requested result
     */
    public String getRightFooter() {
        String value = model.getHeaderFooter().getRightFooter();
        return value != null ? value : "";
    }

    /**
     * Sets the right footer.
     * @param value value to apply
     */
    public void setRightFooter(String value) {
        model.getHeaderFooter().setRightFooter(normalizeText(value));
    }

    /**
     * Returns the print gridlines.
     * @return the requested result
     */
    public boolean getPrintGridlines() {
        return model.getPrintOptions().getGridLines();
    }

    /**
     * Sets the print gridlines.
     * @param value value to apply
     */
    public void setPrintGridlines(boolean value) {
        model.getPrintOptions().setGridLines(value);
    }

    /**
     * Returns the print headings.
     * @return the requested result
     */
    public boolean getPrintHeadings() {
        return model.getPrintOptions().getHeadings();
    }

    /**
     * Sets the print headings.
     * @param value value to apply
     */
    public void setPrintHeadings(boolean value) {
        model.getPrintOptions().setHeadings(value);
    }

    /**
     * Returns the center horizontally.
     * @return the requested result
     */
    public boolean getCenterHorizontally() {
        return model.getPrintOptions().getHorizontalCentered();
    }

    /**
     * Sets the center horizontally.
     * @param value value to apply
     */
    public void setCenterHorizontally(boolean value) {
        model.getPrintOptions().setHorizontalCentered(value);
    }

    /**
     * Returns the center vertically.
     * @return the requested result
     */
    public boolean getCenterVertically() {
        return model.getPrintOptions().getVerticalCentered();
    }

    /**
     * Sets the center vertically.
     * @param value value to apply
     */
    public void setCenterVertically(boolean value) {
        model.getPrintOptions().setVerticalCentered(value);
    }

    /**
     * Returns the horizontal page breaks.
     * @return the requested result
     */
    public List<Integer> getHorizontalPageBreaks() {
        return getOrderedBreaks(model.getHorizontalPageBreaks());
    }

    /**
     * Returns the vertical page breaks.
     * @return the requested result
     */
    public List<Integer> getVerticalPageBreaks() {
        return getOrderedBreaks(model.getVerticalPageBreaks());
    }

    /**
     * Adds horizontal page break.
     * @param rowIndex zero-based row index
     */
    public void addHorizontalPageBreak(int rowIndex) {
        // Handle the relevant branch before the state changes.
        if (rowIndex < 0) {
            throw new CellsException("Horizontal page break row index must be non-negative.");
        }
        addDistinct(model.getHorizontalPageBreaks(), rowIndex);
    }

    /**
     * Adds vertical page break.
     * @param columnIndex zero-based column index
     */
    public void addVerticalPageBreak(int columnIndex) {
        // Handle the relevant branch before the state changes.
        if (columnIndex < 0) {
            throw new CellsException("Vertical page break column index must be non-negative.");
        }
        addDistinct(model.getVerticalPageBreaks(), columnIndex);
    }

    /**
     * Clears the current state maintained by this object.
     */
    public void clearHorizontalPageBreaks() {
        model.getHorizontalPageBreaks().clear();
    }

    /**
     * Clears the current state maintained by this object.
     */
    public void clearVerticalPageBreaks() {
        model.getVerticalPageBreaks().clear();
    }

    /**
     * Processes get ordered breaks.
     * @param breaks breaks
     * @return the requested result
     */
    private static List<Integer> getOrderedBreaks(Collection<Integer> breaks) {
        List<Integer> orderedBreaks = new ArrayList<>(breaks);
        orderedBreaks.sort(null);
        return orderedBreaks;
    }

    /**
     * Adds distinct.
     * @param collection collection
     * @param value value to apply
     */
    private static void addDistinct(Collection<Integer> collection, int value) {
        // Handle the relevant branch before the state changes.
        if (!collection.contains(value)) {
            collection.add(value);
        }
    }

    /**
     * Validates margin.
     * @param value value to apply
     * @param propertyName name to use
     * @return the computed result
     */
    private static double validateMargin(double value, String propertyName) {
        // Handle the relevant branch before the state changes.
        if (value < 0.0) {
            throw new CellsException(propertyName + " must be zero or greater.");
        }
        return value;
    }

    /**
     * Processes to centimeters.
     * @param inches inches
     * @return the computed result
     */
    private static double toCentimeters(double inches) {
        return inches * CENTIMETERS_PER_INCH;
    }

    /**
     * Processes to inches.
     * @param centimeters centimeters
     * @return the computed result
     */
    private static double toInches(double centimeters) {
        return centimeters / CENTIMETERS_PER_INCH;
    }

    /**
     * Normalizes the text.
     * @param value value to apply
     * @return the computed result
     */
    private static String normalizeText(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isBlank()) {
            return null;
        }
        return value.trim();
    }
}
