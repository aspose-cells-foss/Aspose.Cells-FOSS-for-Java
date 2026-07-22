package org.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.List;

/**
 * Represents page setup model for an Excel worksheet.
 */
public final class PageSetupModel {
    private final PageMarginsModel margins = new PageMarginsModel();
    private final PrintOptionsModel printOptions = new PrintOptionsModel();
    private final HeaderFooterModel headerFooter = new HeaderFooterModel();
    private int paperSize;
    private PageOrientation orientation;
    private Integer firstPageNumber;
    private Integer scale;
    private Integer fitToWidth;
    private Integer fitToHeight;
    private String printArea;
    private String printTitleRows;
    private String printTitleColumns;
    private final List<Integer> horizontalPageBreaks = new ArrayList<>();
    private final List<Integer> verticalPageBreaks = new ArrayList<>();

    /**
     * Returns the margins.
     * @return the requested result
     */
    public PageMarginsModel getMargins() { return margins; }

    /**
     * Returns the print options.
     * @return the requested result
     */
    public PrintOptionsModel getPrintOptions() { return printOptions; }

    /**
     * Returns the header footer.
     * @return the requested result
     */
    public HeaderFooterModel getHeaderFooter() { return headerFooter; }

    /**
     * Returns the paper size.
     * @return the requested result
     */
    public int getPaperSize() { return paperSize; }
    /**
     * Sets the paper size.
     * @param paperSize paper size
     */
    public void setPaperSize(int paperSize) { this.paperSize = paperSize; }

    /**
     * Returns the orientation.
     * @return the requested result
     */
    public PageOrientation getOrientation() { return orientation; }
    /**
     * Sets the orientation.
     * @param orientation orientation
     */
    public void setOrientation(PageOrientation orientation) { this.orientation = orientation; }

    /**
     * Returns the first page number.
     * @return the requested result
     */
    public Integer getFirstPageNumber() { return firstPageNumber; }
    /**
     * Sets the first page number.
     * @param firstPageNumber first page number
     */
    public void setFirstPageNumber(Integer firstPageNumber) { this.firstPageNumber = firstPageNumber; }

    /**
     * Returns the scale.
     * @return the requested result
     */
    public Integer getScale() { return scale; }
    /**
     * Sets the scale.
     * @param scale scale
     */
    public void setScale(Integer scale) { this.scale = scale; }

    /**
     * Returns the fit to width.
     * @return the requested result
     */
    public Integer getFitToWidth() { return fitToWidth; }
    /**
     * Sets the fit to width.
     * @param fitToWidth fit to width
     */
    public void setFitToWidth(Integer fitToWidth) { this.fitToWidth = fitToWidth; }

    /**
     * Returns the fit to height.
     * @return the requested result
     */
    public Integer getFitToHeight() { return fitToHeight; }
    /**
     * Sets the fit to height.
     * @param fitToHeight fit to height
     */
    public void setFitToHeight(Integer fitToHeight) { this.fitToHeight = fitToHeight; }

    /**
     * Returns the print area.
     * @return the requested result
     */
    public String getPrintArea() { return printArea; }
    /**
     * Sets the print area.
     * @param printArea print area
     */
    public void setPrintArea(String printArea) { this.printArea = printArea; }

    /**
     * Returns the print title rows.
     * @return the requested result
     */
    public String getPrintTitleRows() { return printTitleRows; }
    /**
     * Sets the print title rows.
     * @param printTitleRows print title rows
     */
    public void setPrintTitleRows(String printTitleRows) { this.printTitleRows = printTitleRows; }

    /**
     * Returns the print title columns.
     * @return the requested result
     */
    public String getPrintTitleColumns() { return printTitleColumns; }
    /**
     * Sets the print title columns.
     * @param printTitleColumns print title columns
     */
    public void setPrintTitleColumns(String printTitleColumns) { this.printTitleColumns = printTitleColumns; }

    /**
     * Returns the horizontal page breaks.
     * @return the requested result
     */
    public List<Integer> getHorizontalPageBreaks() { return horizontalPageBreaks; }

    /**
     * Returns the vertical page breaks.
     * @return the requested result
     */
    public List<Integer> getVerticalPageBreaks() { return verticalPageBreaks; }
}
