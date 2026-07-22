package org.aspose.cells_foss;

import org.aspose.cells_foss.core.CellAddress;
import org.aspose.cells_foss.core.HyperlinkModel;
import java.util.List;
import java.util.Objects;

/**
 * Represents a collection of hyperlinks in a worksheet.
 */
public final class HyperlinkCollection {
    private final List<HyperlinkModel> hyperlinks;

    /**
     * Constructs a new HyperlinkCollection.
     *
     * @param hyperlinks the internal list of hyperlink models
     */
    HyperlinkCollection(List<HyperlinkModel> hyperlinks) {
        this.hyperlinks = hyperlinks;
    }

    /**
     * Gets the number of hyperlinks in the collection.
     *
     * @return the count of hyperlinks
     */
    public int getCount() {
        return hyperlinks.size();
    }

    /**
     * Gets the hyperlink at the specified index.
     *
     * @param index the zero-based index of the hyperlink to get
     * @return the hyperlink at the specified index
     * @throws CellsException if the index is out of range
     */
    public Hyperlink get(int index) {
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= hyperlinks.size()) {
            throw new CellsException("Hyperlink index was out of range.");
        }
        return new Hyperlink(hyperlinks, hyperlinks.get(index));
    }

    /**
     * Adds a hyperlink anchored at a single cell.
     *
     * @param cellName     the cell name (e.g. "A1")
     * @param totalRows    the number of rows the hyperlink spans
     * @param totalColumns the number of columns the hyperlink spans
     * @param address      the hyperlink address
     * @return the zero-based index of the added hyperlink
     */
    public int add(String cellName, int totalRows, int totalColumns, String address) {
        Objects.requireNonNull(address, "address");
        return addInternal(cellName, totalRows, totalColumns, address, null, null);
    }

    /**
     * Adds a hyperlink with explicit row/column coordinates.
     *
     * @param firstRow     the first row index
     * @param firstColumn  the first column index
     * @param totalRows    the number of rows the hyperlink spans
     * @param totalColumns the number of columns the hyperlink spans
     * @param address      the hyperlink address
     * @return the zero-based index of the added hyperlink
     */
    public int add(int firstRow, int firstColumn, int totalRows, int totalColumns, String address) {
        // Handle the relevant branch before the state changes.
        if (firstRow < 0 || firstColumn < 0) {
            throw new CellsException("Hyperlink origin indices must be non-negative.");
        }
        Objects.requireNonNull(address, "address");
        return addInternal(null, totalRows, totalColumns, address, firstRow, firstColumn);
    }

    /**
     * Adds a hyperlink between two cells with optional display text and tooltip.
     *
     * @param startCellName  the start cell name (e.g. "A1")
     * @param endCellName    the end cell name (e.g. "B2")
     * @param address        the hyperlink address
     * @param textToDisplay  the text to display for the hyperlink
     * @param screenTip      the tooltip text
     * @return the zero-based index of the added hyperlink
     */
    public int add(String startCellName, String endCellName, String address, String textToDisplay, String screenTip) {
        Objects.requireNonNull(startCellName, "startCellName");
        Objects.requireNonNull(endCellName, "endCellName");
        Objects.requireNonNull(address, "address");

        CellAddress startAddress;
        CellAddress endAddress;
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            startAddress = CellAddress.parse(startCellName);
            endAddress = CellAddress.parse(endCellName);
        } catch (IllegalArgumentException exception) {
            throw new CellsException(exception.getMessage(), exception);
        }

        int firstRow = Math.min(startAddress.getRowIndex(), endAddress.getRowIndex());
        int firstColumn = Math.min(startAddress.getColumnIndex(), endAddress.getColumnIndex());
        int lastRow = Math.max(startAddress.getRowIndex(), endAddress.getRowIndex());
        int lastColumn = Math.max(startAddress.getColumnIndex(), endAddress.getColumnIndex());

        int index = addInternal(null, lastRow - firstRow + 1, lastColumn - firstColumn + 1, address, firstRow, firstColumn);
        HyperlinkModel hyperlink = hyperlinks.get(index);
        hyperlink.setTextToDisplay(normalizeText(textToDisplay));
        hyperlink.setScreenTip(normalizeText(screenTip));
        return index;
    }

    /**
     * Removes the hyperlink at the specified index.
     *
     * @param index the zero-based index of the hyperlink to remove
     * @throws CellsException if the index is out of range
     */
    public void removeAt(int index) {
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= hyperlinks.size()) {
            throw new CellsException("Hyperlink index was out of range.");
        }
        hyperlinks.remove(index);
    }

    /**
     * Adds internal.
     * @param cellName name to use
     * @param totalRows total rows
     * @param totalColumns total columns
     * @param address address
     * @param firstRowOverride first row override
     * @param firstColumnOverride first column override
     * @return the computed result
     */
    private int addInternal(String cellName, int totalRows, int totalColumns, String address, Integer firstRowOverride, Integer firstColumnOverride) {
        // Handle the relevant branch before the state changes.
        if (totalRows <= 0 || totalColumns <= 0) {
            throw new CellsException("Hyperlink range dimensions must be positive.");
        }

        CellAddress anchor;
        if (firstRowOverride != null && firstColumnOverride != null) {
            anchor = new CellAddress(firstRowOverride, firstColumnOverride);
        } else {
            if (cellName == null || cellName.isBlank()) {
                throw new CellsException("Hyperlink anchor must be a valid cell reference.");
            }
            try {
                anchor = CellAddress.parse(cellName);
            } catch (IllegalArgumentException exception) {
                throw new CellsException(exception.getMessage(), exception);
            }
        }

        HyperlinkModel candidate = new HyperlinkModel();
        candidate.setFirstRow(anchor.getRowIndex());
        candidate.setFirstColumn(anchor.getColumnIndex());
        candidate.setTotalRows(totalRows);
        candidate.setTotalColumns(totalColumns);
        assignAddress(candidate, address);

        for (int index = 0; index < hyperlinks.size(); index++) {
            if (overlaps(hyperlinks.get(index), candidate)) {
                throw new CellsException("Hyperlink ranges must not overlap.");
            }
        }

        hyperlinks.add(candidate);
        return hyperlinks.size() - 1;
    }

    /**
     * Assigns address.
     * @param model model
     * @param address address
     */
    private static void assignAddress(HyperlinkModel model, String address) {
        // Handle the relevant branch before the state changes.
        if (address == null || address.isBlank()) {
            throw new CellsException("Hyperlink address must be non-empty.");
        }

        String normalized = address.trim();
        if (normalized.startsWith("#")) {
            model.setAddress(null);
            model.setSubAddress(normalized.substring(1));
            return;
        }

        if (normalized.indexOf('!') >= 0) {
            model.setAddress(null);
            model.setSubAddress(normalized);
            return;
        }

        model.setAddress(normalized);
        model.setSubAddress(null);
    }

    /**
     * Processes overlaps.
     * @param left left
     * @param right right
     * @return true when the condition is satisfied
     */
    private static boolean overlaps(HyperlinkModel left, HyperlinkModel right) {
        int leftLastRow = left.getFirstRow() + left.getTotalRows() - 1;
        int leftLastColumn = left.getFirstColumn() + left.getTotalColumns() - 1;
        int rightLastRow = right.getFirstRow() + right.getTotalRows() - 1;
        int rightLastColumn = right.getFirstColumn() + right.getTotalColumns() - 1;

        return left.getFirstRow() <= rightLastRow
                && right.getFirstRow() <= leftLastRow
                && left.getFirstColumn() <= rightLastColumn
                && right.getFirstColumn() <= leftLastColumn;
    }

    /**
     * Normalizes the text.
     * @param value value to apply
     * @return the computed result
     */
    private static String normalizeText(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isEmpty()) {
            return null;
        }
        return value;
    }
}
