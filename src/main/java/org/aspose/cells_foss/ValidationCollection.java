package org.aspose.cells_foss;

import org.aspose.cells_foss.core.ValidationModel;
import java.util.List;

/**
 * Represents the collection of data validation rules for a worksheet.
 */
public final class ValidationCollection {
    private final List<ValidationModel> models;

    /**
     * Initializes a new ValidationCollection instance.
     * @param models models
     */
    ValidationCollection(List<ValidationModel> models) {
        this.models = models;
    }

    /**
     * Returns the count.
     * @return the requested result
     */
    public int getCount() {
        return models.size();
    }

    /**
     * Returns the requested item.
     * @param index index
     * @return the requested result
     */
    public Validation get(int index) {
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= models.size()) {
            throw new CellsException("Validation index was out of range.");
        }
        return new Validation(models, models.get(index));
    }

    /**
     * Adds a new item to the current collection.
     * @param area area
     * @return the computed result
     */
    public int add(CellArea area) {
        if (area.getFirstRow() < 0 || area.getFirstColumn() < 0
                || area.getTotalRows() <= 0 || area.getTotalColumns() <= 0) {
            throw new CellsException("Validation area must be a positive cell range.");
        }
        for (ValidationModel existing : models) {
            for (CellArea ea : existing.getAreas()) {
                if (areasOverlap(area, ea))
                    throw new CellsException("Validation area overlaps with an existing validation rule.");
            }
        }
        ValidationModel m = new ValidationModel();
        m.getAreas().add(area);
        models.add(m);
        return models.size() - 1;
    }

    private static boolean areasOverlap(CellArea a, CellArea b) {
        int aR2 = a.getFirstRow() + a.getTotalRows() - 1;
        int aC2 = a.getFirstColumn() + a.getTotalColumns() - 1;
        int bR2 = b.getFirstRow() + b.getTotalRows() - 1;
        int bC2 = b.getFirstColumn() + b.getTotalColumns() - 1;
        return a.getFirstRow() <= bR2 && aR2 >= b.getFirstRow()
            && a.getFirstColumn() <= bC2 && aC2 >= b.getFirstColumn();
    }

    /**
     * Processes get validation in cell.
     * @param row row
     * @param column column
     * @return the requested result
     */
    public Validation getValidationInCell(int row, int column) {
        // Handle the relevant branch before the state changes.
        if (row < 0 || column < 0) {
            throw new CellsException("Row and column indices must be non-negative.");
        }
        for (ValidationModel m : models) {
            for (CellArea area : m.getAreas()) {
                if (row >= area.getFirstRow()
                        && row < area.getFirstRow() + area.getTotalRows()
                        && column >= area.getFirstColumn()
                        && column < area.getFirstColumn() + area.getTotalColumns()) {
                    return new Validation(models, m);
                }
            }
        }
        return null;
    }

    /**
     * Removes a cell.
     * @param row row
     * @param column column
     */
    public void removeACell(int row, int column) {
        if (row < 0 || column < 0) {
            throw new CellsException("Row and column indices must be non-negative.");
        }
        models.removeIf(m -> {
            List<CellArea> toRemove = new java.util.ArrayList<>();
            List<CellArea> toAdd = new java.util.ArrayList<>();
            for (CellArea area : m.getAreas()) {
                int r1 = area.getFirstRow(), c1 = area.getFirstColumn();
                int r2 = r1 + area.getTotalRows() - 1, c2 = c1 + area.getTotalColumns() - 1;
                if (row >= r1 && row <= r2 && column >= c1 && column <= c2) {
                    toRemove.add(area);
                    // top strip: rows r1..row-1
                    if (row > r1) toAdd.add(new CellArea(r1, c1, row - r1, c2 - c1 + 1));
                    // bottom strip: rows row+1..r2
                    if (row < r2) toAdd.add(new CellArea(row + 1, c1, r2 - row, c2 - c1 + 1));
                    // left cell in same row: cols c1..column-1
                    if (column > c1) toAdd.add(new CellArea(row, c1, 1, column - c1));
                    // right cell in same row: cols column+1..c2
                    if (column < c2) toAdd.add(new CellArea(row, column + 1, 1, c2 - column));
                }
            }
            m.getAreas().removeAll(toRemove);
            m.getAreas().addAll(toAdd);
            return m.getAreas().isEmpty();
        });
    }

    /**
     * Removes area.
     * @param cellArea cell area
     */
    public void removeArea(CellArea cellArea) {
        models.removeIf(m -> {
            m.getAreas().remove(cellArea);
            return m.getAreas().isEmpty();
        });
    }
}

