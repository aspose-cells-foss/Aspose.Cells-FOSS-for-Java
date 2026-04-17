package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ValidationModel;
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
        // Handle the relevant branch before the state changes.
        if (area.getFirstRow() < 0 || area.getFirstColumn() < 0
                || area.getTotalRows() <= 0 || area.getTotalColumns() <= 0) {
            throw new CellsException("Validation area must be a positive cell range.");
        }
        ValidationModel m = new ValidationModel();
        m.getAreas().add(area);
        models.add(m);
        return models.size() - 1;
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
        // Handle the relevant branch before the state changes.
        if (row < 0 || column < 0) {
            throw new CellsException("Row and column indices must be non-negative.");
        }
        models.removeIf(m -> {
            m.getAreas().removeIf(area ->
                    row >= area.getFirstRow()
                    && row < area.getFirstRow() + area.getTotalRows()
                    && column >= area.getFirstColumn()
                    && column < area.getFirstColumn() + area.getTotalColumns());
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
