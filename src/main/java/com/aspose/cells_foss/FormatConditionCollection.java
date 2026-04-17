package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ConditionalFormattingModel;
import com.aspose.cells_foss.core.FormatConditionModel;

import java.util.ArrayList;
import java.util.List;

/**
 * Represents a collection of format conditions in Excel.
 */
public final class FormatConditionCollection {

    private final List<ConditionalFormattingModel> owner;
    private final ConditionalFormattingModel model;

    /**
     * Initializes a new FormatConditionCollection instance.
     * @param owner owner
     * @param model model
     */
    FormatConditionCollection(List<ConditionalFormattingModel> owner, ConditionalFormattingModel model) {
        this.owner = owner;
        this.model = model;
    }

    /**
     * Returns the count.
     * @return the requested result
     */
    public int getCount() {
        return model.getConditions().size();
    }

    /**
     * Returns the range count.
     * @return the requested result
     */
    public int getRangeCount() {
        return model.getAreas().size();
    }

    /**
     * Returns the requested item.
     * @param index index
     * @return the requested result
     */
    public FormatCondition get(int index) {
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= model.getConditions().size()) {
            throw new CellsException("Format condition index was out of range.");
        }
        return new FormatCondition(owner, model, model.getConditions().get(index));
    }

    /**
     * Adds a new item to the current collection.
     * @param area area
     * @param type type
     * @param operatorType operator type
     * @param formula1 formula 1
     * @param formula2 formula 2
     * @return the computed result
     */
    public int add(CellArea area, FormatConditionType type, OperatorType operatorType, String formula1, String formula2) {
        addArea(area);
        return addCondition(type, operatorType, formula1, formula2);
    }

    /**
     * Adds condition.
     * @param type type
     * @return the computed result
     */
    public int addCondition(FormatConditionType type) {
        return addCondition(type, OperatorType.NONE, "", "");
    }

    /**
     * Adds condition.
     * @param type type
     * @param operatorType operator type
     * @param formula1 formula 1
     * @param formula2 formula 2
     * @return the computed result
     */
    public int addCondition(FormatConditionType type, OperatorType operatorType, String formula1, String formula2) {
        FormatConditionModel condition = new FormatConditionModel();
        condition.setType(type);
        condition.setOperator(operatorType);
        condition.setFormula1(normalizeFormula(formula1));
        condition.setFormula2(normalizeFormula(formula2));
        condition.setPriority(ConditionalFormattingCollection.getNextPriority(owner));
        condition.setStyle(com.aspose.cells_foss.core.StyleValue.getDefault().clone());
        initializeDefaults(condition);
        model.getConditions().add(condition);
        return model.getConditions().size() - 1;
    }

    /**
     * Adds area.
     * @param area area
     */
    public void addArea(CellArea area) {
        validateArea(area);
        model.getAreas().add(area);
        sortAreas(model.getAreas());
    }

    /**
     * Processes get cell area.
     * @param index index
     * @return the requested result
     */
    public CellArea getCellArea(int index) {
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= model.getAreas().size()) {
            throw new CellsException("Conditional formatting area index was out of range.");
        }

        return model.getAreas().get(index);
    }

    /**
     * Removes area.
     * @param index index
     */
    public void removeArea(int index) {
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= model.getAreas().size()) {
            throw new CellsException("Conditional formatting area index was out of range.");
        }

        model.getAreas().remove(index);
        removeCollectionIfEmpty(owner, model);
    }

    /**
     * Removes area.
     * @param startRow start row
     * @param startColumn start column
     * @param totalRows total rows
     * @param totalColumns total columns
     */
    public void removeArea(int startRow, int startColumn, int totalRows, int totalColumns) {
        removeArea(new CellArea(startRow, startColumn, totalRows, totalColumns));
    }

    /**
     * Removes area.
     * @param area area
     */
    void removeArea(CellArea area) {
        validateArea(area);
        replaceAreas(model, subtractAreas(model.getAreas(), area));
        removeCollectionIfEmpty(owner, model);
    }

    /**
     * Removes condition.
     * @param index index
     */
    public void removeCondition(int index) {
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= model.getConditions().size()) {
            throw new CellsException("Format condition index was out of range.");
        }

        model.getConditions().remove(index);
        removeCollectionIfEmpty(owner, model);
    }

    /**
     * Removes condition.
     * @param owner owner
     * @param collection collection
     * @param model model
     */
    static void removeCondition(List<ConditionalFormattingModel> owner, ConditionalFormattingModel collection, FormatConditionModel model) {
        collection.getConditions().remove(model);
        removeCollectionIfEmpty(owner, collection);
    }

    /**
     * Initializes defaults.
     * @param condition condition
     */
    private static void initializeDefaults(FormatConditionModel condition) {
        // Translate the internal value into the matching public representation.
        switch (condition.getType()) {
            case DUPLICATE_VALUES:
                condition.setDuplicate(true);
                break;
            case UNIQUE_VALUES:
                condition.setDuplicate(false);
                break;
            case TOP_10:
                condition.setTop(true);
                condition.setRank(10);
                break;
            case BOTTOM_10:
                condition.setTop(false);
                condition.setRank(10);
                break;
            case ABOVE_AVERAGE:
                condition.setAbove(true);
                break;
            case BELOW_AVERAGE:
                condition.setAbove(false);
                break;
            case COLOR_SCALE:
                condition.setColorScaleCount(2);
                break;
            case DATA_BAR:
                condition.setBarColor(new com.aspose.cells_foss.core.ColorValue((byte) 255, (byte) 99, (byte) 142, (byte) 198));
                break;
            case ICON_SET:
                condition.setIconSetType("3TrafficLights1");
                break;
        }
    }

    /**
     * Removes collection if empty.
     * @param owner owner
     * @param collection collection
     */
    private static void removeCollectionIfEmpty(List<ConditionalFormattingModel> owner, ConditionalFormattingModel collection) {
        // Handle the relevant branch before the state changes.
        if (collection.getAreas().size() == 0 || collection.getConditions().size() == 0) {
            owner.remove(collection);
        }
    }

    /**
     * Replaces areas.
     * @param model model
     * @param areas areas
     */
    private static void replaceAreas(ConditionalFormattingModel model, List<CellArea> areas) {
        model.getAreas().clear();
        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < areas.size(); index++) {
            model.getAreas().add(areas.get(index));
        }
    }

    /**
     * Normalizes the formula.
     * @param value value to apply
     * @return the computed result
     */
    private static String normalizeFormula(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            return null;
        }

        String trimmed = value.trim();
        if (trimmed.length() == 0) {
            return null;
        }

        if (trimmed.charAt(0) == '=') {
            return trimmed.substring(1);
        }

        return trimmed;
    }

    /**
     * Validates area.
     * @param area area
     */
    private static void validateArea(CellArea area) {
        // Handle the relevant branch before the state changes.
        if (area.getFirstRow() < 0 || area.getFirstColumn() < 0 || area.getTotalRows() <= 0 || area.getTotalColumns() <= 0) {
            throw new CellsException("Conditional formatting area must be a positive cell range.");
        }
    }

    /**
     * Subtracts areas.
     * @param sourceAreas source areas
     * @param removal removal
     * @return the computed result
     */
    private static List<CellArea> subtractAreas(List<CellArea> sourceAreas, CellArea removal) {
        List<CellArea> remaining = new ArrayList<>();
        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < sourceAreas.size(); index++) {
            subtractArea(sourceAreas.get(index), removal, remaining);
        }

        sortAreas(remaining);
        return remaining;
    }

    /**
     * Subtracts area.
     * @param source source
     * @param removal removal
     * @param output output
     */
    private static void subtractArea(CellArea source, CellArea removal, List<CellArea> output) {
        // Handle the relevant branch before the state changes.
        if (!areasOverlap(source, removal)) {
            output.add(source);
            return;
        }

        int sourceLastRow = source.getFirstRow() + source.getTotalRows() - 1;
        int sourceLastColumn = source.getFirstColumn() + source.getTotalColumns() - 1;
        int removalLastRow = removal.getFirstRow() + removal.getTotalRows() - 1;
        int removalLastColumn = removal.getFirstColumn() + removal.getTotalColumns() - 1;

        int overlapFirstRow = Math.max(source.getFirstRow(), removal.getFirstRow());
        int overlapFirstColumn = Math.max(source.getFirstColumn(), removal.getFirstColumn());
        int overlapLastRow = Math.min(sourceLastRow, removalLastRow);
        int overlapLastColumn = Math.min(sourceLastColumn, removalLastColumn);

        addIfNonEmpty(output, source.getFirstRow(), source.getFirstColumn(), overlapFirstRow - 1, sourceLastColumn);
        addIfNonEmpty(output, overlapLastRow + 1, source.getFirstColumn(), sourceLastRow, sourceLastColumn);
        addIfNonEmpty(output, overlapFirstRow, source.getFirstColumn(), overlapLastRow, overlapFirstColumn - 1);
        addIfNonEmpty(output, overlapFirstRow, overlapLastColumn + 1, overlapLastRow, sourceLastColumn);
    }

    /**
     * Adds if non empty.
     * @param areas areas
     * @param firstRow first row
     * @param firstColumn first column
     * @param lastRow last row
     * @param lastColumn last column
     */
    private static void addIfNonEmpty(List<CellArea> areas, int firstRow, int firstColumn, int lastRow, int lastColumn) {
        // Handle the relevant branch before the state changes.
        if (lastRow < firstRow || lastColumn < firstColumn) {
            return;
        }

        areas.add(CellArea.createCellArea(firstRow, firstColumn, lastRow, lastColumn));
    }

    /**
     * Processes areas overlap.
     * @param left left
     * @param right right
     * @return true when the condition is satisfied
     */
    static boolean areasOverlap(CellArea left, CellArea right) {
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
     * Sorts areas.
     * @param areas areas
     */
    static void sortAreas(List<CellArea> areas) {
        areas.sort(FormatConditionCollection::compareAreas);
    }

    /**
     * Compares areas.
     * @param left left
     * @param right right
     * @return the computed result
     */
    static int compareAreas(CellArea left, CellArea right) {
        int rowComparison = Integer.compare(left.getFirstRow(), right.getFirstRow());
        // Handle the relevant branch before the state changes.
        if (rowComparison != 0) {
            return rowComparison;
        }

        int columnComparison = Integer.compare(left.getFirstColumn(), right.getFirstColumn());
        if (columnComparison != 0) {
            return columnComparison;
        }

        int rowCountComparison = Integer.compare(left.getTotalRows(), right.getTotalRows());
        if (rowCountComparison != 0) {
            return rowCountComparison;
        }

        return Integer.compare(left.getTotalColumns(), right.getTotalColumns());
    }
}