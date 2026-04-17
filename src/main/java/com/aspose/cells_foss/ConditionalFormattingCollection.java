package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ConditionalFormattingModel;
import java.util.List;

/**
 * A collection of conditional formatting objects.
 */
public final class ConditionalFormattingCollection {

    private final List<ConditionalFormattingModel> collections;

    /**
     * Initializes a new ConditionalFormattingCollection instance.
     * @param collections collections
     */
    ConditionalFormattingCollection(List<ConditionalFormattingModel> collections) {
        this.collections = collections;
    }

    /**
     * Returns the count.
     * @return the requested result
     */
    public int getCount() {
        return collections.size();
    }

    /**
     * Returns the requested item.
     * @param index index
     * @return the requested result
     */
    public FormatConditionCollection get(int index) {
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= collections.size()) {
            throw new CellsException("Conditional formatting index was out of range.");
        }

        return new FormatConditionCollection(collections, collections.get(index));
    }

    /**
     * Adds a new item to the current collection.
     * @return the computed result
     */
    public int add() {
        collections.add(new ConditionalFormattingModel());
        return collections.size() - 1;
    }

    /**
     * Removes at.
     * @param index index
     */
    public void removeAt(int index) {
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= collections.size()) {
            throw new CellsException("Conditional formatting index was out of range.");
        }

        collections.remove(index);
    }

    /**
     * Removes area.
     * @param startRow start row
     * @param startColumn start column
     * @param totalRows total rows
     * @param totalColumns total columns
     */
    public void removeArea(int startRow, int startColumn, int totalRows, int totalColumns) {
        CellArea area = new CellArea(startRow, startColumn, totalRows, totalColumns);
        // Walk the current collection so every entry is processed consistently.
        for (int index = collections.size() - 1; index >= 0; index--) {
            FormatConditionCollection collection = new FormatConditionCollection(collections, collections.get(index));
            collection.removeArea(area);
        }
    }

    /**
     * Processes get next priority.
     * @param collections collections
     * @return the requested result
     */
    static int getNextPriority(List<ConditionalFormattingModel> collections) {
        int maxPriority = 0;
        // Walk the current collection so every entry is processed consistently.
        for (ConditionalFormattingModel collection : collections) {
            for (int conditionIndex = 0; conditionIndex < collection.getConditions().size(); conditionIndex++) {
                maxPriority = Math.max(maxPriority, collection.getConditions().get(conditionIndex).getPriority());
            }
        }

        return maxPriority + 1;
    }
}