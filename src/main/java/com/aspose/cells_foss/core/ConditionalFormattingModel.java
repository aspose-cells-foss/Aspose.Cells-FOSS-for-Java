package com.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.List;
import com.aspose.cells_foss.CellArea;

/**
 * Represents the conditional formatting model for a worksheet.
 */
public final class ConditionalFormattingModel {
    private final List<CellArea> areas = new ArrayList<>();
    private final List<FormatConditionModel> conditions = new ArrayList<>();

    /**
     * Initializes a new ConditionalFormattingModel instance.
     */
    public ConditionalFormattingModel() {
        // Constructor initializes empty lists (C# object initializers become field initializers in Java)
    }

    /**
     * Returns the areas.
     * @return the requested result
     */
    public List<CellArea> getAreas() {
        return areas;
    }

    /**
     * Returns the conditions.
     * @return the requested result
     */
    public List<FormatConditionModel> getConditions() {
        return conditions;
    }
}