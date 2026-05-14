package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ChartModel;

/** Represents an embedded chart in a worksheet (read-only; charts are round-tripped verbatim). */
public final class Chart {
    private final ChartModel model;

    Chart(ChartModel model) {
        this.model = model;
    }

    ChartModel getModel() { return model; }

    public String getName() { return model.getName(); }

    public ChartType getChartType() { return ChartCollection.parseChartType(model.getChartType()); }

    public int getUpperLeftRow() { return model.getUpperLeftRow(); }
    public int getUpperLeftColumn() { return model.getUpperLeftColumn(); }
    public int getLowerRightRow() { return model.getLowerRightRow(); }
    public int getLowerRightColumn() { return model.getLowerRightColumn(); }

    /** Width in English Metric Units (EMUs). */
    public long getExtentCx() { return model.getExtentCx(); }

    /** Height in English Metric Units (EMUs). */
    public long getExtentCy() { return model.getExtentCy(); }
}
