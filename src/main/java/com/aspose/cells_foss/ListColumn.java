package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ListColumnModel;

/** Represents a column within an Excel table (ListObject). */
public final class ListColumn {
    private final ListColumnModel model;

    ListColumn(ListColumnModel model) {
        this.model = model;
    }

    ListColumnModel getModel() { return model; }

    /** One-based column identifier as stored in OOXML. */
    public int getId() { return model.getId(); }

    public String getName() { return model.getName(); }
    public void setName(String name) {
        if (name == null || name.isBlank())
            throw new CellsException("Column name must be non-empty.");
        model.setName(name.trim());
    }

    public TotalsCalculation getTotalsCalculation() {
        return parseTotalsCalculation(model.getTotalsRowFunction());
    }

    public void setTotalsCalculation(TotalsCalculation calc) {
        model.setTotalsRowFunction(totalsCalculationToString(calc));
    }

    public String getTotalsRowLabel() { return model.getTotalsRowLabel(); }
    public void setTotalsRowLabel(String label) { model.setTotalsRowLabel(label != null ? label : ""); }

    public String getTotalsRowFormula() { return model.getTotalsRowFormula(); }
    public void setTotalsRowFormula(String formula) { model.setTotalsRowFormula(formula != null ? formula : ""); }

    // -------------------------------------------------------------------------

    public static TotalsCalculation parseTotalsCalculation(String raw) {
        if (raw == null) return TotalsCalculation.NONE;
        return switch (raw.toLowerCase()) {
            case "sum"       -> TotalsCalculation.SUM;
            case "count"     -> TotalsCalculation.COUNT;
            case "average"   -> TotalsCalculation.AVERAGE;
            case "max"       -> TotalsCalculation.MAX;
            case "min"       -> TotalsCalculation.MIN;
            case "stddev"    -> TotalsCalculation.STD_DEV;
            case "var"       -> TotalsCalculation.VAR;
            case "countnums" -> TotalsCalculation.COUNT_NUMS;
            case "custom"    -> TotalsCalculation.CUSTOM;
            default          -> TotalsCalculation.NONE;
        };
    }

    public static String totalsCalculationToString(TotalsCalculation calc) {
        return switch (calc) {
            case SUM        -> "sum";
            case COUNT      -> "count";
            case AVERAGE    -> "average";
            case MAX        -> "max";
            case MIN        -> "min";
            case STD_DEV    -> "stdDev";
            case VAR        -> "var";
            case COUNT_NUMS -> "countNums";
            case CUSTOM     -> "custom";
            default         -> "none";
        };
    }
}
