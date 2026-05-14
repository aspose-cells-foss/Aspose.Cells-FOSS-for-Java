package com.aspose.cells_foss.core;

/**
 * Internal model for a table (ListObject) column.
 */
public final class ListColumnModel {
    private int id;
    private String name = "";
    private String totalsRowFunction = "none";
    private String totalsRowLabel = "";
    private String totalsRowFormula = "";

    public int getId() { return id; }
    public void setId(int id) { this.id = id; }

    public String getName() { return name; }
    public void setName(String name) { this.name = name != null ? name : ""; }

    public String getTotalsRowFunction() { return totalsRowFunction; }
    public void setTotalsRowFunction(String totalsRowFunction) { this.totalsRowFunction = totalsRowFunction != null ? totalsRowFunction : "none"; }

    public String getTotalsRowLabel() { return totalsRowLabel; }
    public void setTotalsRowLabel(String totalsRowLabel) { this.totalsRowLabel = totalsRowLabel != null ? totalsRowLabel : ""; }

    public String getTotalsRowFormula() { return totalsRowFormula; }
    public void setTotalsRowFormula(String totalsRowFormula) { this.totalsRowFormula = totalsRowFormula != null ? totalsRowFormula : ""; }
}
