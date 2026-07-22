package org.aspose.cells_foss;

import org.aspose.cells_foss.core.WorkbookPropertiesModel;

/** Represents workbook calculation settings. */
public final class CalculationProperties {
    private final WorkbookPropertiesModel.CalculationPropertiesModel model;

    CalculationProperties(WorkbookPropertiesModel.CalculationPropertiesModel model) {
        this.model = model;
    }

    public Integer getCalculationId() { return model.getCalculationId(); }
    public void setCalculationId(Integer v) {
        if (v != null && v < 0) throw new CellsException("CalculationId must be non-negative.");
        model.setCalculationId(v);
    }

    public String getCalculationMode() {
        String m = model.getCalculationMode();
        return m.isEmpty() ? "auto" : m;
    }
    public void setCalculationMode(String v) { model.setCalculationMode(v != null ? v : ""); }

    public boolean getFullCalculationOnLoad() { return model.getFullCalculationOnLoad(); }
    public void setFullCalculationOnLoad(boolean v) { model.setFullCalculationOnLoad(v); }

    public String getReferenceMode() {
        String m = model.getReferenceMode();
        return m.isEmpty() ? "A1" : m;
    }
    public void setReferenceMode(String v) { model.setReferenceMode(v != null ? v : ""); }

    public boolean getIterate() { return model.getIterate(); }
    public void setIterate(boolean v) { model.setIterate(v); }

    public int getIterateCount() {
        Integer v = model.getIterateCount();
        return v != null ? v : 100;
    }
    public void setIterateCount(int v) {
        if (v < 0) throw new CellsException("IterateCount must be non-negative.");
        model.setIterateCount(v);
    }

    public double getIterateDelta() {
        Double v = model.getIterateDelta();
        return v != null ? v : 0.001;
    }
    public void setIterateDelta(double v) {
        if (v < 0.0) throw new CellsException("IterateDelta must be non-negative.");
        model.setIterateDelta(v);
    }

    public boolean getFullPrecision() {
        Boolean v = model.getFullPrecision();
        return v == null || v;
    }
    public void setFullPrecision(boolean v) { model.setFullPrecision(v); }

    public boolean getCalculationCompleted() {
        Boolean v = model.getCalculationCompleted();
        return v == null || v;
    }
    public void setCalculationCompleted(boolean v) { model.setCalculationCompleted(v); }

    public boolean getCalculationOnSave() {
        Boolean v = model.getCalculationOnSave();
        return v == null || v;
    }
    public void setCalculationOnSave(boolean v) { model.setCalculationOnSave(v); }

    public boolean getConcurrentCalculation() {
        Boolean v = model.getConcurrentCalculation();
        return v == null || v;
    }
    public void setConcurrentCalculation(boolean v) { model.setConcurrentCalculation(v); }

    public boolean getForceFullCalculation() { return model.getForceFullCalculation(); }
    public void setForceFullCalculation(boolean v) { model.setForceFullCalculation(v); }
}

