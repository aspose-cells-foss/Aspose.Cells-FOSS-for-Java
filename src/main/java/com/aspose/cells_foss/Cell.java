package com.aspose.cells_foss;

import com.aspose.cells_foss.core.CellAddress;
import com.aspose.cells_foss.core.CellRecord;
import com.aspose.cells_foss.core.CellValueKind;
import java.util.Locale;

/**
 * Represents a cell in a worksheet.
 */
public class Cell {
    private final Worksheet worksheet;
    private final CellAddress address;

    /**
     * Initializes a new Cell instance.
     * @param worksheet worksheet to apply
     * @param address address
     */
    Cell(Worksheet worksheet, CellAddress address) {
        this.worksheet = worksheet;
        this.address = address;
    }

    /**
     * Returns the value.
     * @return the requested result
     */
    public Object getValue() {
        CellRecord record = tryGetRecord();
        return record != null ? record.getValue() : null;
    }

    /**
     * Sets the value.
     * @param value value to apply
     */
    public void setValue(Object value) {
        assignValue(value);
    }

    /**
     * Returns the string value.
     * @return the requested result
     */
    public String getStringValue() {
        CellRecord record = tryGetRecord();
        return DisplayTextFormatter.formatStringValue(record != null ? record.getValue() : null);
    }

    /**
     * Returns the display string value.
     * @return the requested result
     */
    public String getDisplayStringValue() {
        CellRecord record = tryGetRecord();
        com.aspose.cells_foss.core.StyleValue style =
                record != null && record.getStyle() != null ? record.getStyle() : com.aspose.cells_foss.core.StyleValue.getDefault();
        return DisplayTextFormatter.formatDisplayValue(
                record != null ? record.getValue() : null,
                style,
                Locale.US,
                worksheet.getWorkbook().getModel().getSettings().getDateSystem());
    }

    /**
     * Returns the formula.
     * @return the requested result
     */
    public String getFormula() {
        CellRecord record = tryGetRecord();
        // Handle the relevant branch before the state changes.
        if (record == null || record.getFormula() == null || record.getFormula().isBlank()) {
            return "";
        }
        return "=" + record.getFormula();
    }

    /**
     * Sets the formula.
     * @param value value to apply
     */
    public void setFormula(String value) {
        CellRecord record = getOrCreateRecord();
        record.setFormula(normalizeFormula(value));
        // Handle the relevant branch before the state changes.
        if (record.getFormula() == null || record.getFormula().isBlank()) {
            if (record.getValue() == null) {
                record.setKind(CellValueKind.BLANK);
            }
            return;
        }
        record.setKind(CellValueKind.FORMULA);
    }

    /**
     * Returns the type.
     * @return the requested result
     */
    public CellValueType getType() {
        CellRecord record = tryGetRecord();
        // Handle the relevant branch before the state changes.
        if (record == null) {
            return CellValueType.BLANK;
        }
        if (record.getFormula() != null && !record.getFormula().isBlank()) {
            return CellValueType.FORMULA;
        }
        switch (record.getKind()) {
            case STRING:
                return CellValueType.STRING;
            case NUMBER:
                return CellValueType.NUMBER;
            case BOOLEAN:
                return CellValueType.BOOLEAN;
            case DATE_TIME:
                return CellValueType.DATE_TIME;
            case FORMULA:
                return CellValueType.FORMULA;
            default:
                return CellValueType.BLANK;
        }
    }

    /**
     * Stores the value.
     * @param value value to apply
     */
    public void putValue(String value) {
        java.util.Objects.requireNonNull(value, "value");
        setScalar(value, CellValueKind.STRING);
    }

    /**
     * Stores the value.
     * @param value value to apply
     */
    public void putValue(int value) {
        setScalar(value, CellValueKind.NUMBER);
    }

    /**
     * Stores the value.
     * @param value value to apply
     */
    public void putValue(double value) {
        setScalar(value, CellValueKind.NUMBER);
    }

    /**
     * Stores the value.
     * @param value value to apply
     */
    public void putValue(boolean value) {
        setScalar(value, CellValueKind.BOOLEAN);
    }

    /**
     * Stores the value.
     * @param value value to apply
     */
    public void putValue(java.time.LocalDateTime value) {
        setScalar(value, CellValueKind.DATE_TIME);
    }

    /**
     * Returns the style.
     * @return the requested result
     */
    public Style getStyle() {
        CellRecord record = tryGetRecord();
        com.aspose.cells_foss.core.StyleValue sv =
                record != null ? record.getStyle() : com.aspose.cells_foss.core.StyleValue.getDefault();
        return Style.fromModel(sv);
    }

    /**
     * Sets the style.
     * @param style style to apply
     */
    public void setStyle(Style style) {
        java.util.Objects.requireNonNull(style, "style");
        getOrCreateRecord().setStyle(style.toModel());
    }

    /**
     * Assigns value.
     * @param value value to apply
     */
    private void assignValue(Object value) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            clearValue();
            return;
        }
        if (value instanceof String text) {
            putValue(text);
            return;
        }
        if (value instanceof Boolean booleanValue) {
            putValue(booleanValue);
            return;
        }
        if (value instanceof java.time.LocalDateTime dateTime) {
            putValue(dateTime);
            return;
        }
        if (value instanceof Number) {
            if (value instanceof Integer intValue) {
                putValue(intValue);
            } else if (value instanceof Double doubleValue) {
                putValue(doubleValue);
            } else {
                setScalar(value, CellValueKind.NUMBER);
            }
            return;
        }
        putValue(value.toString());
    }

    /**
     * Clears the current state maintained by this object.
     */
    private void clearValue() {
        CellRecord record = getOrCreateRecord();
        record.setValue(null);
        record.setFormula(null);
        record.setKind(CellValueKind.BLANK);
    }

    /**
     * Attempts to process get record.
     * @return the computed result
     */
    private CellRecord tryGetRecord() {
        return worksheet.getModel().getCells().get(address);
    }

    /**
     * Returns the or create record.
     * @return the requested result
     */
    private CellRecord getOrCreateRecord() {
        // Lazily create the backing record the first time this path is used.
        return worksheet.getModel().getCells().computeIfAbsent(address, k -> {
            CellRecord record = new CellRecord();
            record.setStyle(new com.aspose.cells_foss.core.StyleValue());
            return record;
        });
    }

    /**
     * Sets the scalar.
     * @param value value to apply
     * @param kind kind
     */
    private void setScalar(Object value, CellValueKind kind) {
        CellRecord record = getOrCreateRecord();
        record.setValue(value);
        record.setKind(kind);
        record.setFormula(null);
    }

    /**
     * Normalizes the formula.
     * @param value value to apply
     * @return the computed result
     */
    private static String normalizeFormula(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isBlank()) {
            return "";
        }
        return value.charAt(0) == '=' ? value.substring(1) : value;
    }
}
