package org.aspose.cells_foss;

import org.aspose.cells_foss.core.ConditionalFormattingModel;
import org.aspose.cells_foss.core.FormatConditionModel;
import org.aspose.cells_foss.core.StyleValue;
import java.util.List;

/**
 * Represents a conditional formatting rule.
 */
public final class FormatCondition {

    private final List<ConditionalFormattingModel> owner;
    private final ConditionalFormattingModel collection;
    private final FormatConditionModel model;

    /**
     * Initializes a new FormatCondition instance.
     * @param owner owner
     * @param collection collection
     * @param model model
     */
    FormatCondition(List<ConditionalFormattingModel> owner,
                    ConditionalFormattingModel collection,
                    FormatConditionModel model) {
        this.owner = owner;
        this.collection = collection;
        this.model = model;
    }

    /**
     * Returns the type.
     * @return the requested result
     */
    public FormatConditionType getType() { return model.getType(); }
    /**
     * Sets the type.
     * @param type type
     */
    public void setType(FormatConditionType type) { model.setType(type); }

    /**
     * Returns the operator.
     * @return the requested result
     */
    public OperatorType getOperator() { return model.getOperator(); }
    /**
     * Sets the operator.
     * @param operator operator
     */
    public void setOperator(OperatorType operator) { model.setOperator(operator); }

    /**
     * Returns the formula 1.
     * @return the requested result
     */
    public String getFormula1() { return nvl(model.getFormula1()); }
    /**
     * Sets the formula 1.
     * @param value value to apply
     */
    public void setFormula1(String value) { model.setFormula1(normalizeFormula(value)); }

    /**
     * Returns the formula 2.
     * @return the requested result
     */
    public String getFormula2() { return nvl(model.getFormula2()); }
    /**
     * Sets the formula 2.
     * @param value value to apply
     */
    public void setFormula2(String value) { model.setFormula2(normalizeFormula(value)); }

    /**
     * Returns the formula.
     * @return the requested result
     */
    public String getFormula() { return nvl(model.getFormula1()); }
    /**
     * Sets the formula.
     * @param value value to apply
     */
    public void setFormula(String value) { model.setFormula1(normalizeFormula(value)); }

    /**
     * Returns the time period.
     * @return the requested result
     */
    public String getTimePeriod() { return nvl(model.getTimePeriod()); }
    /**
     * Sets the time period.
     * @param value value to apply
     */
    public void setTimePeriod(String value) { model.setTimePeriod(normalizeText(value)); }

    /**
     * Returns the duplicate.
     * @return the requested result
     */
    public boolean getDuplicate() { return model.getDuplicate(); }
    /**
     * Sets the duplicate.
     * @param duplicate duplicate
     */
    public void setDuplicate(boolean duplicate) { model.setDuplicate(duplicate); }

    /**
     * Returns the top.
     * @return the requested result
     */
    public boolean getTop() { return model.getTop(); }
    /**
     * Sets the top.
     * @param top top
     */
    public void setTop(boolean top) { model.setTop(top); }

    /**
     * Returns the percent.
     * @return the requested result
     */
    public boolean getPercent() { return model.getPercent(); }
    /**
     * Sets the percent.
     * @param percent percent
     */
    public void setPercent(boolean percent) { model.setPercent(percent); }

    /**
     * Returns the rank.
     * @return the requested result
     */
    public int getRank() { return model.getRank(); }
    /**
     * Sets the rank.
     * @param rank rank
     */
    public void setRank(int rank) {
        // Validate the caller input before continuing.
        if (rank < 0) throw new CellsException("Conditional formatting rank must be zero or greater.");
        model.setRank(rank);
    }

    /**
     * Returns the above.
     * @return the requested result
     */
    public boolean getAbove() { return model.getAbove(); }
    /**
     * Sets the above.
     * @param above above
     */
    public void setAbove(boolean above) { model.setAbove(above); }

    /**
     * Returns the standard deviation.
     * @return the requested result
     */
    public int getStandardDeviation() { return model.getStandardDeviation(); }
    /**
     * Sets the standard deviation.
     * @param standardDeviation standard deviation
     */
    public void setStandardDeviation(int standardDeviation) {
        // Validate the caller input before continuing.
        if (standardDeviation < 0) throw new CellsException("Conditional formatting standard deviation must be zero or greater.");
        model.setStandardDeviation(standardDeviation);
    }

    /**
     * Returns the color scale count.
     * @return the requested result
     */
    public int getColorScaleCount() { return model.getColorScaleCount(); }
    /**
     * Sets the color scale count.
     * @param colorScaleCount color scale count
     */
    public void setColorScaleCount(int colorScaleCount) {
        // Validate the caller input before continuing.
        if (colorScaleCount != 2 && colorScaleCount != 3) throw new CellsException("ColorScaleCount must be 2 or 3.");
        model.setColorScaleCount(colorScaleCount);
    }

    /**
     * Returns the min color.
     * @return the requested result
     */
    public Color getMinColor() { return model.getMinColor() != null ? Color.fromCore(model.getMinColor()) : Color.getEmpty(); }
    /**
     * Sets the min color.
     * @param value value to apply
     */
    public void setMinColor(Color value) { model.setMinColor(value == null || value.equals(Color.getEmpty()) ? null : value.toCore()); }

    /**
     * Returns the mid color.
     * @return the requested result
     */
    public Color getMidColor() { return model.getMidColor() != null ? Color.fromCore(model.getMidColor()) : Color.getEmpty(); }
    /**
     * Sets the mid color.
     * @param value value to apply
     */
    public void setMidColor(Color value) { model.setMidColor(value == null || value.equals(Color.getEmpty()) ? null : value.toCore()); }

    /**
     * Returns the max color.
     * @return the requested result
     */
    public Color getMaxColor() { return model.getMaxColor() != null ? Color.fromCore(model.getMaxColor()) : Color.getEmpty(); }
    /**
     * Sets the max color.
     * @param value value to apply
     */
    public void setMaxColor(Color value) { model.setMaxColor(value == null || value.equals(Color.getEmpty()) ? null : value.toCore()); }

    /**
     * Returns the bar color.
     * @return the requested result
     */
    public Color getBarColor() { return model.getBarColor() != null ? Color.fromCore(model.getBarColor()) : Color.getEmpty(); }
    /**
     * Sets the bar color.
     * @param value value to apply
     */
    public void setBarColor(Color value) { model.setBarColor(value == null || value.equals(Color.getEmpty()) ? null : value.toCore()); }

    /**
     * Returns the negative bar color.
     * @return the requested result
     */
    public Color getNegativeBarColor() { return model.getNegativeBarColor() != null ? Color.fromCore(model.getNegativeBarColor()) : Color.getEmpty(); }
    /**
     * Sets the negative bar color.
     * @param value value to apply
     */
    public void setNegativeBarColor(Color value) { model.setNegativeBarColor(value == null || value.equals(Color.getEmpty()) ? null : value.toCore()); }

    /**
     * Returns the show border.
     * @return the requested result
     */
    public boolean getShowBorder() { return model.getShowBorder(); }
    /**
     * Sets the show border.
     * @param showBorder show border
     */
    public void setShowBorder(boolean showBorder) { model.setShowBorder(showBorder); }

    /**
     * Returns the direction.
     * @return the requested result
     */
    public String getDirection() { return nvl(model.getDirection()); }
    /**
     * Sets the direction.
     * @param value value to apply
     */
    public void setDirection(String value) { model.setDirection(normalizeText(value)); }

    /**
     * Returns the bar length.
     * @return the requested result
     */
    public String getBarLength() { return nvl(model.getBarLength()); }
    /**
     * Sets the bar length.
     * @param value value to apply
     */
    public void setBarLength(String value) { model.setBarLength(normalizeText(value)); }

    /**
     * Returns the icon set type.
     * @return the requested result
     */
    public String getIconSetType() { return nvl(model.getIconSetType()); }
    /**
     * Sets the icon set type.
     * @param value value to apply
     */
    public void setIconSetType(String value) { model.setIconSetType(normalizeText(value)); }

    /**
     * Returns the reverse icons.
     * @return the requested result
     */
    public boolean getReverseIcons() { return model.getReverseIcons(); }
    /**
     * Sets the reverse icons.
     * @param reverseIcons reverse icons
     */
    public void setReverseIcons(boolean reverseIcons) { model.setReverseIcons(reverseIcons); }

    /**
     * Returns the show icon only.
     * @return the requested result
     */
    public boolean getShowIconOnly() { return model.getShowIconOnly(); }
    /**
     * Sets the show icon only.
     * @param showIconOnly show icon only
     */
    public void setShowIconOnly(boolean showIconOnly) { model.setShowIconOnly(showIconOnly); }

    /**
     * Returns the priority.
     * @return the requested result
     */
    public int getPriority() { return model.getPriority(); }
    /**
     * Sets the priority.
     * @param priority priority
     */
    public void setPriority(int priority) {
        // Validate the caller input before continuing.
        if (priority <= 0) throw new CellsException("Conditional formatting priority must be greater than zero.");
        model.setPriority(priority);
    }

    /**
     * Returns the stop if true.
     * @return the requested result
     */
    public boolean getStopIfTrue() { return model.getStopIfTrue(); }
    /**
     * Sets the stop if true.
     * @param stopIfTrue stop if true
     */
    public void setStopIfTrue(boolean stopIfTrue) { model.setStopIfTrue(stopIfTrue); }

    /**
     * Returns the style.
     * @return the requested result
     */
    public Style getStyle() { return Style.fromModel(model.getStyle()).clone(); }
    /**
     * Sets the style.
     * @param value value to apply
     */
    public void setStyle(Style value) {
        model.setStyle(value == null ? StyleValue.getDefault().clone() : value.toModel());
    }

    /**
     * Removes the requested content.
     */
    public void remove() {
        FormatConditionCollection.removeCondition(owner, collection, model);
    }

    // --- helpers ---
    /**
     * Returns the fallback value when the primary value is empty.
     * @param s s
     * @return the computed result
     */
    private static String nvl(String s) { return s == null ? "" : s; }

    /**
     * Normalizes the formula.
     * @param v v
     * @return the computed result
     */
    private static String normalizeFormula(String v) {
        // Handle the relevant branch before the state changes.
        if (v == null) return null;
        String t = v.trim();
        if (t.isEmpty()) return null;
        return t.charAt(0) == '=' ? t.substring(1) : t;
    }

    /**
     * Normalizes the text.
     * @param v v
     * @return the computed result
     */
    private static String normalizeText(String v) {
        // Handle the relevant branch before the state changes.
        if (v == null) return null;
        String t = v.trim();
        return t.isEmpty() ? null : t;
    }
}

