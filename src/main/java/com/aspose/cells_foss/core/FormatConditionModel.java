package com.aspose.cells_foss.core;
import com.aspose.cells_foss.OperatorType;
import com.aspose.cells_foss.FormatConditionType;

/**
 * Represents a format condition model used in Excel conditional formatting.
 */
public final class FormatConditionModel {
    private FormatConditionType type;
    private OperatorType operator = OperatorType.NONE;
    private String formula1;
    private String formula2;
    private String timePeriod;
    private boolean duplicate = true;
    private boolean top = true;
    private boolean percent;
    private int rank;
    private boolean above = true;
    private int standardDeviation;
    private int colorScaleCount = 2;
    private ColorValue minColor;
    private ColorValue midColor;
    private ColorValue maxColor;
    private ColorValue barColor;
    private ColorValue negativeBarColor;
    private boolean showBorder;
    private String direction;
    private String barLength;
    private String iconSetType;
    private boolean reverseIcons;
    private boolean showIconOnly;
    private int priority;
    private boolean stopIfTrue;
    private StyleValue style = StyleValue.getDefault().clone();

    /**
     * Returns the type.
     * @return the requested result
     */
    public FormatConditionType getType() { return type; }
    /**
     * Sets the type.
     * @param type type
     */
    public void setType(FormatConditionType type) { this.type = type; }

    /**
     * Returns the operator.
     * @return the requested result
     */
    public OperatorType getOperator() { return operator; }
    /**
     * Sets the operator.
     * @param operator operator
     */
    public void setOperator(OperatorType operator) { this.operator = operator; }

    /**
     * Returns the formula 1.
     * @return the requested result
     */
    public String getFormula1() { return formula1; }
    /**
     * Sets the formula 1.
     * @param formula1 formula 1
     */
    public void setFormula1(String formula1) { this.formula1 = formula1; }

    /**
     * Returns the formula 2.
     * @return the requested result
     */
    public String getFormula2() { return formula2; }
    /**
     * Sets the formula 2.
     * @param formula2 formula 2
     */
    public void setFormula2(String formula2) { this.formula2 = formula2; }

    /**
     * Returns the time period.
     * @return the requested result
     */
    public String getTimePeriod() { return timePeriod; }
    /**
     * Sets the time period.
     * @param timePeriod time period
     */
    public void setTimePeriod(String timePeriod) { this.timePeriod = timePeriod; }

    /**
     * Returns the duplicate.
     * @return the requested result
     */
    public boolean getDuplicate() { return duplicate; }
    /**
     * Sets the duplicate.
     * @param duplicate duplicate
     */
    public void setDuplicate(boolean duplicate) { this.duplicate = duplicate; }

    /**
     * Returns the top.
     * @return the requested result
     */
    public boolean getTop() { return top; }
    /**
     * Sets the top.
     * @param top top
     */
    public void setTop(boolean top) { this.top = top; }

    /**
     * Returns the percent.
     * @return the requested result
     */
    public boolean getPercent() { return percent; }
    /**
     * Sets the percent.
     * @param percent percent
     */
    public void setPercent(boolean percent) { this.percent = percent; }

    /**
     * Returns the rank.
     * @return the requested result
     */
    public int getRank() { return rank; }
    /**
     * Sets the rank.
     * @param rank rank
     */
    public void setRank(int rank) { this.rank = rank; }

    /**
     * Returns the above.
     * @return the requested result
     */
    public boolean getAbove() { return above; }
    /**
     * Sets the above.
     * @param above above
     */
    public void setAbove(boolean above) { this.above = above; }

    /**
     * Returns the standard deviation.
     * @return the requested result
     */
    public int getStandardDeviation() { return standardDeviation; }
    /**
     * Sets the standard deviation.
     * @param standardDeviation standard deviation
     */
    public void setStandardDeviation(int standardDeviation) { this.standardDeviation = standardDeviation; }

    /**
     * Returns the color scale count.
     * @return the requested result
     */
    public int getColorScaleCount() { return colorScaleCount; }
    /**
     * Sets the color scale count.
     * @param colorScaleCount color scale count
     */
    public void setColorScaleCount(int colorScaleCount) { this.colorScaleCount = colorScaleCount; }

    /**
     * Returns the min color.
     * @return the requested result
     */
    public ColorValue getMinColor() { return minColor; }
    /**
     * Sets the min color.
     * @param minColor min color
     */
    public void setMinColor(ColorValue minColor) { this.minColor = minColor; }

    /**
     * Returns the mid color.
     * @return the requested result
     */
    public ColorValue getMidColor() { return midColor; }
    /**
     * Sets the mid color.
     * @param midColor mid color
     */
    public void setMidColor(ColorValue midColor) { this.midColor = midColor; }

    /**
     * Returns the max color.
     * @return the requested result
     */
    public ColorValue getMaxColor() { return maxColor; }
    /**
     * Sets the max color.
     * @param maxColor max color
     */
    public void setMaxColor(ColorValue maxColor) { this.maxColor = maxColor; }

    /**
     * Returns the bar color.
     * @return the requested result
     */
    public ColorValue getBarColor() { return barColor; }
    /**
     * Sets the bar color.
     * @param barColor bar color
     */
    public void setBarColor(ColorValue barColor) { this.barColor = barColor; }

    /**
     * Returns the negative bar color.
     * @return the requested result
     */
    public ColorValue getNegativeBarColor() { return negativeBarColor; }
    /**
     * Sets the negative bar color.
     * @param negativeBarColor negative bar color
     */
    public void setNegativeBarColor(ColorValue negativeBarColor) { this.negativeBarColor = negativeBarColor; }

    /**
     * Returns the show border.
     * @return the requested result
     */
    public boolean getShowBorder() { return showBorder; }
    /**
     * Sets the show border.
     * @param showBorder show border
     */
    public void setShowBorder(boolean showBorder) { this.showBorder = showBorder; }

    /**
     * Returns the direction.
     * @return the requested result
     */
    public String getDirection() { return direction; }
    /**
     * Sets the direction.
     * @param direction direction
     */
    public void setDirection(String direction) { this.direction = direction; }

    /**
     * Returns the bar length.
     * @return the requested result
     */
    public String getBarLength() { return barLength; }
    /**
     * Sets the bar length.
     * @param barLength bar length
     */
    public void setBarLength(String barLength) { this.barLength = barLength; }

    /**
     * Returns the icon set type.
     * @return the requested result
     */
    public String getIconSetType() { return iconSetType; }
    /**
     * Sets the icon set type.
     * @param iconSetType icon set type
     */
    public void setIconSetType(String iconSetType) { this.iconSetType = iconSetType; }

    /**
     * Returns the reverse icons.
     * @return the requested result
     */
    public boolean getReverseIcons() { return reverseIcons; }
    /**
     * Sets the reverse icons.
     * @param reverseIcons reverse icons
     */
    public void setReverseIcons(boolean reverseIcons) { this.reverseIcons = reverseIcons; }

    /**
     * Returns the show icon only.
     * @return the requested result
     */
    public boolean getShowIconOnly() { return showIconOnly; }
    /**
     * Sets the show icon only.
     * @param showIconOnly show icon only
     */
    public void setShowIconOnly(boolean showIconOnly) { this.showIconOnly = showIconOnly; }

    /**
     * Returns the priority.
     * @return the requested result
     */
    public int getPriority() { return priority; }
    /**
     * Sets the priority.
     * @param priority priority
     */
    public void setPriority(int priority) { this.priority = priority; }

    /**
     * Returns the stop if true.
     * @return the requested result
     */
    public boolean getStopIfTrue() { return stopIfTrue; }
    /**
     * Sets the stop if true.
     * @param stopIfTrue stop if true
     */
    public void setStopIfTrue(boolean stopIfTrue) { this.stopIfTrue = stopIfTrue; }

    /**
     * Returns the style.
     * @return the requested result
     */
    public StyleValue getStyle() { return style; }
    /**
     * Sets the style.
     * @param style style to apply
     */
    public void setStyle(StyleValue style) { this.style = style; }
}