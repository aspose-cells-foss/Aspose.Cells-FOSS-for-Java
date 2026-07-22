package org.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.List;

/**
 * Represents a model for auto-filter configuration in Excel.
 */
public final class AutoFilterModel {
    private String range;
    private final List<FilterColumnModel> filterColumns;
    private final AutoFilterSortStateModel sortState;

    /**
     * Initializes a new AutoFilterModel instance.
     */
    public AutoFilterModel() {
        this.range = "";
        this.filterColumns = new ArrayList<>();
        this.sortState = new AutoFilterSortStateModel();
    }

    /**
     * Returns the range.
     * @return the requested result
     */
    public String getRange() { return range; }
    /**
     * Sets the range.
     * @param range range
     */
    public void setRange(String range) { this.range = range; }

    /**
     * Returns the filter columns.
     * @return the requested result
     */
    public List<FilterColumnModel> getFilterColumns() { return filterColumns; }

    /**
     * Returns the sort state.
     * @return the requested result
     */
    public AutoFilterSortStateModel getSortState() { return sortState; }

    /**
     * Clears the current state maintained by this object.
     */
    public void clear() {
        range = "";
        filterColumns.clear();
        sortState.clear();
    }

    /**
     * Indicates whether this instance has stored state.
     * @return true when the condition is satisfied
     */
    public boolean hasStoredState() {
        return !range.isEmpty();
    }

    // Inner class: FilterColumnModel
    /**
     * Represents the FilterColumnModel component.
     */
    public static final class FilterColumnModel {
        private int columnIndex;
        private boolean hiddenButton;
        private final List<String> filters;
        private final List<AutoFilterCustomFilterModel> customFilters;
        private boolean customFiltersAnd;
        private final AutoFilterColorFilterModel colorFilter;
        private final AutoFilterDynamicFilterModel dynamicFilter;
        private final AutoFilterTop10Model top10;

        /**
         * Initializes a new FilterColumnModel instance.
         */
        public FilterColumnModel() {
            this.filters = new ArrayList<>();
            this.customFilters = new ArrayList<>();
            this.colorFilter = new AutoFilterColorFilterModel();
            this.dynamicFilter = new AutoFilterDynamicFilterModel();
            this.top10 = new AutoFilterTop10Model();
        }

        /**
         * Returns the column index.
         * @return the requested result
         */
        public int getColumnIndex() { return columnIndex; }
        /**
         * Sets the column index.
         * @param value value to apply
         */
        public void setColumnIndex(int value) { this.columnIndex = value; }

        /**
         * Returns the hidden button.
         * @return the requested result
         */
        public boolean getHiddenButton() { return hiddenButton; }
        /**
         * Sets the hidden button.
         * @param value value to apply
         */
        public void setHiddenButton(boolean value) { this.hiddenButton = value; }

        /**
         * Returns the filters.
         * @return the requested result
         */
        public List<String> getFilters() { return filters; }

        /**
         * Returns the custom filters.
         * @return the requested result
         */
        public List<AutoFilterCustomFilterModel> getCustomFilters() { return customFilters; }

        /**
         * Indicates whether custom filters and.
         * @return true when the condition is satisfied
         */
        public boolean isCustomFiltersAnd() { return customFiltersAnd; }
        /**
         * Sets the custom filters and.
         * @param value value to apply
         */
        public void setCustomFiltersAnd(boolean value) { this.customFiltersAnd = value; }

        /**
         * Returns the color filter.
         * @return the requested result
         */
        public AutoFilterColorFilterModel getColorFilter() { return colorFilter; }

        /**
         * Returns the dynamic filter.
         * @return the requested result
         */
        public AutoFilterDynamicFilterModel getDynamicFilter() { return dynamicFilter; }

        /**
         * Returns the top 10.
         * @return the requested result
         */
        public AutoFilterTop10Model getTop10() { return top10; }

        /**
         * Clears the current state maintained by this object.
         */
        public void clearCriteria() {
            hiddenButton = false;
            filters.clear();
            customFilters.clear();
            customFiltersAnd = false;
            colorFilter.clear();
            dynamicFilter.clear();
            top10.clear();
        }

        /**
         * Indicates whether this instance has stored state.
         * @return true when the condition is satisfied
         */
        public boolean hasStoredState() {
            return hiddenButton
                || filters.size() > 0
                || customFilters.size() > 0
                || colorFilter.isEnabled()
                || dynamicFilter.isEnabled()
                || top10.isEnabled();
        }
    }

    // Inner class: AutoFilterCustomFilterModel
    /**
     * Represents the AutoFilterCustomFilterModel component.
     */
    public static final class AutoFilterCustomFilterModel {
        private String operator = "";
        private String value = "";

        /**
         * Returns the operator.
         * @return the requested result
         */
        public String getOperator() { return operator; }
        /**
         * Sets the operator.
         * @param value value to apply
         */
        public void setOperator(String value) { this.operator = value; }

        /**
         * Returns the value.
         * @return the requested result
         */
        public String getValue() { return value; }
        /**
         * Sets the value.
         * @param value value to apply
         */
        public void setValue(String value) { this.value = value; }
    }

    // Inner class: AutoFilterColorFilterModel
    /**
     * Represents the AutoFilterColorFilterModel component.
     */
    public static final class AutoFilterColorFilterModel {
        private boolean enabled;
        private Integer differentialStyleId;
        private boolean cellColor;

        /**
         * Indicates whether enabled.
         * @return true when the condition is satisfied
         */
        public boolean isEnabled() { return enabled; }
        /**
         * Sets the enabled.
         * @param value value to apply
         */
        public void setEnabled(boolean value) { this.enabled = value; }

        /**
         * Returns the differential style id.
         * @return the requested result
         */
        public Integer getDifferentialStyleId() { return differentialStyleId; }
        /**
         * Sets the differential style id.
         * @param value value to apply
         */
        public void setDifferentialStyleId(Integer value) { this.differentialStyleId = value; }

        /**
         * Indicates whether cell color.
         * @return true when the condition is satisfied
         */
        public boolean isCellColor() { return cellColor; }
        /**
         * Sets the cell color.
         * @param value value to apply
         */
        public void setCellColor(boolean value) { this.cellColor = value; }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            enabled = false;
            differentialStyleId = null;
            cellColor = false;
        }
    }

    // Inner class: AutoFilterDynamicFilterModel
    /**
     * Represents the AutoFilterDynamicFilterModel component.
     */
    public static final class AutoFilterDynamicFilterModel {
        private boolean enabled;
        private String type = "";
        private Double value;
        private Double maxValue;

        /**
         * Indicates whether enabled.
         * @return true when the condition is satisfied
         */
        public boolean isEnabled() { return enabled; }
        /**
         * Sets the enabled.
         * @param value value to apply
         */
        public void setEnabled(boolean value) { this.enabled = value; }

        /**
         * Returns the type.
         * @return the requested result
         */
        public String getType() { return type; }
        /**
         * Sets the type.
         * @param value value to apply
         */
        public void setType(String value) { this.type = value; }

        /**
         * Returns the value.
         * @return the requested result
         */
        public Double getValue() { return value; }
        /**
         * Sets the value.
         * @param value value to apply
         */
        public void setValue(Double value) { this.value = value; }

        /**
         * Returns the max value.
         * @return the requested result
         */
        public Double getMaxValue() { return maxValue; }
        /**
         * Sets the max value.
         * @param value value to apply
         */
        public void setMaxValue(Double value) { this.maxValue = value; }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            enabled = false;
            type = "";
            value = null;
            maxValue = null;
        }
    }

    // Inner class: AutoFilterTop10Model
    /**
     * Represents the AutoFilterTop10Model component.
     */
    public static final class AutoFilterTop10Model {
        private boolean enabled;
        private boolean top = true;
        private boolean percent;
        private Double value;
        private Double filterValue;

        /**
         * Indicates whether enabled.
         * @return true when the condition is satisfied
         */
        public boolean isEnabled() { return enabled; }
        /**
         * Sets the enabled.
         * @param value value to apply
         */
        public void setEnabled(boolean value) { this.enabled = value; }

        /**
         * Indicates whether top.
         * @return true when the condition is satisfied
         */
        public boolean isTop() { return top; }
        /**
         * Sets the top.
         * @param value value to apply
         */
        public void setTop(boolean value) { this.top = value; }

        /**
         * Indicates whether percent.
         * @return true when the condition is satisfied
         */
        public boolean isPercent() { return percent; }
        /**
         * Sets the percent.
         * @param value value to apply
         */
        public void setPercent(boolean value) { this.percent = value; }

        /**
         * Returns the value.
         * @return the requested result
         */
        public Double getValue() { return value; }
        /**
         * Sets the value.
         * @param value value to apply
         */
        public void setValue(Double value) { this.value = value; }

        /**
         * Returns the filter value.
         * @return the requested result
         */
        public Double getFilterValue() { return filterValue; }
        /**
         * Sets the filter value.
         * @param value value to apply
         */
        public void setFilterValue(Double value) { this.filterValue = value; }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            enabled = false;
            top = true;
            percent = false;
            value = null;
            filterValue = null;
        }
    }

    // Inner class: AutoFilterSortStateModel
    /**
     * Represents the AutoFilterSortStateModel component.
     */
    public static final class AutoFilterSortStateModel {
        private boolean columnSort;
        private boolean caseSensitive;
        private String sortMethod = "";
        private String ref = "";
        private final List<AutoFilterSortConditionModel> conditions;

        /**
         * Initializes a new AutoFilterSortStateModel instance.
         */
        public AutoFilterSortStateModel() {
            this.conditions = new ArrayList<>();
        }

        /**
         * Indicates whether column sort.
         * @return true when the condition is satisfied
         */
        public boolean isColumnSort() { return columnSort; }
        /**
         * Sets the column sort.
         * @param value value to apply
         */
        public void setColumnSort(boolean value) { this.columnSort = value; }

        /**
         * Indicates whether case sensitive.
         * @return true when the condition is satisfied
         */
        public boolean isCaseSensitive() { return caseSensitive; }
        /**
         * Sets the case sensitive.
         * @param value value to apply
         */
        public void setCaseSensitive(boolean value) { this.caseSensitive = value; }

        /**
         * Returns the sort method.
         * @return the requested result
         */
        public String getSortMethod() { return sortMethod; }
        /**
         * Sets the sort method.
         * @param value value to apply
         */
        public void setSortMethod(String value) { this.sortMethod = value; }

        /**
         * Returns the ref.
         * @return the requested result
         */
        public String getRef() { return ref; }
        /**
         * Sets the ref.
         * @param value value to apply
         */
        public void setRef(String value) { this.ref = value; }

        /**
         * Returns the conditions.
         * @return the requested result
         */
        public List<AutoFilterSortConditionModel> getConditions() { return conditions; }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            columnSort = false;
            caseSensitive = false;
            sortMethod = "";
            ref = "";
            conditions.clear();
        }

        /**
         * Indicates whether this instance has stored state.
         * @return true when the condition is satisfied
         */
        public boolean hasStoredState() {
            return !ref.isEmpty()
                || columnSort
                || caseSensitive
                || !sortMethod.isEmpty()
                || conditions.size() > 0;
        }
    }

    // Inner class: AutoFilterSortConditionModel
    /**
     * Represents the AutoFilterSortConditionModel component.
     */
    public static final class AutoFilterSortConditionModel {
        private String ref = "";
        private boolean descending;
        private String sortBy = "";
        private String customList = "";
        private Integer differentialStyleId;
        private String iconSet = "";
        private Integer iconId;

        /**
         * Returns the ref.
         * @return the requested result
         */
        public String getRef() { return ref; }
        /**
         * Sets the ref.
         * @param value value to apply
         */
        public void setRef(String value) { this.ref = value; }

        /**
         * Indicates whether descending.
         * @return true when the condition is satisfied
         */
        public boolean isDescending() { return descending; }
        /**
         * Sets the descending.
         * @param value value to apply
         */
        public void setDescending(boolean value) { this.descending = value; }

        /**
         * Returns the sort by.
         * @return the requested result
         */
        public String getSortBy() { return sortBy; }
        /**
         * Sets the sort by.
         * @param value value to apply
         */
        public void setSortBy(String value) { this.sortBy = value; }

        /**
         * Returns the custom list.
         * @return the requested result
         */
        public String getCustomList() { return customList; }
        /**
         * Sets the custom list.
         * @param value value to apply
         */
        public void setCustomList(String value) { this.customList = value; }

        /**
         * Returns the differential style id.
         * @return the requested result
         */
        public Integer getDifferentialStyleId() { return differentialStyleId; }
        /**
         * Sets the differential style id.
         * @param value value to apply
         */
        public void setDifferentialStyleId(Integer value) { this.differentialStyleId = value; }

        /**
         * Returns the icon set.
         * @return the requested result
         */
        public String getIconSet() { return iconSet; }
        /**
         * Sets the icon set.
         * @param value value to apply
         */
        public void setIconSet(String value) { this.iconSet = value; }

        /**
         * Returns the icon id.
         * @return the requested result
         */
        public Integer getIconId() { return iconId; }
        /**
         * Sets the icon id.
         * @param value value to apply
         */
        public void setIconId(Integer value) { this.iconId = value; }
    }

    // Inner class: FilterOperatorType
    /**
     * Enumerates the supported FilterOperatorType values.
     */
    public enum FilterOperatorType {
        /** Equal to. */
        EQUAL,
        /** Less than. */
        LESS_THAN,
        /** Less than or equal. */
        LESS_OR_EQUAL,
        /** Not equal. */
        NOT_EQUAL,
        /** Greater than or equal. */
        GREATER_OR_EQUAL,
        /** Greater than. */
        GREATER_THAN
    }
}
