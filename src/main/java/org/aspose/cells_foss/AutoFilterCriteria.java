package org.aspose.cells_foss;

import org.aspose.cells_foss.core.AutoFilterModel;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Represents auto-filter criteria for filtering data in a worksheet.
 */
public class AutoFilterCriteria {

    /**
     * Represents a collection of auto-filter custom filters.
     */
    public static final class AutoFilterCustomFilterCollection implements Iterable<AutoFilterCustomFilter> {
        private final List<AutoFilterModel.AutoFilterCustomFilterModel> models;
        private final AutoFilterModel.FilterColumnModel columnModel;

        /**
         * Initializes a new AutoFilterCustomFilterCollection instance.
         * @param models models
         * @param columnModel column model
         */
        AutoFilterCustomFilterCollection(List<AutoFilterModel.AutoFilterCustomFilterModel> models, AutoFilterModel.FilterColumnModel columnModel) {
            this.models = models;
            this.columnModel = columnModel;
        }

        /**
         * Indicates whether match all.
         * @return true when the condition is satisfied
         */
        public boolean isMatchAll() {
            return columnModel.isCustomFiltersAnd();
        }

        /**
         * Sets the match all.
         * @param matchAll match all
         */
        public void setMatchAll(boolean matchAll) {
            columnModel.setCustomFiltersAnd(matchAll);
        }

        /**
         * Returns the count.
         * @return the requested result
         */
        public int getCount() {
            return models.size();
        }

        /**
         * Returns the requested item.
         * @param index index
         * @return the requested result
         */
        public AutoFilterCustomFilter get(int index) {
            // Handle the relevant branch before the state changes.
            if (index < 0 || index >= models.size()) {
                throw new CellsException("Custom filter index was out of range.");
            }
            return new AutoFilterCustomFilter(models.get(index));
        }

        /**
         * Adds a new item to the current collection.
         * @param operatorType operator type
         * @param value value to apply
         * @return the computed result
         */
        public int add(AutoFilterModel.FilterOperatorType operatorType, String value) {
            // Handle the relevant branch before the state changes.
            if (models.size() >= 2) {
                throw new CellsException("Custom filters support at most two filter conditions.");
            }
            AutoFilterModel.AutoFilterCustomFilterModel model = new AutoFilterModel.AutoFilterCustomFilterModel();
            model.setOperator(AutoFilterSupport.toOperatorName(operatorType));
            model.setValue(AutoFilterSupport.normalizeText(value, "value"));
            models.add(model);
            return models.size() - 1;
        }

        /**
         * Removes at.
         * @param index index
         */
        public void removeAt(int index) {
            // Handle the relevant branch before the state changes.
            if (index < 0 || index >= models.size()) {
                throw new CellsException("Custom filter index was out of range.");
            }
            models.remove(index);
        }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            models.clear();
            columnModel.setCustomFiltersAnd(false);
        }

        /**
         * Returns an iterator over the current collection.
         * @return the computed result
         */
        @Override
        public Iterator<AutoFilterCustomFilter> iterator() {
            List<AutoFilterCustomFilter> filters = new ArrayList<>(models.size());
            // Walk the current collection so every entry is processed consistently.
            for (int i = 0; i < models.size(); i++) {
                filters.add(new AutoFilterCustomFilter(models.get(i)));
            }
            return filters.iterator();
        }
    }

    /**
     * Represents a custom filter in an auto-filter collection.
     */
    public static final class AutoFilterCustomFilter {
        private final AutoFilterModel.AutoFilterCustomFilterModel model;

        /**
         * Initializes a new AutoFilterCustomFilter instance.
         * @param model model
         */
        AutoFilterCustomFilter(AutoFilterModel.AutoFilterCustomFilterModel model) {
            this.model = model;
        }

        /**
         * Returns the operator.
         * @return the requested result
         */
        public AutoFilterModel.FilterOperatorType getOperator() {
            return AutoFilterSupport.parseOperatorOrDefault(model.getOperator());
        }

        /**
         * Sets the operator.
         * @param operator operator
         */
        public void setOperator(AutoFilterModel.FilterOperatorType operator) {
            model.setOperator(AutoFilterSupport.toOperatorName(operator));
        }

        /**
         * Returns the value.
         * @return the requested result
         */
        public String getValue() {
            return model.getValue();
        }

        /**
         * Sets the value.
         * @param value value to apply
         */
        public void setValue(String value) {
            model.setValue(AutoFilterSupport.normalizeText(value, "value"));
        }
    }

    /**
     * Represents a color filter for auto-filtering.
     */
    public static final class AutoFilterColorFilter {
        private final AutoFilterModel.AutoFilterColorFilterModel model;

        /**
         * Initializes a new AutoFilterColorFilter instance.
         * @param model model
         */
        AutoFilterColorFilter(AutoFilterModel.AutoFilterColorFilterModel model) {
            this.model = model;
        }

        /**
         * Indicates whether enabled.
         * @return true when the condition is satisfied
         */
        public boolean isEnabled() {
            return model.isEnabled();
        }

        /**
         * Sets the enabled.
         * @param enabled enabled
         */
        public void setEnabled(boolean enabled) {
            // Handle the relevant branch before the state changes.
            if (!enabled) {
                model.clear();
                return;
            }
            model.setEnabled(true);
        }

        /**
         * Returns the differential style id.
         * @return the requested result
         */
        public Integer getDifferentialStyleId() {
            return model.getDifferentialStyleId();
        }

        /**
         * Sets the differential style id.
         * @param value value to apply
         */
        public void setDifferentialStyleId(Integer value) {
            // Handle the relevant branch before the state changes.
            if (value != null && value < 0) {
                throw new CellsException("Differential style id must be zero or greater.");
            }
            model.setDifferentialStyleId(value);
            if (value != null) {
                model.setEnabled(true);
            }
        }

        /**
         * Indicates whether cell color.
         * @return true when the condition is satisfied
         */
        public boolean isCellColor() {
            return model.isCellColor();
        }

        /**
         * Sets the cell color.
         * @param cellColor cell color
         */
        public void setCellColor(boolean cellColor) {
            model.setCellColor(cellColor);
            // Handle the relevant branch before the state changes.
            if (cellColor) {
                model.setEnabled(true);
            }
        }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            model.clear();
        }
    }

    /**
     * Represents a dynamic filter for auto-filtering.
     */
    public static final class AutoFilterDynamicFilter {
        private final AutoFilterModel.AutoFilterDynamicFilterModel model;

        /**
         * Initializes a new AutoFilterDynamicFilter instance.
         * @param model model
         */
        AutoFilterDynamicFilter(AutoFilterModel.AutoFilterDynamicFilterModel model) {
            this.model = model;
        }

        /**
         * Indicates whether enabled.
         * @return true when the condition is satisfied
         */
        public boolean isEnabled() {
            return model.isEnabled();
        }

        /**
         * Sets the enabled.
         * @param enabled enabled
         */
        public void setEnabled(boolean enabled) {
            // Handle the relevant branch before the state changes.
            if (!enabled) {
                model.clear();
                return;
            }
            model.setEnabled(true);
        }

        /**
         * Returns the type.
         * @return the requested result
         */
        public String getType() {
            return model.getType();
        }

        /**
         * Sets the type.
         * @param type type
         */
        public void setType(String type) {
            model.setType(AutoFilterSupport.normalizeOptionalText(type));
            // Handle the relevant branch before the state changes.
            if (!model.getType().isEmpty()) {
                model.setEnabled(true);
            }
        }

        /**
         * Returns the value.
         * @return the requested result
         */
        public Double getValue() {
            return model.getValue();
        }

        /**
         * Sets the value.
         * @param value value to apply
         */
        public void setValue(Double value) {
            model.setValue(value);
            // Handle the relevant branch before the state changes.
            if (value != null) {
                model.setEnabled(true);
            }
        }

        /**
         * Returns the max value.
         * @return the requested result
         */
        public Double getMaxValue() {
            return model.getMaxValue();
        }

        /**
         * Sets the max value.
         * @param maxValue max value
         */
        public void setMaxValue(Double maxValue) {
            model.setMaxValue(maxValue);
            // Handle the relevant branch before the state changes.
            if (maxValue != null) {
                model.setEnabled(true);
            }
        }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            model.clear();
        }
    }

    /**
     * Represents a top 10 filter for auto-filtering.
     */
    public static final class AutoFilterTop10 {
        private final AutoFilterModel.AutoFilterTop10Model model;

        /**
         * Initializes a new AutoFilterTop10 instance.
         * @param model model
         */
        AutoFilterTop10(AutoFilterModel.AutoFilterTop10Model model) {
            this.model = model;
        }

        /**
         * Indicates whether enabled.
         * @return true when the condition is satisfied
         */
        public boolean isEnabled() {
            return model.isEnabled();
        }

        /**
         * Sets the enabled.
         * @param enabled enabled
         */
        public void setEnabled(boolean enabled) {
            // Handle the relevant branch before the state changes.
            if (!enabled) {
                model.clear();
                return;
            }
            model.setEnabled(true);
        }

        /**
         * Indicates whether top.
         * @return true when the condition is satisfied
         */
        public boolean isTop() {
            return model.isTop();
        }

        /**
         * Sets the top.
         * @param top top
         */
        public void setTop(boolean top) {
            model.setTop(top);
            model.setEnabled(true);
        }

        /**
         * Indicates whether percent.
         * @return true when the condition is satisfied
         */
        public boolean isPercent() {
            return model.isPercent();
        }

        /**
         * Sets the percent.
         * @param percent percent
         */
        public void setPercent(boolean percent) {
            model.setPercent(percent);
            model.setEnabled(true);
        }

        /**
         * Returns the value.
         * @return the requested result
         */
        public Double getValue() {
            return model.getValue();
        }

        /**
         * Sets the value.
         * @param value value to apply
         */
        public void setValue(Double value) {
            model.setValue(value);
            // Handle the relevant branch before the state changes.
            if (value != null) {
                model.setEnabled(true);
            }
        }

        /**
         * Returns the filter value.
         * @return the requested result
         */
        public Double getFilterValue() {
            return model.getFilterValue();
        }

        /**
         * Sets the filter value.
         * @param filterValue filter value
         */
        public void setFilterValue(Double filterValue) {
            model.setFilterValue(filterValue);
            // Handle the relevant branch before the state changes.
            if (filterValue != null) {
                model.setEnabled(true);
            }
        }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            model.clear();
        }
    }

    /**
     * Represents the sort state for auto-filtering.
     */
    public static final class AutoFilterSortState {
        private final AutoFilterModel.AutoFilterSortStateModel model;
        private final AutoFilterSortConditionCollection conditions;

        /**
         * Initializes a new AutoFilterSortState instance.
         * @param model model
         */
        AutoFilterSortState(AutoFilterModel.AutoFilterSortStateModel model) {
            this.model = model;
            this.conditions = new AutoFilterSortConditionCollection(model.getConditions());
        }

        /**
         * Indicates whether column sort.
         * @return true when the condition is satisfied
         */
        public boolean isColumnSort() {
            return model.isColumnSort();
        }

        /**
         * Sets the column sort.
         * @param columnSort column sort
         */
        public void setColumnSort(boolean columnSort) {
            model.setColumnSort(columnSort);
        }

        /**
         * Indicates whether case sensitive.
         * @return true when the condition is satisfied
         */
        public boolean isCaseSensitive() {
            return model.isCaseSensitive();
        }

        /**
         * Sets the case sensitive.
         * @param caseSensitive case sensitive
         */
        public void setCaseSensitive(boolean caseSensitive) {
            model.setCaseSensitive(caseSensitive);
        }

        /**
         * Returns the sort method.
         * @return the requested result
         */
        public String getSortMethod() {
            return model.getSortMethod();
        }

        /**
         * Sets the sort method.
         * @param sortMethod sort method
         */
        public void setSortMethod(String sortMethod) {
            model.setSortMethod(AutoFilterSupport.normalizeOptionalText(sortMethod));
        }

        /**
         * Returns the ref.
         * @return the requested result
         */
        public String getRef() {
            return model.getRef();
        }

        /**
         * Sets the ref.
         * @param ref ref
         */
        public void setRef(String ref) {
            model.setRef(AutoFilterSupport.normalizeOptionalRange(ref, "ref"));
        }

        /**
         * Returns the sort conditions.
         * @return the requested result
         */
        public AutoFilterSortConditionCollection getSortConditions() {
            return conditions;
        }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            model.clear();
        }
    }

    /**
     * Represents a collection of auto-filter sort conditions.
     */
    public static final class AutoFilterSortConditionCollection implements Iterable<AutoFilterSortCondition> {
        private final List<AutoFilterModel.AutoFilterSortConditionModel> models;

        /**
         * Initializes a new AutoFilterSortConditionCollection instance.
         * @param models models
         */
        AutoFilterSortConditionCollection(List<AutoFilterModel.AutoFilterSortConditionModel> models) {
            this.models = models;
        }

        /**
         * Returns the count.
         * @return the requested result
         */
        public int getCount() {
            return models.size();
        }

        /**
         * Returns the requested item.
         * @param index index
         * @return the requested result
         */
        public AutoFilterSortCondition get(int index) {
            // Handle the relevant branch before the state changes.
            if (index < 0 || index >= models.size()) {
                throw new CellsException("Sort condition index was out of range.");
            }
            return new AutoFilterSortCondition(models.get(index));
        }

        /**
         * Adds a new item to the current collection.
         * @param reference reference
         * @return the computed result
         */
        public int add(String reference) {
            AutoFilterModel.AutoFilterSortConditionModel model = new AutoFilterModel.AutoFilterSortConditionModel();
            model.setRef(AutoFilterSupport.normalizeRequiredRange(reference, "reference"));
            models.add(model);
            return models.size() - 1;
        }

        /**
         * Removes at.
         * @param index index
         */
        public void removeAt(int index) {
            // Handle the relevant branch before the state changes.
            if (index < 0 || index >= models.size()) {
                throw new CellsException("Sort condition index was out of range.");
            }
            models.remove(index);
        }

        /**
         * Clears the current state maintained by this object.
         */
        public void clear() {
            models.clear();
        }

        /**
         * Returns an iterator over the current collection.
         * @return the computed result
         */
        @Override
        public Iterator<AutoFilterSortCondition> iterator() {
            List<AutoFilterSortCondition> conditions = new ArrayList<>(models.size());
            // Walk the current collection so every entry is processed consistently.
            for (int i = 0; i < models.size(); i++) {
                conditions.add(new AutoFilterSortCondition(models.get(i)));
            }
            return conditions.iterator();
        }
    }

    /**
     * Represents a sort condition for auto-filtering.
     */
    public static final class AutoFilterSortCondition {
        private final AutoFilterModel.AutoFilterSortConditionModel model;

        /**
         * Initializes a new AutoFilterSortCondition instance.
         * @param model model
         */
        AutoFilterSortCondition(AutoFilterModel.AutoFilterSortConditionModel model) {
            this.model = model;
        }

        /**
         * Returns the ref.
         * @return the requested result
         */
        public String getRef() {
            return model.getRef();
        }

        /**
         * Sets the ref.
         * @param ref ref
         */
        public void setRef(String ref) {
            model.setRef(AutoFilterSupport.normalizeRequiredRange(ref, "ref"));
        }

        /**
         * Indicates whether descending.
         * @return true when the condition is satisfied
         */
        public boolean isDescending() {
            return model.isDescending();
        }

        /**
         * Sets the descending.
         * @param descending descending
         */
        public void setDescending(boolean descending) {
            model.setDescending(descending);
        }

        /**
         * Returns the sort by.
         * @return the requested result
         */
        public String getSortBy() {
            return model.getSortBy();
        }

        /**
         * Sets the sort by.
         * @param sortBy sort by
         */
        public void setSortBy(String sortBy) {
            model.setSortBy(AutoFilterSupport.normalizeOptionalText(sortBy));
        }

        /**
         * Returns the custom list.
         * @return the requested result
         */
        public String getCustomList() {
            return model.getCustomList();
        }

        /**
         * Sets the custom list.
         * @param customList custom list
         */
        public void setCustomList(String customList) {
            model.setCustomList(AutoFilterSupport.normalizeOptionalText(customList));
        }

        /**
         * Returns the differential style id.
         * @return the requested result
         */
        public Integer getDifferentialStyleId() {
            return model.getDifferentialStyleId();
        }

        /**
         * Sets the differential style id.
         * @param value value to apply
         */
        public void setDifferentialStyleId(Integer value) {
            // Handle the relevant branch before the state changes.
            if (value != null && value < 0) {
                throw new CellsException("Differential style id must be zero or greater.");
            }
            model.setDifferentialStyleId(value);
        }

        /**
         * Returns the icon set.
         * @return the requested result
         */
        public String getIconSet() {
            return model.getIconSet();
        }

        /**
         * Sets the icon set.
         * @param iconSet icon set
         */
        public void setIconSet(String iconSet) {
            model.setIconSet(AutoFilterSupport.normalizeOptionalText(iconSet));
        }

        /**
         * Returns the icon id.
         * @return the requested result
         */
        public Integer getIconId() {
            return model.getIconId();
        }

        /**
         * Sets the icon id.
         * @param value value to apply
         */
        public void setIconId(Integer value) {
            // Handle the relevant branch before the state changes.
            if (value != null && value < 0) {
                throw new CellsException("Icon id must be zero or greater.");
            }
            model.setIconId(value);
        }
    }

    /**
     * Initializes a new AutoFilterCriteria instance.
     */
    private AutoFilterCriteria() {
        // Utility class
    }
}
