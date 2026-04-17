package com.aspose.cells_foss;

import com.aspose.cells_foss.core.DefinedNameModel;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Represents a collection of defined names in a workbook.
 */
public final class DefinedNameCollection implements Iterable<DefinedName> {
    private final Workbook workbook;

    /**
     * Initializes a new DefinedNameCollection instance.
     * @param workbook workbook to apply
     */
    DefinedNameCollection(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * Returns the count.
     * @return the requested result
     */
    public int getCount() {
        return workbook.getModel().getDefinedNames().size();
    }

    /**
     * Returns the requested item.
     * @param index index
     * @return the requested result
     */
    public DefinedName get(int index) {
        List<DefinedNameModel> names = workbook.getModel().getDefinedNames();
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= names.size()) {
            throw new CellsException("Defined name index was out of range.");
        }
        return new DefinedName(workbook, names.get(index));
    }

    /**
     * Adds a new item to the current collection.
     * @param name name
     * @param formula formula
     * @return the computed result
     */
    public int add(String name, String formula) {
        return add(name, formula, null);
    }

    /**
     * Adds a new item to the current collection.
     * @param name name
     * @param formula formula
     * @param localSheetIndex zero-based local sheet index
     * @return the computed result
     */
    public int add(String name, String formula, Integer localSheetIndex) {
        String normalizedName = DefinedNameUtility.normalizeName(name);
        String normalizedFormula = DefinedNameUtility.normalizeFormula(formula);
        workbook.ensureValidDefinedNameScope(localSheetIndex);
        workbook.ensureUniqueDefinedName(null, normalizedName, localSheetIndex);

        DefinedNameModel model = new DefinedNameModel();
        model.setName(normalizedName);
        model.setFormula(normalizedFormula);
        model.setLocalSheetIndex(localSheetIndex);

        workbook.getModel().getDefinedNames().add(model);
        return workbook.getModel().getDefinedNames().size() - 1;
    }

    /**
     * Removes at.
     * @param index index
     */
    public void removeAt(int index) {
        List<DefinedNameModel> names = workbook.getModel().getDefinedNames();
        // Handle the relevant branch before the state changes.
        if (index < 0 || index >= names.size()) {
            throw new CellsException("Defined name index was out of range.");
        }
        names.remove(index);
    }

    /**
     * Returns an iterator over the current collection.
     * @return the computed result
     */
    @Override
    public Iterator<DefinedName> iterator() {
        List<DefinedNameModel> names = workbook.getModel().getDefinedNames();
        List<DefinedName> result = new ArrayList<>(names.size());
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < names.size(); i++) {
            result.add(new DefinedName(workbook, names.get(i)));
        }
        return result.iterator();
    }
}