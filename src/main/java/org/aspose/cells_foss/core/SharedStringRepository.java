package org.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * A repository that manages shared strings for Excel files.
 */
public final class SharedStringRepository {
    private final Map<String, Integer> indices = new HashMap<>();
    private final List<String> values = new ArrayList<>();

    /**
     * Gets the list of interned string values.
     */
    public List<String> getValues() {
        return values;
    }

    /**
     * Clears all stored strings and indices.
     */
    public void clear() {
        indices.clear();
        values.clear();
    }

    /**
     * Tries to get the string value at the specified index.
     *
     * @param index the zero-based index
     * @param value single-element array used as an output holder; value[0] is set on return
     * @return true if the index is valid; false otherwise
     */
    public boolean tryGetValue(int index, String[] value) {
        // Handle the relevant branch before the state changes.
        if (index >= 0 && index < values.size()) {
            value[0] = values.get(index);
            return true;
        }
        value[0] = "";
        return false;
    }

    /**
     * Interns a string, adding it to the repository if not already present.
     *
     * @param value the string to intern
     * @return the zero-based index of the string in the repository
     */
    public int intern(String value) {
        Integer index = indices.get(value);
        // Handle the relevant branch before the state changes.
        if (index != null) {
            return index;
        }
        index = values.size();
        values.add(value);
        indices.put(value, index);
        return index;
    }
}
