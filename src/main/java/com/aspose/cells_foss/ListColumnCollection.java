package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ListColumnModel;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;

/** Ordered collection of columns in an Excel table. */
public final class ListColumnCollection implements Iterable<ListColumn> {
    private final List<ListColumnModel> models;

    ListColumnCollection(List<ListColumnModel> models) {
        this.models = models;
    }

    public int getCount() { return models.size(); }

    public ListColumn get(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("Column index out of range.");
        return new ListColumn(models.get(index));
    }

    @Override
    public Iterator<ListColumn> iterator() {
        return new Iterator<>() {
            private int i = 0;
            @Override public boolean hasNext() { return i < models.size(); }
            @Override public ListColumn next() {
                if (!hasNext()) throw new NoSuchElementException();
                return new ListColumn(models.get(i++));
            }
        };
    }
}
