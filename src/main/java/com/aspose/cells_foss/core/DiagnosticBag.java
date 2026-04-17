package com.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Represents a bag of diagnostic entries.
 */
public final class DiagnosticBag {
    private final List<DiagnosticEntry> entries = new ArrayList<>();

    /**
     * Gets the read-only list of diagnostic entries.
     */
    public List<DiagnosticEntry> getEntries() {
        return Collections.unmodifiableList(entries);
    }

    /**
     * Adds a diagnostic entry to the bag.
     *
     * @param entry The entry to add.
     */
    public void add(DiagnosticEntry entry) {
        entries.add(entry);
    }
}