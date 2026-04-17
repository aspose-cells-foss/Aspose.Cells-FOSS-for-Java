package com.aspose.cells_foss;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Represents diagnostics information during workbook loading.
 */
public final class LoadDiagnostics {
    private final List<LoadIssue> issues = new ArrayList<>();

    /**
     * Gets the read-only list of load issues.
     */
    public List<LoadIssue> getIssues() {
        return Collections.unmodifiableList(issues);
    }

    /**
     * Gets a value indicating whether any repairs were applied during loading.
     */
    public boolean hasRepairs() {
        // Walk the current collection so every entry is processed consistently.
        for (LoadIssue issue : issues) {
            if (issue.getRepairApplied()) {
                return true;
            }
        }
        return false;
    }

    /**
     * Gets a value indicating whether any issues have data loss risk.
     */
    public boolean hasDataLossRisk() {
        // Walk the current collection so every entry is processed consistently.
        for (LoadIssue issue : issues) {
            if (issue.getDataLossRisk()) {
                return true;
            }
        }
        return false;
    }

    /**
     * Adds a load issue to the diagnostics.
     * Package-private method (C# internal).
     *
     * @param issue the issue to add
     */
    void add(LoadIssue issue) {
        issues.add(issue);
    }
}