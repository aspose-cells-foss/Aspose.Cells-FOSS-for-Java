package com.aspose.cells_foss.packaging;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Represents the model of a package (e.g., XLSX file structure) with parts,
 * relationships, and unsupported parts.
 */
public final class PackageModel {
    private final List<PackagePartDescriptor> parts = new ArrayList<>();
    private final List<RelationshipDescriptor> relationships = new ArrayList<>();
    private final Map<String, byte[]> unsupportedParts = new HashMap<>();

    /**
     * Returns the parts.
     * @return the requested result
     */
    public List<PackagePartDescriptor> getParts() {
        return parts;
    }

    /**
     * Returns the relationships.
     * @return the requested result
     */
    public List<RelationshipDescriptor> getRelationships() {
        return relationships;
    }

    /**
     * Returns the unsupported parts.
     * @return the requested result
     */
    public Map<String, byte[]> getUnsupportedParts() {
        return unsupportedParts;
    }
}