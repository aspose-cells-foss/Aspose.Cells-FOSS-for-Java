package org.aspose.cells_foss.packaging;

/**
 * Represents a relationship descriptor in the XLSX package.
 */
public final class RelationshipDescriptor {
    private String id = "";
    private String type = "";
    private String target = "";
    private boolean isExternal;

    /**
     * Returns the id.
     * @return the requested result
     */
    public String getId() { return id; }
    /**
     * Sets the id.
     * @param id id
     */
    public void setId(String id) { this.id = id; }

    /**
     * Returns the type.
     * @return the requested result
     */
    public String getType() { return type; }
    /**
     * Sets the type.
     * @param type type
     */
    public void setType(String type) { this.type = type; }

    /**
     * Returns the target.
     * @return the requested result
     */
    public String getTarget() { return target; }
    /**
     * Sets the target.
     * @param target target
     */
    public void setTarget(String target) { this.target = target; }

    /**
     * Returns the external.
     * @return the requested result
     */
    public boolean getIsExternal() { return isExternal; }
    /**
     * Sets the external.
     * @param isExternal is external
     */
    public void setIsExternal(boolean isExternal) { this.isExternal = isExternal; }
}
