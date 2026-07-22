package org.aspose.cells_foss.packaging;

/**
 * Represents a descriptor for a package part in the XLSX package.
 */
public final class PackagePartDescriptor {
    private String partUri = "";
    private String contentType = "";
    private String category = "";

    /**
     * Returns the part uri.
     * @return the requested result
     */
    public String getPartUri() { return partUri; }
    /**
     * Sets the part uri.
     * @param partUri part uri
     */
    public void setPartUri(String partUri) { this.partUri = partUri; }

    /**
     * Returns the content type.
     * @return the requested result
     */
    public String getContentType() { return contentType; }
    /**
     * Sets the content type.
     * @param contentType content type
     */
    public void setContentType(String contentType) { this.contentType = contentType; }

    /**
     * Returns the category.
     * @return the requested result
     */
    public String getCategory() { return category; }
    /**
     * Sets the category.
     * @param category category
     */
    public void setCategory(String category) { this.category = category; }
}
