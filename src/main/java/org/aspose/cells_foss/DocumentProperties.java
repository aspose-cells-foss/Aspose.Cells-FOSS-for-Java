package org.aspose.cells_foss;

import org.aspose.cells_foss.core.DocumentPropertiesModel;

/**
 * Represents the document properties of a workbook (docProps/core.xml and docProps/app.xml).
 *
 * <p>Core properties (dc: / cp: namespace in core.xml):
 * title, subject, author (dc:creator), keywords, comments (dc:description), category.
 *
 * <p>Extended properties (app.xml):
 * company, manager.
 */
public final class DocumentProperties {

    private final DocumentPropertiesModel model;

    /**
     * Initializes a new DocumentProperties instance.
     * @param model model
     */
    DocumentProperties(DocumentPropertiesModel model) {
        this.model = model;
    }

    // --- Core properties (docProps/core.xml) ---

    /**
     * Returns the title.
     * @return the requested result
     */
    public String getTitle() { return nvl(model.getCore().getTitle()); }
    /**
     * Sets the title.
     * @param value value to apply
     */
    public void setTitle(String value) { model.getCore().setTitle(nvl(value)); }

    /**
     * Returns the subject.
     * @return the requested result
     */
    public String getSubject() { return nvl(model.getCore().getSubject()); }
    /**
     * Sets the subject.
     * @param value value to apply
     */
    public void setSubject(String value) { model.getCore().setSubject(nvl(value)); }

    /** Maps to dc:creator in core.xml. */
    public String getAuthor() { return nvl(model.getCore().getCreator()); }
    /**
     * Sets the author.
     * @param value value to apply
     */
    public void setAuthor(String value) { model.getCore().setCreator(nvl(value)); }

    /**
     * Returns the keywords.
     * @return the requested result
     */
    public String getKeywords() { return nvl(model.getCore().getKeywords()); }
    /**
     * Sets the keywords.
     * @param value value to apply
     */
    public void setKeywords(String value) { model.getCore().setKeywords(nvl(value)); }

    /** Maps to dc:description in core.xml. */
    public String getComments() { return nvl(model.getCore().getDescription()); }
    /**
     * Sets the comments.
     * @param value value to apply
     */
    public void setComments(String value) { model.getCore().setDescription(nvl(value)); }

    /**
     * Returns the category.
     * @return the requested result
     */
    public String getCategory() { return nvl(model.getCore().getCategory()); }
    /**
     * Sets the category.
     * @param value value to apply
     */
    public void setCategory(String value) { model.getCore().setCategory(nvl(value)); }

    // --- Extended properties (docProps/app.xml) ---

    /**
     * Returns the company.
     * @return the requested result
     */
    public String getCompany() { return nvl(model.getExtended().getCompany()); }
    /**
     * Sets the company.
     * @param value value to apply
     */
    public void setCompany(String value) { model.getExtended().setCompany(nvl(value)); }

    /**
     * Returns the manager.
     * @return the requested result
     */
    public String getManager() { return nvl(model.getExtended().getManager()); }
    /**
     * Sets the manager.
     * @param value value to apply
     */
    public void setManager(String value) { model.getExtended().setManager(nvl(value)); }

    /**
     * Returns the fallback value when the primary value is empty.
     * @param s s
     * @return the computed result
     */
    private static String nvl(String s) { return s != null ? s : ""; }
}

