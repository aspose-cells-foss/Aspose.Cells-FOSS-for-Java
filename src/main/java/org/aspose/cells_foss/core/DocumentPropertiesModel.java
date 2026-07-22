package org.aspose.cells_foss.core;

import java.time.LocalDateTime;

/**
 * Represents the document properties model for an Excel file.
 */
public final class DocumentPropertiesModel {
    private final CoreDocumentPropertiesModel core;
    private final ExtendedDocumentPropertiesModel extended;

    /**
     * Initializes a new DocumentPropertiesModel instance.
     */
    public DocumentPropertiesModel() {
        this.core = new CoreDocumentPropertiesModel();
        this.extended = new ExtendedDocumentPropertiesModel();
    }

    /**
     * Returns the core.
     * @return the requested result
     */
    public CoreDocumentPropertiesModel getCore() { return core; }

    /**
     * Returns the extended.
     * @return the requested result
     */
    public ExtendedDocumentPropertiesModel getExtended() { return extended; }

    /**
     * Copies the from.
     * @param source source
     */
    public void copyFrom(DocumentPropertiesModel source) {
        core.copyFrom(source.getCore());
        extended.copyFrom(source.getExtended());
    }

    /**
     * Indicates whether this instance has stored state.
     * @return true when the condition is satisfied
     */
    public boolean hasStoredState() {
        return core.hasStoredState() || extended.hasStoredState();
    }

    // Inner class: CoreDocumentPropertiesModel
    /**
     * Represents the CoreDocumentPropertiesModel component.
     */
    public static final class CoreDocumentPropertiesModel {
        private String title = "";
        private String subject = "";
        private String creator = "";
        private String keywords = "";
        private String description = "";
        private String lastModifiedBy = "";
        private String revision = "";
        private String category = "";
        private String contentStatus = "";
        private LocalDateTime created;
        private LocalDateTime modified;

        /**
         * Returns the title.
         * @return the requested result
         */
        public String getTitle() { return title; }
        /**
         * Sets the title.
         * @param value value to apply
         */
        public void setTitle(String value) { this.title = value; }

        /**
         * Returns the subject.
         * @return the requested result
         */
        public String getSubject() { return subject; }
        /**
         * Sets the subject.
         * @param value value to apply
         */
        public void setSubject(String value) { this.subject = value; }

        /**
         * Returns the creator.
         * @return the requested result
         */
        public String getCreator() { return creator; }
        /**
         * Sets the creator.
         * @param value value to apply
         */
        public void setCreator(String value) { this.creator = value; }

        /**
         * Returns the keywords.
         * @return the requested result
         */
        public String getKeywords() { return keywords; }
        /**
         * Sets the keywords.
         * @param value value to apply
         */
        public void setKeywords(String value) { this.keywords = value; }

        /**
         * Returns the description.
         * @return the requested result
         */
        public String getDescription() { return description; }
        /**
         * Sets the description.
         * @param value value to apply
         */
        public void setDescription(String value) { this.description = value; }

        /**
         * Returns the last modified by.
         * @return the requested result
         */
        public String getLastModifiedBy() { return lastModifiedBy; }
        /**
         * Sets the last modified by.
         * @param value value to apply
         */
        public void setLastModifiedBy(String value) { this.lastModifiedBy = value; }

        /**
         * Returns the revision.
         * @return the requested result
         */
        public String getRevision() { return revision; }
        /**
         * Sets the revision.
         * @param value value to apply
         */
        public void setRevision(String value) { this.revision = value; }

        /**
         * Returns the category.
         * @return the requested result
         */
        public String getCategory() { return category; }
        /**
         * Sets the category.
         * @param value value to apply
         */
        public void setCategory(String value) { this.category = value; }

        /**
         * Returns the content status.
         * @return the requested result
         */
        public String getContentStatus() { return contentStatus; }
        /**
         * Sets the content status.
         * @param value value to apply
         */
        public void setContentStatus(String value) { this.contentStatus = value; }

        /**
         * Returns the created.
         * @return the requested result
         */
        public LocalDateTime getCreated() { return created; }
        /**
         * Sets the created.
         * @param value value to apply
         */
        public void setCreated(LocalDateTime value) { this.created = value; }

        /**
         * Returns the modified.
         * @return the requested result
         */
        public LocalDateTime getModified() { return modified; }
        /**
         * Sets the modified.
         * @param value value to apply
         */
        public void setModified(LocalDateTime value) { this.modified = value; }

        /**
         * Copies the from.
         * @param source source
         */
        public void copyFrom(CoreDocumentPropertiesModel source) {
            title = source.title;
            subject = source.subject;
            creator = source.creator;
            keywords = source.keywords;
            description = source.description;
            lastModifiedBy = source.lastModifiedBy;
            revision = source.revision;
            category = source.category;
            contentStatus = source.contentStatus;
            created = source.created;
            modified = source.modified;
        }

        /**
         * Indicates whether this instance has stored state.
         * @return true when the condition is satisfied
         */
        public boolean hasStoredState() {
            return !title.isEmpty()
                || !subject.isEmpty()
                || !creator.isEmpty()
                || !keywords.isEmpty()
                || !description.isEmpty()
                || !lastModifiedBy.isEmpty()
                || !revision.isEmpty()
                || !category.isEmpty()
                || !contentStatus.isEmpty()
                || created != null
                || modified != null;
        }
    }

    // Inner class: ExtendedDocumentPropertiesModel
    /**
     * Represents the ExtendedDocumentPropertiesModel component.
     */
    public static final class ExtendedDocumentPropertiesModel {
        private String application = "";
        private String appVersion = "";
        private String company = "";
        private String manager = "";
        private Integer docSecurity;
        private String hyperlinkBase = "";
        private Boolean scaleCrop;
        private Boolean linksUpToDate;
        private Boolean sharedDoc;

        /**
         * Returns the application.
         * @return the requested result
         */
        public String getApplication() { return application; }
        /**
         * Sets the application.
         * @param value value to apply
         */
        public void setApplication(String value) { this.application = value; }

        /**
         * Returns the app version.
         * @return the requested result
         */
        public String getAppVersion() { return appVersion; }
        /**
         * Sets the app version.
         * @param value value to apply
         */
        public void setAppVersion(String value) { this.appVersion = value; }

        /**
         * Returns the company.
         * @return the requested result
         */
        public String getCompany() { return company; }
        /**
         * Sets the company.
         * @param value value to apply
         */
        public void setCompany(String value) { this.company = value; }

        /**
         * Returns the manager.
         * @return the requested result
         */
        public String getManager() { return manager; }
        /**
         * Sets the manager.
         * @param value value to apply
         */
        public void setManager(String value) { this.manager = value; }

        /**
         * Returns the doc security.
         * @return the requested result
         */
        public Integer getDocSecurity() { return docSecurity; }
        /**
         * Sets the doc security.
         * @param value value to apply
         */
        public void setDocSecurity(Integer value) { this.docSecurity = value; }

        /**
         * Returns the hyperlink base.
         * @return the requested result
         */
        public String getHyperlinkBase() { return hyperlinkBase; }
        /**
         * Sets the hyperlink base.
         * @param value value to apply
         */
        public void setHyperlinkBase(String value) { this.hyperlinkBase = value; }

        /**
         * Returns the scale crop.
         * @return the requested result
         */
        public Boolean getScaleCrop() { return scaleCrop; }
        /**
         * Sets the scale crop.
         * @param value value to apply
         */
        public void setScaleCrop(Boolean value) { this.scaleCrop = value; }

        /**
         * Returns the links up to date.
         * @return the requested result
         */
        public Boolean getLinksUpToDate() { return linksUpToDate; }
        /**
         * Sets the links up to date.
         * @param value value to apply
         */
        public void setLinksUpToDate(Boolean value) { this.linksUpToDate = value; }

        /**
         * Returns the shared doc.
         * @return the requested result
         */
        public Boolean getSharedDoc() { return sharedDoc; }
        /**
         * Sets the shared doc.
         * @param value value to apply
         */
        public void setSharedDoc(Boolean value) { this.sharedDoc = value; }

        /**
         * Copies the from.
         * @param source source
         */
        public void copyFrom(ExtendedDocumentPropertiesModel source) {
            application = source.application;
            appVersion = source.appVersion;
            company = source.company;
            manager = source.manager;
            docSecurity = source.docSecurity;
            hyperlinkBase = source.hyperlinkBase;
            scaleCrop = source.scaleCrop;
            linksUpToDate = source.linksUpToDate;
            sharedDoc = source.sharedDoc;
        }

        /**
         * Indicates whether this instance has stored state.
         * @return true when the condition is satisfied
         */
        public boolean hasStoredState() {
            return !application.isEmpty()
                || !appVersion.isEmpty()
                || !company.isEmpty()
                || !manager.isEmpty()
                || docSecurity != null
                || !hyperlinkBase.isEmpty()
                || scaleCrop != null
                || linksUpToDate != null
                || sharedDoc != null;
        }
    }
}
