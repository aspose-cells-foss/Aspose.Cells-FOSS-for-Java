package com.aspose.cells_foss.packaging;

/**
 * Defines constants for Open XML package part paths and relationship types.
 */
public final class PackagingConventions {

    /** Path of the Content Types part in the package. */
    public static final String CONTENT_TYPES_PART = "/[Content_Types].xml";

    /** Path of the root relationships part in the package. */
    public static final String ROOT_RELATIONSHIPS_PART = "/_rels/.rels";

    /** Path of the main workbook part. */
    public static final String WORKBOOK_PART = "/xl/workbook.xml";

    /** Path of the workbook relationships part. */
    public static final String WORKBOOK_RELATIONSHIPS_PART = "/xl/_rels/workbook.xml.rels";

    /** Relationship type for the main office document. */
    public static final String OFFICE_DOCUMENT_RELATIONSHIP =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    /** Relationship type for a worksheet part. */
    public static final String WORKSHEET_RELATIONSHIP =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    /** Relationship type for the styles part. */
    public static final String STYLES_RELATIONSHIP =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    /** Relationship type for the shared strings part. */
    public static final String SHARED_STRINGS_RELATIONSHIP =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

    /**
     * Initializes a new PackagingConventions instance.
     */
    private PackagingConventions() {}
}