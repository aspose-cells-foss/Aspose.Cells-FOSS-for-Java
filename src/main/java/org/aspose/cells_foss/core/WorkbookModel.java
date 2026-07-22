package org.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.List;

/**
 * Represents the top-level model of a workbook.
 */
public final class WorkbookModel {
    private final List<WorksheetModel> worksheets = new ArrayList<>();
    private final WorkbookSettingsModel settings = new WorkbookSettingsModel();
    private final WorkbookPropertiesModel properties = new WorkbookPropertiesModel();
    private final DocumentPropertiesModel documentProperties = new DocumentPropertiesModel();
    private final DiagnosticBag diagnostics = new DiagnosticBag();
    private final StyleRepository styles = new StyleRepository();
    private final SharedStringRepository sharedStrings = new SharedStringRepository();
    private StyleValue defaultStyle = StyleValue.getDefault().clone();
    private int activeSheetIndex = 0;
    private final List<DefinedNameModel> definedNames = new ArrayList<>();
    /** Raw bytes of xl/theme/theme1.xml from the source file, preserved for round-trip fidelity. */
    private byte[] rawThemeXml;
    /** Raw XML of the default &lt;font&gt; element (fonts[0]) from the source styles.xml. */
    private String rawDefaultFontXml;
    /** Verbatim bytes of xl/externalLinks/externalLink{N}.xml files, in order. */
    private final List<byte[]> externalLinkXmls = new ArrayList<>();
    /** Verbatim bytes of the corresponding xl/externalLinks/_rels/externalLink{N}.xml.rels files (null if absent). */
    private final List<byte[]> externalLinkRels = new ArrayList<>();

    /**
     * Initializes a new WorkbookModel instance.
     */
    public WorkbookModel() {
        worksheets.add(new WorksheetModel("Sheet1"));
    }

    /**
     * Returns the worksheets.
     * @return the requested result
     */
    public List<WorksheetModel> getWorksheets() {
        return worksheets;
    }

    /**
     * Returns the settings.
     * @return the requested result
     */
    public WorkbookSettingsModel getSettings() {
        return settings;
    }

    /**
     * Returns the properties.
     * @return the requested result
     */
    public WorkbookPropertiesModel getProperties() {
        return properties;
    }

    /**
     * Returns the document properties.
     * @return the requested result
     */
    public DocumentPropertiesModel getDocumentProperties() {
        return documentProperties;
    }

    /**
     * Returns the diagnostics.
     * @return the requested result
     */
    public DiagnosticBag getDiagnostics() {
        return diagnostics;
    }

    /**
     * Returns the styles.
     * @return the requested result
     */
    public StyleRepository getStyles() {
        return styles;
    }

    /**
     * Returns the shared strings.
     * @return the requested result
     */
    public SharedStringRepository getSharedStrings() {
        return sharedStrings;
    }

    /**
     * Returns the default style.
     * @return the requested result
     */
    public StyleValue getDefaultStyle() {
        return defaultStyle;
    }

    /**
     * Sets the default style.
     * @param defaultStyle default style
     */
    public void setDefaultStyle(StyleValue defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    /**
     * Returns the active sheet index.
     * @return the requested result
     */
    public int getActiveSheetIndex() {
        return activeSheetIndex;
    }

    /**
     * Sets the active sheet index.
     * @param activeSheetIndex zero-based active sheet index
     */
    public void setActiveSheetIndex(int activeSheetIndex) {
        this.activeSheetIndex = activeSheetIndex;
    }

    /**
     * Returns the defined names.
     * @return the requested result
     */
    public List<DefinedNameModel> getDefinedNames() {
        return definedNames;
    }

    public byte[] getRawThemeXml() { return rawThemeXml; }
    public void setRawThemeXml(byte[] rawThemeXml) { this.rawThemeXml = rawThemeXml; }

    public String getRawDefaultFontXml() { return rawDefaultFontXml; }
    public void setRawDefaultFontXml(String rawDefaultFontXml) { this.rawDefaultFontXml = rawDefaultFontXml; }

    /** Returns the external link XML files (in index order, matching externalLink1, externalLink2, 鈥?. */
    public List<byte[]> getExternalLinkXmls() { return externalLinkXmls; }
    /** Returns the external link rels files (parallel to externalLinkXmls; element may be null). */
    public List<byte[]> getExternalLinkRels() { return externalLinkRels; }
}
