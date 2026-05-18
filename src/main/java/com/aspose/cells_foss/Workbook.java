package com.aspose.cells_foss;

import com.aspose.cells_foss.core.WorkbookModel;
import com.aspose.cells_foss.core.WorksheetModel;
import com.aspose.cells_foss.core.DefinedNameModel;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.IOException;

/**
 * Represents an Excel workbook.
 */
public class Workbook implements AutoCloseable {
    private final WorkbookModel model;
    private final WorksheetCollection worksheets;
    private final WorkbookSettings settings;
    private final WorkbookProperties properties;
    private final DocumentProperties documentProperties;
    private final DefinedNameCollection definedNames;
    private final LoadDiagnostics loadDiagnostics;
    private final CalculationProperties calculationProperties;

    /**
     * Initializes a new Workbook instance.
     */
    public Workbook() {
        this.model = new WorkbookModel();
        this.worksheets = new WorksheetCollection(this);
        this.settings = new WorkbookSettings(model.getSettings());
        this.properties = new WorkbookProperties(model.getProperties(), model);
        this.documentProperties = new DocumentProperties(model.getDocumentProperties());
        this.definedNames = new DefinedNameCollection(this);
        this.loadDiagnostics = new LoadDiagnostics();
        this.calculationProperties = this.properties.getCalculation();
    }

    /**
     * Initializes a new Workbook instance.
     * @param fileName name to use
     */
    public Workbook(String fileName) {
        this(fileName, new LoadOptions());
    }

    /**
     * Initializes a new Workbook instance.
     * @param stream stream to apply
     */
    public Workbook(InputStream stream) {
        this(stream, new LoadOptions());
    }

    /**
     * Initializes a new Workbook instance.
     * @param fileName name to use
     * @param options options to apply
     */
    public Workbook(String fileName, LoadOptions options) {
        this();
        // Validate the caller input before continuing.
        if (fileName == null) throw new IllegalArgumentException("fileName");
        if (options == null) throw new IllegalArgumentException("options");
        try (InputStream stream = new FileInputStream(fileName)) {
            loadFromStream(stream, options);
        } catch (IOException e) {
            throw new WorkbookLoadException("Failed to open file.", e);
        }
    }

    /**
     * Initializes a new Workbook instance.
     * @param stream stream to apply
     * @param options options to apply
     */
    public Workbook(InputStream stream, LoadOptions options) {
        this();
        // Validate the caller input before continuing.
        if (stream == null) throw new IllegalArgumentException("stream");
        if (options == null) throw new IllegalArgumentException("options");
        loadFromStream(stream, options);
    }

    /**
     * Returns the worksheets.
     * @return the requested result
     */
    public WorksheetCollection getWorksheets() { return worksheets; }
    /**
     * Returns the settings.
     * @return the requested result
     */
    public WorkbookSettings getSettings() { return settings; }
    /**
     * Returns the properties.
     * @return the requested result
     */
    public WorkbookProperties getProperties() { return properties; }
    /**
     * Returns the document properties.
     * @return the requested result
     */
    public DocumentProperties getDocumentProperties() { return documentProperties; }
    /**
     * Returns the defined names.
     * @return the requested result
     */
    public DefinedNameCollection getDefinedNames() { return definedNames; }
    /**
     * Returns the load diagnostics.
     * @return the requested result
     */
    public LoadDiagnostics getLoadDiagnostics() { return loadDiagnostics; }

    /** Returns the workbook calculation properties. */
    public CalculationProperties getCalculationProperties() { return calculationProperties; }

    /**
     * Returns the model.
     * @return the requested result
     */
    WorkbookModel getModel() { return model; }

    /**
     * Ensures that unique sheet name.
     * @param sheetName name to use
     * @param currentSheet current sheet
     */
    void ensureUniqueSheetName(String sheetName, WorksheetModel currentSheet) {
        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < model.getWorksheets().size(); index++) {
            WorksheetModel existing = model.getWorksheets().get(index);
            if (existing == currentSheet) continue;
            if (existing.getName().equalsIgnoreCase(sheetName)) {
                throw new CellsException("Worksheet name '" + sheetName + "' already exists.");
            }
        }
    }

    /**
     * Ensures that valid defined name scope.
     * @param localSheetIndex zero-based local sheet index
     */
    void ensureValidDefinedNameScope(Integer localSheetIndex) {
        // Handle the relevant branch before the state changes.
        if (localSheetIndex == null) return;
        if (localSheetIndex < 0 || localSheetIndex >= model.getWorksheets().size()) {
            throw new CellsException("Defined name scope must refer to an existing worksheet.");
        }
    }

    /**
     * Ensures that unique defined name.
     * @param currentDefinedName name to use
     * @param name name
     * @param localSheetIndex zero-based local sheet index
     */
    void ensureUniqueDefinedName(DefinedNameModel currentDefinedName, String name, Integer localSheetIndex) {
        // Walk the current collection so every entry is processed consistently.
        for (int index = 0; index < model.getDefinedNames().size(); index++) {
            DefinedNameModel existing = model.getDefinedNames().get(index);
            if (existing == currentDefinedName) continue;
            if (!existing.getName().equalsIgnoreCase(name)) continue;
            if (DefinedNameUtility.sameScope(existing.getLocalSheetIndex(), localSheetIndex)) {
                throw new CellsException("Defined name '" + name + "' already exists in the same scope.");
            }
        }
    }

    /** Saves the workbook to a file path (XLSX format). */
    public void save(String fileName) {
        save(fileName, new SaveOptions());
    }

    /** Saves the workbook to a file path with the given format. */
    public void save(String fileName, SaveFormat format) {
        SaveOptions options = new SaveOptions();
        options.setSaveFormat(format);
        save(fileName, options);
    }

    /** Saves the workbook to a file path with the given options. */
    public void save(String fileName, SaveOptions options) {
        // Validate the caller input before continuing.
        if (fileName == null) throw new IllegalArgumentException("fileName");
        if (options == null) throw new IllegalArgumentException("options");
        try (OutputStream stream = new FileOutputStream(fileName)) {
            save(stream, options);
        } catch (IOException e) {
            throw new WorkbookSaveException("Failed to create file.", e);
        }
    }

    /** Saves the workbook to a stream with the given format. */
    public void save(OutputStream stream, SaveFormat format) {
        SaveOptions options = new SaveOptions();
        options.setSaveFormat(format);
        save(stream, options);
    }

    /** Saves the workbook to a stream with the given options. */
    public void save(OutputStream stream, SaveOptions options) {
        // Validate the caller input before continuing.
        if (stream == null) throw new IllegalArgumentException("stream");
        if (options == null) throw new IllegalArgumentException("options");
        try {
            XlsxWorkbookSerializer.save(model, stream, options);
        } catch (CellsException e) {
            throw e;
        } catch (Exception e) {
            throw new WorkbookSaveException("Failed to save XLSX workbook.", e);
        }
    }

    /**
     * Releases any resources associated with this instance.
     */
    @Override
    public void close() {}

    /**
     * Loads the from stream.
     * @param stream stream to apply
     * @param options options to apply
     */
    private void loadFromStream(InputStream stream, LoadOptions options) {
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            WorkbookModel loadedModel = XlsxWorkbookSerializer.load(stream, options, loadDiagnostics);
            model.getWorksheets().clear();
            model.getWorksheets().addAll(loadedModel.getWorksheets());
            model.getSettings().setDateSystem(loadedModel.getSettings().getDateSystem());
            model.getDefinedNames().clear();
            model.getDefinedNames().addAll(loadedModel.getDefinedNames());
            model.setActiveSheetIndex(loadedModel.getActiveSheetIndex());
            model.getProperties().copyFrom(loadedModel.getProperties());
            model.getDocumentProperties().copyFrom(loadedModel.getDocumentProperties());
            model.setRawThemeXml(loadedModel.getRawThemeXml());
            model.setRawDefaultFontXml(loadedModel.getRawDefaultFontXml());
            model.getExternalLinkXmls().clear();
            model.getExternalLinkXmls().addAll(loadedModel.getExternalLinkXmls());
            model.getExternalLinkRels().clear();
            model.getExternalLinkRels().addAll(loadedModel.getExternalLinkRels());
        } catch (CellsException e) {
            throw e;
        } catch (Exception e) {
            throw new WorkbookLoadException("Failed to load XLSX workbook.", e);
        }
    }
}
