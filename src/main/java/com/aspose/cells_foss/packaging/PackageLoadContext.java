package com.aspose.cells_foss.packaging;

/**
 * Represents the context for loading a package.
 */
public final class PackageLoadContext {
    private final Object workbook;
    private final PackageModel model;

    /**
     * Initializes a new PackageLoadContext instance.
     * @param workbook workbook to apply
     * @param model model
     */
    public PackageLoadContext(Object workbook, PackageModel model) {
        this.workbook = workbook;
        this.model = model;
    }

    /**
     * Returns the workbook.
     * @return the requested result
     */
    public Object getWorkbook() { return workbook; }

    /**
     * Returns the package.
     * @return the requested result
     */
    public PackageModel getPackage() { return model; }
}