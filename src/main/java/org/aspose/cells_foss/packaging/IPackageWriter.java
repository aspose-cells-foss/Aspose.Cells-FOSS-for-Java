package org.aspose.cells_foss.packaging;

import java.io.OutputStream;
import java.io.IOException;

/**
 * Provides a contract for writing package models to a stream.
 */
public interface IPackageWriter {
    /**
     * Writes the package model to the specified stream.
     *
     * @param stream the output stream to write to
     * @param packageModel the package model to write
     * @throws IOException if an I/O error occurs
     */
    void write(OutputStream stream, PackageModel packageModel) throws IOException;
}
