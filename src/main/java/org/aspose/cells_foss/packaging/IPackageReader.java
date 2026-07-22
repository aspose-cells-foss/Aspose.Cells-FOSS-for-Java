package org.aspose.cells_foss.packaging;

import java.io.InputStream;

/**
 * Provides a reader interface for reading package models from streams.
 */
public interface IPackageReader {
    /**
     * Reads a package model from the specified stream.
     *
     * @param stream the input stream to read from
     * @return the package model
     */
    PackageModel read(InputStream stream);
}
