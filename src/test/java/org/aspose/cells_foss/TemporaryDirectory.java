package org.aspose.cells_foss;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.UUID;

/**
 * Creates a temporary directory under the project output folder for test artifacts.
 * Mirrors the C# TemporaryDirectory from shared test infrastructure.
 * Artifacts are preserved after the test so they can be inspected.
 */
public final class TemporaryDirectory implements AutoCloseable {

    private final Path rootPath;

    /**
     * Verifies that temporary directory.
     * @param suiteName name to use
     */
    public TemporaryDirectory(String suiteName) {
        Path outputDir = resolveOutputDir();
        this.rootPath = outputDir.resolve(suiteName).resolve(UUID.randomUUID().toString().replace("-", ""));
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            Files.createDirectories(rootPath);
        } catch (IOException e) {
            throw new RuntimeException("Failed to create temporary directory: " + rootPath, e);
        }
    }

    /**
     * Verifies that get root path.
     */
    public Path getRootPath() { return rootPath; }

    /**
     * Verifies that get path.
     * @param fileName name to use
     */
    public String getPath(String fileName) {
        return rootPath.resolve(fileName).toString();
    }

    /** Artifacts are kept for inspection; no cleanup on close. */
    @Override
    public void close() {}

    /**
     * Verifies that resolve output dir.
     */
    private static Path resolveOutputDir() {
        // Walk up the directory tree looking for pom.xml (project root)
        Path dir = Paths.get(System.getProperty("user.dir")).toAbsolutePath();
        while (dir != null) {
            if (Files.exists(dir.resolve("pom.xml"))) {
                return dir.resolve("output");
            }
            dir = dir.getParent();
        }
        return Paths.get(System.getProperty("java.io.tmpdir"), "cells-foss-tests");
    }
}

