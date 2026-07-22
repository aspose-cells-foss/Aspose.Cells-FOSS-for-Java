package org.aspose.cells_foss;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.zip.*;

/**
 * Helpers for reading and manipulating XLSX/ZIP packages in tests.
 * Uses Java's built-in zip support (replaces C# ZipFile / OpenXML SDK usage).
 * For structured XLSX inspection, use Apache POI (XSSFWorkbook) instead.
 */
public final class ZipPackageHelper {

    /**
     * Verifies that zip package helper.
     */
    private ZipPackageHelper() {}

    /** Reads a UTF-8 text entry from a ZIP file. */
    public static String readEntryText(String packagePath, String entryPath) {
        String normalized = normalizeEntryPath(entryPath);
        // Wrap lower-level failures in the library-specific exception flow.
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(packagePath), StandardCharsets.UTF_8)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.getName().equals(normalized)) {
                    return new String(zis.readAllBytes(), StandardCharsets.UTF_8);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException("Failed to read zip entry '" + entryPath + "' from '" + packagePath + "'.", e);
        }
        throw new RuntimeException("Missing zip entry '" + entryPath + "' in '" + packagePath + "'.");
    }

    /** Returns true if an entry exists inside the ZIP file. */
    public static boolean entryExists(String packagePath, String entryPath) {
        String normalized = normalizeEntryPath(entryPath);
        // Wrap lower-level failures in the library-specific exception flow.
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(packagePath), StandardCharsets.UTF_8)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.getName().equals(normalized)) return true;
            }
        } catch (IOException e) {
            throw new RuntimeException("Failed to inspect zip '" + packagePath + "'.", e);
        }
        return false;
    }

    /**
     * Verifies that normalize entry path.
     * @param entryPath path to use
     */
    private static String normalizeEntryPath(String entryPath) {
        return entryPath.replaceAll("^/+", "").replace('\\', '/');
    }
}

