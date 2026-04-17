package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;
import javax.xml.parsers.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.zip.*;

/**
 * Package-private helpers for reading/writing the XLSX zip archive.
 */
final class XlsxWorkbookArchiveHelpers {

    /**
     * Initializes a new XlsxWorkbookArchiveHelpers instance.
     */
    private XlsxWorkbookArchiveHelpers() {}

    /**
     * Writes the current content to the target output.
     * @param zip zip
     * @param name name
     * @param data data
     */
    static void write(ZipOutputStream zip, String name, byte[] data) throws IOException {
        ZipEntry entry = new ZipEntry(name);
        zip.putNextEntry(entry);
        zip.write(data);
        zip.closeEntry();
    }

    /**
     * Parses the requested content.
     * @param bytes bytes
     * @return the computed result
     */
    static Document parse(byte[] bytes) {
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            dbf.setNamespaceAware(true);
            dbf.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            dbf.setFeature("http://xml.org/sax/features/external-general-entities", false);
            dbf.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            dbf.setXIncludeAware(false);
            return dbf.newDocumentBuilder().parse(new ByteArrayInputStream(bytes));
        } catch (Exception e) {
            throw new InvalidFileFormatException("Failed to parse XML: " + e.getMessage());
        }
    }

    /**
     * Loads the rels.
     * @param entries entries
     * @param path path to use
     * @return the requested result
     */
    static Map<String, String> loadRels(Map<String, byte[]> entries, String path) {
        byte[] bytes = entries.get(path);
        // Handle the relevant branch before the state changes.
        if (bytes == null) return new HashMap<>();
        Document doc = parse(bytes);
        Map<String, String> rels = new HashMap<>();
        NodeList nodes = doc.getElementsByTagNameNS("*", "Relationship");
        for (int i = 0; i < nodes.getLength(); i++) {
            Element el = (Element) nodes.item(i);
            rels.put(el.getAttribute("Id"), el.getAttribute("Target"));
        }
        return rels;
    }

    /**
     * Processes content types xml.
     * @param sheetCount sheet count
     * @param hasSst has sst
     * @param hasDocProps has doc props
     * @return the computed result
     */
    static byte[] contentTypesXml(int sheetCount, boolean hasSst, boolean hasDocProps) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        sb.append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        // Handle the relevant branch before the state changes.
        if (hasSst) sb.append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
        for (int i = 1; i <= sheetCount; i++)
            sb.append("<Override PartName=\"/xl/worksheets/sheet").append(i)
              .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        if (hasDocProps) {
            sb.append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            sb.append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        }
        sb.append("</Types>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    /**
     * Processes package rels xml.
     * @param hasDocProps has doc props
     * @return the computed result
     */
    static byte[] packageRelsXml(boolean hasDocProps) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        sb.append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>");
        // Handle the relevant branch before the state changes.
        if (hasDocProps) {
            sb.append("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>");
            sb.append("<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>");
        }
        sb.append("</Relationships>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    /**
     * Processes workbook rels xml.
     * @param sheetCount sheet count
     * @param hasSst has sst
     * @return the computed result
     */
    static byte[] workbookRelsXml(int sheetCount, boolean hasSst) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 1; i <= sheetCount; i++)
            sb.append("<Relationship Id=\"rId").append(i)
              .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet")
              .append(i).append(".xml\"/>");
        int nextId = sheetCount + 1;
        if (hasSst)
            sb.append("<Relationship Id=\"rId").append(nextId++)
              .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
        sb.append("<Relationship Id=\"rId").append(nextId)
          .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        sb.append("</Relationships>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    /**
     * Processes shared strings xml.
     * @param sst sst
     * @return the computed result
     */
    static byte[] sharedStringsXml(SharedStringRepository sst) {
        List<String> vals = sst.getValues();
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"")
          .append(vals.size()).append("\" uniqueCount=\"").append(vals.size()).append("\">");
        // Walk the current collection so every entry is processed consistently.
        for (String v : vals) sb.append("<si><t>").append(XlsxWorkbookSerializerCommon.xmlText(v)).append("</t></si>");
        sb.append("</sst>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }
}
