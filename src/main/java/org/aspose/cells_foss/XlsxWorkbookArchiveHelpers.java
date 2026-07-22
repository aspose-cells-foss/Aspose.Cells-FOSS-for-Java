package org.aspose.cells_foss;

import org.aspose.cells_foss.core.*;
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
        return contentTypesXml(sheetCount, hasSst, hasDocProps, false, false,
                java.util.Collections.emptySet(), java.util.Collections.emptyMap());
    }

    static byte[] contentTypesXml(int sheetCount, boolean hasSst, boolean hasDocProps,
                                   boolean hasDrawings, boolean hasCharts,
                                   java.util.Set<String> chartParts,
                                   java.util.Map<String, byte[]> mediaParts) {
        return contentTypesXml(sheetCount, hasSst, hasDocProps, hasDrawings, hasCharts, chartParts, mediaParts, 0, false, false);
    }

    static byte[] contentTypesXml(int sheetCount, boolean hasSst, boolean hasDocProps,
                                   boolean hasDrawings, boolean hasCharts,
                                   java.util.Set<String> chartParts,
                                   java.util.Map<String, byte[]> mediaParts,
                                   int drawingCount) {
        return contentTypesXml(sheetCount, hasSst, hasDocProps, hasDrawings, hasCharts, chartParts, mediaParts, drawingCount, false, false);
    }

    static byte[] contentTypesXml(int sheetCount, boolean hasSst, boolean hasDocProps,
                                   boolean hasDrawings, boolean hasCharts,
                                   java.util.Set<String> chartParts,
                                   java.util.Map<String, byte[]> mediaParts,
                                   int drawingCount, boolean hasVml) {
        return contentTypesXml(sheetCount, hasSst, hasDocProps, hasDrawings, hasCharts, chartParts, mediaParts, drawingCount, hasVml, false, 0);
    }

    static byte[] contentTypesXml(int sheetCount, boolean hasSst, boolean hasDocProps,
                                   boolean hasDrawings, boolean hasCharts,
                                   java.util.Set<String> chartParts,
                                   java.util.Map<String, byte[]> mediaParts,
                                   int drawingCount, boolean hasVml, boolean hasTheme) {
        return contentTypesXml(sheetCount, hasSst, hasDocProps, hasDrawings, hasCharts, chartParts, mediaParts, drawingCount, hasVml, hasTheme, 0);
    }

    static byte[] contentTypesXml(int sheetCount, boolean hasSst, boolean hasDocProps,
                                   boolean hasDrawings, boolean hasCharts,
                                   java.util.Set<String> chartParts,
                                   java.util.Map<String, byte[]> mediaParts,
                                   int drawingCount, boolean hasVml, boolean hasTheme,
                                   int externalLinkCount) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        if (hasVml)
            sb.append("<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>");

        // Add image type defaults for all distinct extensions
        java.util.Set<String> seenExts = new java.util.HashSet<>();
        for (String mediaPath : mediaParts.keySet()) {
            int dot = mediaPath.lastIndexOf('.');
            if (dot >= 0) {
                String ext = mediaPath.substring(dot + 1).toLowerCase();
                if (seenExts.add(ext)) {
                    String ct = PictureCollection.contentTypeFromExtension(ext);
                    sb.append("<Default Extension=\"").append(ext).append("\" ContentType=\"").append(ct).append("\"/>");
                }
            }
        }

        sb.append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        sb.append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        if (hasTheme) sb.append("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
        if (hasSst) sb.append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
        for (int i = 1; i <= sheetCount; i++)
            sb.append("<Override PartName=\"/xl/worksheets/sheet").append(i)
              .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        if (hasDocProps) {
            sb.append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            sb.append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        }
        if (hasDrawings) {
            for (int i = 1; i <= drawingCount; i++) {
                sb.append("<Override PartName=\"/xl/drawings/drawing").append(i)
                  .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
            }
            for (String chartPart : chartParts) {
                if (chartPart.contains("/_rels/")) continue; // rels covered by Default
                String filename = chartPart.substring(chartPart.lastIndexOf('/') + 1);
                sb.append("<Override PartName=\"/").append(chartPart)
                  .append("\" ContentType=\"").append(chartPartContentType(filename)).append("\"/>");
            }
        }
        for (int i = 1; i <= externalLinkCount; i++)
            sb.append("<Override PartName=\"/xl/externalLinks/externalLink").append(i)
              .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml\"/>");
        sb.append("</Types>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    /** Inserts table Override entries into an already-built content-types byte array. */
    static byte[] addTableContentTypes(byte[] contentTypesBytes, int tableCount) {
        if (tableCount == 0) return contentTypesBytes;
        String xml = new String(contentTypesBytes, StandardCharsets.UTF_8);
        StringBuilder inserts = new StringBuilder();
        for (int i = 1; i <= tableCount; i++) {
            inserts.append("<Override PartName=\"/xl/tables/table").append(i)
              .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
        }
        xml = xml.replace("</Types>", inserts + "</Types>");
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    /**
     * Inserts comment part Override entries into an already-built content-types byte array.
     * Without these the parts fall back to the {@code xml} Default (application/xml), which
     * strict consumers reject when opening the package.
     */
    static byte[] addCommentContentTypes(byte[] contentTypesBytes, java.util.Collection<Integer> sheetNumbers) {
        if (sheetNumbers.isEmpty()) return contentTypesBytes;
        String xml = new String(contentTypesBytes, StandardCharsets.UTF_8);
        StringBuilder inserts = new StringBuilder();
        for (int sn : sheetNumbers) {
            inserts.append("<Override PartName=\"/xl/comments").append(sn)
              .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>");
        }
        xml = xml.replace("</Types>", inserts + "</Types>");
        return xml.getBytes(StandardCharsets.UTF_8);
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
        return workbookRelsXml(sheetCount, hasSst, false);
    }

    static byte[] workbookRelsXml(int sheetCount, boolean hasSst, boolean hasTheme) {
        return workbookRelsXml(sheetCount, hasSst, hasTheme, 0);
    }

    static byte[] workbookRelsXml(int sheetCount, boolean hasSst, boolean hasTheme, int externalLinkCount) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for (int i = 1; i <= sheetCount; i++)
            sb.append("<Relationship Id=\"rId").append(i)
              .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet")
              .append(i).append(".xml\"/>");
        int nextId = sheetCount + 1;
        if (hasSst)
            sb.append("<Relationship Id=\"rId").append(nextId++)
              .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
        if (hasTheme)
            sb.append("<Relationship Id=\"rId").append(nextId++)
              .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>");
        sb.append("<Relationship Id=\"rId").append(nextId++)
          .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        for (int i = 1; i <= externalLinkCount; i++)
            sb.append("<Relationship Id=\"rIdExt").append(i)
              .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink\" Target=\"externalLinks/externalLink")
              .append(i).append(".xml\"/>");
        sb.append("</Relationships>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    static byte[] defaultThemeXml() {
        return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            + "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">"
            + "<a:themeElements>"
            + "<a:clrScheme name=\"Office\">"
            + "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>"
            + "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>"
            + "<a:dk2><a:srgbClr val=\"0E2841\"/></a:dk2>"
            + "<a:lt2><a:srgbClr val=\"E8E8E8\"/></a:lt2>"
            + "<a:accent1><a:srgbClr val=\"156082\"/></a:accent1>"
            + "<a:accent2><a:srgbClr val=\"E97132\"/></a:accent2>"
            + "<a:accent3><a:srgbClr val=\"196B24\"/></a:accent3>"
            + "<a:accent4><a:srgbClr val=\"0F9ED5\"/></a:accent4>"
            + "<a:accent5><a:srgbClr val=\"A02B93\"/></a:accent5>"
            + "<a:accent6><a:srgbClr val=\"4EA72E\"/></a:accent6>"
            + "<a:hlink><a:srgbClr val=\"467886\"/></a:hlink>"
            + "<a:folHlink><a:srgbClr val=\"96607D\"/></a:folHlink>"
            + "</a:clrScheme>"
            + "<a:fontScheme name=\"Office\">"
            + "<a:majorFont><a:latin typeface=\"Aptos Display\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>"
            + "<a:minorFont><a:latin typeface=\"Aptos Narrow\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>"
            + "</a:fontScheme>"
            + "<a:fmtScheme name=\"Office\">"
            + "<a:fillStyleLst>"
            + "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
            + "<a:gradFill rotWithShape=\"1\"><a:gsLst>"
            + "<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs>"
            + "</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>"
            + "<a:gradFill rotWithShape=\"1\"><a:gsLst>"
            + "<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs>"
            + "</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>"
            + "</a:fillStyleLst>"
            + "<a:lnStyleLst>"
            + "<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
            + "<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
            + "<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>"
            + "</a:lnStyleLst>"
            + "<a:effectStyleLst>"
            + "<a:effectStyle><a:effectLst/></a:effectStyle>"
            + "<a:effectStyle><a:effectLst/></a:effectStyle>"
            + "<a:effectStyle><a:effectLst>"
            + "<a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\">"
            + "<a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr>"
            + "</a:outerShdw></a:effectLst></a:effectStyle>"
            + "</a:effectStyleLst>"
            + "<a:bgFillStyleLst>"
            + "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
            + "<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill>"
            + "<a:gradFill rotWithShape=\"1\"><a:gsLst>"
            + "<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs>"
            + "</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>"
            + "</a:bgFillStyleLst>"
            + "</a:fmtScheme>"
            + "</a:themeElements>"
            + "</a:theme>"
        ).getBytes(StandardCharsets.UTF_8);
    }

    private static String chartPartContentType(String filename) {
        if (filename.startsWith("chartEx"))    return "application/vnd.ms-office.chartex+xml";
        if (filename.startsWith("style"))      return "application/vnd.ms-office.chartstyle+xml";
        if (filename.startsWith("colors"))     return "application/vnd.ms-office.chartcolorstyle+xml";
        if (filename.startsWith("userShapes")) return "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
        return "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
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

