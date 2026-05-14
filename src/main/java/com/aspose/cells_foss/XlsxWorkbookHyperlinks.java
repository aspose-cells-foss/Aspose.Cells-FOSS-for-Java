package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * Package-private helpers for building and loading XLSX hyperlinks.
 */
final class XlsxWorkbookHyperlinks {

    /**
     * Initializes a new XlsxWorkbookHyperlinks instance.
     */
    private XlsxWorkbookHyperlinks() {}

    // =========================================================================
    // Build
    // =========================================================================

    /**
     * Builds the hyperlinks section.
     * @param ws ws
     * @param sb sb
     * @param externalRels external rels
     */
    static void buildHyperlinksSection(WorksheetModel ws, StringBuilder sb, List<String[]> externalRels) {
        List<HyperlinkModel> hyperlinks = ws.getHyperlinks();
        // Handle the relevant branch before the state changes.
        if (hyperlinks.isEmpty()) return;

        sb.append("<hyperlinks>");
        int rIdCounter = 1;
        for (HyperlinkModel hl : hyperlinks) {
            int r1 = hl.getFirstRow(), c1 = hl.getFirstColumn();
            int r2 = r1 + hl.getTotalRows() - 1, c2 = c1 + hl.getTotalColumns() - 1;
            String ref = (r2 == r1 && c2 == c1)
                    ? XlsxWorkbookSerializerCommon.colLetter(c1) + (r1 + 1)
                    : XlsxWorkbookSerializerCommon.colLetter(c1) + (r1 + 1) + ":"
                      + XlsxWorkbookSerializerCommon.colLetter(c2) + (r2 + 1);
            sb.append("<hyperlink ref=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(ref)).append("\"");
            if (hl.getAddress() != null && !hl.getAddress().isEmpty()) {
                String rId = "rId" + rIdCounter++;
                sb.append(" r:id=\"").append(rId).append("\"");
                externalRels.add(new String[]{
                    rId,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    hl.getAddress(),
                    "External"
                });
            }
            if (hl.getSubAddress() != null && !hl.getSubAddress().isEmpty())
                sb.append(" location=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(hl.getSubAddress())).append("\"");
            if (hl.getScreenTip() != null && !hl.getScreenTip().isEmpty())
                sb.append(" tooltip=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(hl.getScreenTip())).append("\"");
            if (hl.getTextToDisplay() != null && !hl.getTextToDisplay().isEmpty())
                sb.append(" display=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(hl.getTextToDisplay())).append("\"");
            sb.append("/>");
        }
        sb.append("</hyperlinks>");
    }

    /**
     * Builds the sheet rels XML.
     * Each entry in {@code rels} is {@code [rId, type, target, targetMode?]}.
     * {@code targetMode} may be null or empty to omit the attribute (internal rels).
     */
    static byte[] buildSheetRelsXml(List<String[]> rels) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        // Walk the current collection so every entry is processed consistently.
        for (String[] rel : rels) {
            sb.append("<Relationship Id=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(rel[0]))
              .append("\" Type=\"").append(rel[1]).append("\"")
              .append(" Target=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(rel[2])).append("\"");
            if (rel.length > 3 && rel[3] != null && !rel[3].isEmpty())
                sb.append(" TargetMode=\"").append(rel[3]).append("\"");
            sb.append("/>");
        }
        sb.append("</Relationships>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    // =========================================================================
    // Load
    // =========================================================================

    /**
     * Loads the hyperlinks.
     * @param ws ws
     * @param doc doc
     * @param sheetRels sheet rels
     */
    static void loadHyperlinks(WorksheetModel ws, Document doc, Map<String, String> sheetRels) {
        NodeList hlNodes = doc.getElementsByTagNameNS("*", "hyperlink");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < hlNodes.getLength(); i++) {
            Element hl = (Element) hlNodes.item(i);
            String ref = hl.getAttribute("ref");
            if (ref.isEmpty()) continue;
            HyperlinkModel hlm = new HyperlinkModel();
            String[] refParts = ref.split(":");
            CellAddress a1 = XlsxWorkbookSerializerCommon.parseRef(refParts[0]);
            hlm.setFirstRow(a1.getRowIndex());
            hlm.setFirstColumn(a1.getColumnIndex());
            if (refParts.length == 2) {
                CellAddress a2 = XlsxWorkbookSerializerCommon.parseRef(refParts[1]);
                hlm.setTotalRows(a2.getRowIndex() - a1.getRowIndex() + 1);
                hlm.setTotalColumns(a2.getColumnIndex() - a1.getColumnIndex() + 1);
            }
            String rId = hl.getAttribute("r:id");
            if (rId.isEmpty())
                rId = hl.getAttributeNS(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
            if (!rId.isEmpty()) {
                String target = sheetRels.get(rId);
                if (target != null) hlm.setAddress(target);
            }
            String location = hl.getAttribute("location");
            if (!location.isEmpty()) hlm.setSubAddress(location);
            String tooltip = hl.getAttribute("tooltip");
            if (!tooltip.isEmpty()) hlm.setScreenTip(tooltip);
            String display = hl.getAttribute("display");
            if (!display.isEmpty()) hlm.setTextToDisplay(display);
            ws.getHyperlinks().add(hlm);
        }
    }
}
