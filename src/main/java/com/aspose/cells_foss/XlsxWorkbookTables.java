package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ListColumnModel;
import com.aspose.cells_foss.core.ListObjectModel;
import com.aspose.cells_foss.core.WorksheetModel;
import org.w3c.dom.*;
import java.nio.charset.StandardCharsets;
import java.util.List;

/**
 * Builds and loads XLSX table parts (xl/tables/table{N}.xml).
 */
final class XlsxWorkbookTables {

    private XlsxWorkbookTables() {}

    // =========================================================================
    // Build
    // =========================================================================

    /** Returns the XML bytes for xl/tables/table{N}.xml for the given table model. */
    static byte[] buildTableXml(ListObjectModel t) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
        sb.append(" id=\"").append(t.getId()).append("\"");
        sb.append(" name=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(t.getDisplayName())).append("\"");
        sb.append(" displayName=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(t.getDisplayName())).append("\"");
        sb.append(" ref=\"").append(buildRef(t)).append("\"");
        if (!t.isShowHeaderRow()) sb.append(" headerRowCount=\"0\"");
        if (t.isShowTotals()) sb.append(" totalsRowCount=\"1\"");
        if (!t.getComment().isBlank())
            sb.append(" comment=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(t.getComment())).append("\"");
        sb.append(">");

        if (t.isAutoFilterEnabled() && t.isShowHeaderRow()) {
            sb.append("<autoFilter ref=\"").append(buildRef(t)).append("\"/>");
        }

        sb.append("<tableColumns count=\"").append(t.getColumns().size()).append("\">");
        for (ListColumnModel col : t.getColumns()) {
            sb.append("<tableColumn id=\"").append(col.getId()).append("\"");
            sb.append(" name=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(col.getName())).append("\"");
            if (!col.getTotalsRowFunction().equals("none"))
                sb.append(" totalsRowFunction=\"").append(col.getTotalsRowFunction()).append("\"");
            if (!col.getTotalsRowLabel().isBlank())
                sb.append(" totalsRowLabel=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(col.getTotalsRowLabel())).append("\"");
            sb.append("/>");
        }
        sb.append("</tableColumns>");

        sb.append("<tableStyleInfo");
        if (!t.getTableStyleName().isBlank())
            sb.append(" name=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(t.getTableStyleName())).append("\"");
        sb.append(" showFirstColumn=\"").append(t.isShowTableStyleFirstColumn() ? "1" : "0").append("\"");
        sb.append(" showLastColumn=\"").append(t.isShowTableStyleLastColumn() ? "1" : "0").append("\"");
        sb.append(" showRowStripes=\"").append(t.isShowTableStyleRowStripes() ? "1" : "0").append("\"");
        sb.append(" showColumnStripes=\"").append(t.isShowTableStyleColumnStripes() ? "1" : "0").append("\"");
        sb.append("/>");

        sb.append("</table>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    /** Builds the sheet-level <tableParts> element and populates rels. Returns the XML snippet. */
    static String buildTablePartsSnippet(WorksheetModel ws, int sheetIndex,
                                          List<String[]> sheetRels, int[] tableCounter) {
        List<ListObjectModel> tables = ws.getListObjects();
        if (tables.isEmpty()) return "";
        StringBuilder sb = new StringBuilder();
        sb.append("<tableParts count=\"").append(tables.size()).append("\">");
        for (int i = 0; i < tables.size(); i++) {
            int tIdx = tableCounter[0]++;
            String rId = "rIdT" + tIdx;
            String partName = "../tables/table" + tIdx + ".xml";
            sheetRels.add(new String[]{
                rId,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
                partName,
                null
            });
            sb.append("<tablePart r:id=\"").append(rId).append("\"/>");
        }
        sb.append("</tableParts>");
        return sb.toString();
    }

    // =========================================================================
    // Load
    // =========================================================================

    static void loadTable(WorksheetModel ws, Document doc) {
        if (doc == null) return;
        Element root = doc.getDocumentElement();
        ListObjectModel model = new ListObjectModel();

        String idStr = root.getAttribute("id");
        if (!idStr.isBlank()) {
            try { model.setId(Integer.parseInt(idStr)); } catch (NumberFormatException ignored) {}
        }
        model.setDisplayName(root.getAttribute("displayName"));
        if (model.getDisplayName().isBlank()) model.setDisplayName(root.getAttribute("name"));
        model.setComment(root.getAttribute("comment"));

        String ref = root.getAttribute("ref");
        if (!ref.isBlank()) parseRef(ref, model);

        String hrc = root.getAttribute("headerRowCount");
        model.setShowHeaderRow(hrc.isBlank() || !hrc.equals("0"));

        String trc = root.getAttribute("totalsRowCount");
        model.setShowTotals(!trc.isBlank() && !trc.equals("0"));

        NodeList styleInfo = root.getElementsByTagNameNS("*", "tableStyleInfo");
        if (styleInfo.getLength() > 0) {
            Element si = (Element) styleInfo.item(0);
            model.setTableStyleName(si.getAttribute("name"));
            model.setShowTableStyleFirstColumn("1".equals(si.getAttribute("showFirstColumn")));
            model.setShowTableStyleLastColumn("1".equals(si.getAttribute("showLastColumn")));
            model.setShowTableStyleRowStripes(!"0".equals(si.getAttribute("showRowStripes")));
            model.setShowTableStyleColumnStripes("1".equals(si.getAttribute("showColumnStripes")));
        }

        NodeList af = root.getElementsByTagNameNS("*", "autoFilter");
        model.setAutoFilterEnabled(af.getLength() > 0);

        NodeList cols = root.getElementsByTagNameNS("*", "tableColumn");
        for (int i = 0; i < cols.getLength(); i++) {
            Element ce = (Element) cols.item(i);
            ListColumnModel col = new ListColumnModel();
            String cid = ce.getAttribute("id");
            if (!cid.isBlank()) { try { col.setId(Integer.parseInt(cid)); } catch (NumberFormatException ignored) {} }
            col.setName(ce.getAttribute("name"));
            col.setTotalsRowFunction(ce.getAttribute("totalsRowFunction").isBlank() ? "none" : ce.getAttribute("totalsRowFunction"));
            col.setTotalsRowLabel(ce.getAttribute("totalsRowLabel"));
            model.getColumns().add(col);
        }

        ws.getListObjects().add(model);
    }

    // =========================================================================
    // Private helpers
    // =========================================================================

    private static String buildRef(ListObjectModel t) {
        int startRow = t.getStartRow(), startCol = t.getStartColumn();
        int endRow = t.getEndRow(), endCol = t.getEndColumn();
        return XlsxWorkbookSerializerCommon.colLetter(startCol) + (startRow + 1)
             + ":" + XlsxWorkbookSerializerCommon.colLetter(endCol) + (endRow + 1);
    }

    private static void parseRef(String ref, ListObjectModel model) {
        String[] parts = ref.split(":");
        if (parts.length != 2) return;
        try {
            com.aspose.cells_foss.core.CellAddress a1 = XlsxWorkbookSerializerCommon.parseRef(parts[0]);
            com.aspose.cells_foss.core.CellAddress a2 = XlsxWorkbookSerializerCommon.parseRef(parts[1]);
            model.setStartRow(a1.getRowIndex());
            model.setStartColumn(a1.getColumnIndex());
            model.setEndRow(a2.getRowIndex());
            model.setEndColumn(a2.getColumnIndex());
        } catch (Exception ignored) {}
    }
}
