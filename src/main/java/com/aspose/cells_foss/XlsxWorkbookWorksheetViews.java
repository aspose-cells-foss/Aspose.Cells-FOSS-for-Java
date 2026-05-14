package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;

/**
 * Package-private helpers for building and loading XLSX worksheet view elements
 * (tab color, sheetViews, cols).
 */
final class XlsxWorkbookWorksheetViews {

    /**
     * Initializes a new XlsxWorkbookWorksheetViews instance.
     */
    private XlsxWorkbookWorksheetViews() {}

    // =========================================================================
    // Build
    // =========================================================================

    /**
     * Builds the sheet pr.
     * @param ws ws
     * @param sb sb
     */
    static void buildSheetPr(WorksheetModel ws, StringBuilder sb) {
        ColorValue tabColor = ws.getTabColor();
        // Handle the relevant branch before the state changes.
        if (tabColor != null) {
            sb.append("<sheetPr><tabColor rgb=\"")
              .append(String.format("FF%02X%02X%02X",
                  Byte.toUnsignedInt(tabColor.getR()),
                  Byte.toUnsignedInt(tabColor.getG()),
                  Byte.toUnsignedInt(tabColor.getB())))
              .append("\"/></sheetPr>");
        }
    }

    /**
     * Builds the sheet views.
     * @param ws ws
     * @param sb sb
     */
    static void buildSheetViews(WorksheetModel ws, StringBuilder sb) {
        WorksheetViewModel view = ws.getView();
        sb.append("<sheetViews><sheetView workbookViewId=\"0\"");
        // Handle the relevant branch before the state changes.
        if (!view.getShowGridLines()) sb.append(" showGridLines=\"0\"");
        if (!view.getShowRowColumnHeaders()) sb.append(" showRowColHeaders=\"0\"");
        if (!view.getShowZeros()) sb.append(" showZeros=\"0\"");
        if (view.getRightToLeft()) sb.append(" rightToLeft=\"1\"");
        if (view.getZoomScale() != 100) sb.append(" zoomScale=\"").append(view.getZoomScale()).append("\"");

        int freezeRow = view.getFreezeRow();
        int freezeCol = view.getFreezeColumn();
        boolean hasFrozenRows = freezeRow > 0;
        boolean hasFrozenCols = freezeCol > 0;
        if (hasFrozenRows || hasFrozenCols) {
            sb.append(">");
            appendFreezePane(sb, freezeRow, freezeCol);
            sb.append("</sheetView>");
        } else {
            sb.append("/>");
        }
        sb.append("</sheetViews>");
    }

    private static void appendFreezePane(StringBuilder sb, int freezeRow, int freezeCol) {
        boolean hasFrozenRows = freezeRow > 0;
        boolean hasFrozenCols = freezeCol > 0;

        // Determine the top-left cell of the unfrozen pane
        String topLeftCell = colName(freezeCol) + (freezeRow + 1);

        // Determine active pane
        String activePane;
        if (hasFrozenRows && hasFrozenCols) activePane = "bottomRight";
        else if (hasFrozenRows)             activePane = "bottomLeft";
        else                                activePane = "topRight";

        sb.append("<pane");
        if (hasFrozenCols) sb.append(" xSplit=\"").append(freezeCol).append("\"");
        if (hasFrozenRows) sb.append(" ySplit=\"").append(freezeRow).append("\"");
        sb.append(" topLeftCell=\"").append(topLeftCell).append("\"");
        sb.append(" activePane=\"").append(activePane).append("\"");
        sb.append(" state=\"frozen\"/>");

        // Selection elements for each pane quadrant
        sb.append("<selection pane=\"topLeft\"/>");
        if (hasFrozenRows && hasFrozenCols) {
            sb.append("<selection pane=\"topRight\"/>");
            sb.append("<selection pane=\"bottomLeft\"/>");
        }
        sb.append("<selection pane=\"").append(activePane)
          .append("\" activeCell=\"").append(topLeftCell)
          .append("\" sqref=\"").append(topLeftCell).append("\"/>");
    }

    /** Converts a zero-based column index to its Excel column letter(s) (e.g. 0→"A", 25→"Z", 26→"AA"). */
    private static String colName(int colIndex) {
        int idx = colIndex + 1;
        StringBuilder result = new StringBuilder();
        while (idx > 0) {
            idx--;
            result.append((char) ('A' + (idx % 26)));
            idx /= 26;
        }
        return result.reverse().toString();
    }

    /**
     * Builds the cols.
     * @param ws ws
     * @param sb sb
     */
    static void buildCols(WorksheetModel ws, StringBuilder sb) {
        // Handle the relevant branch before the state changes.
        if (ws.getColumns().isEmpty()) return;
        sb.append("<cols>");
        for (ColumnRangeModel col : ws.getColumns()) {
            sb.append("<col min=\"").append(col.getMinColumnIndex() + 1)
              .append("\" max=\"").append(col.getMaxColumnIndex() + 1).append("\"");
            if (col.getWidth() != null) {
                sb.append(" width=\"").append(XlsxWorkbookSerializerCommon.fmt(col.getWidth())).append("\"");
                if (col.getBestFit()) sb.append(" bestFit=\"1\"");
                sb.append(" customWidth=\"1\"");
            }
            if (col.getHidden()) sb.append(" hidden=\"1\"");
            if (col.getOutlineLevel() > 0) sb.append(" outlineLevel=\"").append(col.getOutlineLevel()).append("\"");
            if (col.getCollapsed()) sb.append(" collapsed=\"1\"");
            sb.append("/>");
        }
        sb.append("</cols>");
    }

    // =========================================================================
    // Load
    // =========================================================================

    /**
     * Loads the tab color.
     * @param ws ws
     * @param doc doc
     */
    static void loadTabColor(WorksheetModel ws, Document doc) {
        NodeList tabColors = doc.getElementsByTagNameNS("*", "tabColor");
        // Handle the relevant branch before the state changes.
        if (tabColors.getLength() > 0) {
            String rgb = ((Element) tabColors.item(0)).getAttribute("rgb");
            if (rgb.length() == 8) {
                int r = Integer.parseInt(rgb.substring(2, 4), 16);
                int g = Integer.parseInt(rgb.substring(4, 6), 16);
                int b = Integer.parseInt(rgb.substring(6, 8), 16);
                ws.setTabColor(new ColorValue((byte) 255, (byte) r, (byte) g, (byte) b));
            }
        }
    }

    /**
     * Loads the sheet view.
     * @param ws ws
     * @param doc doc
     */
    static void loadSheetView(WorksheetModel ws, Document doc) {
        NodeList svNodes = doc.getElementsByTagNameNS("*", "sheetView");
        // Handle the relevant branch before the state changes.
        if (svNodes.getLength() == 0) return;
        Element sv = (Element) svNodes.item(0);
        if ("0".equals(sv.getAttribute("showGridLines"))) ws.getView().setShowGridLines(false);
        if ("0".equals(sv.getAttribute("showRowColHeaders"))) ws.getView().setShowRowColumnHeaders(false);
        if ("0".equals(sv.getAttribute("showZeros"))) ws.getView().setShowZeros(false);
        if ("1".equals(sv.getAttribute("rightToLeft"))) ws.getView().setRightToLeft(true);
        String zoom = sv.getAttribute("zoomScale");
        if (!zoom.isEmpty()) ws.getView().setZoomScale(Integer.parseInt(zoom));

        // Load freeze pane: only "frozen" state is supported (not split)
        NodeList paneNodes = sv.getElementsByTagNameNS("*", "pane");
        if (paneNodes.getLength() > 0) {
            Element pane = (Element) paneNodes.item(0);
            if ("frozen".equals(pane.getAttribute("state"))) {
                String xSplit = pane.getAttribute("xSplit");
                String ySplit = pane.getAttribute("ySplit");
                if (!xSplit.isEmpty()) ws.getView().setFreezeColumn(
                    XlsxWorkbookSerializerCommon.parseInt(xSplit, 0));
                if (!ySplit.isEmpty()) ws.getView().setFreezeRow(
                    XlsxWorkbookSerializerCommon.parseInt(ySplit, 0));
            }
        }
    }

    /**
     * Loads the cols.
     * @param ws ws
     * @param doc doc
     */
    static void loadCols(WorksheetModel ws, Document doc) {
        NodeList colNodes = doc.getElementsByTagNameNS("*", "col");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < colNodes.getLength(); i++) {
            Element col = (Element) colNodes.item(i);
            int min = XlsxWorkbookSerializerCommon.parseInt(col.getAttribute("min"), 1) - 1;
            int max = XlsxWorkbookSerializerCommon.parseInt(col.getAttribute("max"), 1) - 1;
            ColumnRangeModel cr = new ColumnRangeModel();
            cr.setMinColumnIndex(min);
            cr.setMaxColumnIndex(max);
            String w = col.getAttribute("width");
            if (!w.isEmpty()) cr.setWidth(Double.parseDouble(w));
            if ("1".equals(col.getAttribute("hidden"))) cr.setHidden(true);
            if ("1".equals(col.getAttribute("bestFit"))) cr.setBestFit(true);
            String ol = col.getAttribute("outlineLevel");
            if (!ol.isEmpty()) { try { cr.setOutlineLevel(Integer.parseInt(ol)); } catch (NumberFormatException ignored) {} }
            if ("1".equals(col.getAttribute("collapsed"))) cr.setCollapsed(true);
            ws.getColumns().add(cr);
        }
    }
}
