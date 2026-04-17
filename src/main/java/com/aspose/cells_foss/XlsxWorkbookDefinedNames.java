package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;
import java.util.*;

/**
 * Package-private helpers for building and loading XLSX defined names
 * (print area, print titles).
 */
final class XlsxWorkbookDefinedNames {

    /**
     * Initializes a new XlsxWorkbookDefinedNames instance.
     */
    private XlsxWorkbookDefinedNames() {}

    // =========================================================================
    // Build
    // =========================================================================

    /**
     * Builds the defined names xml.
     * @param model model
     * @param sb sb
     */
    static void buildDefinedNamesXml(WorkbookModel model, StringBuilder sb) {
        List<WorksheetModel> sheets = model.getWorksheets();
        boolean hasDefinedNames = false;
        // Walk the current collection so every entry is processed consistently.
        for (WorksheetModel ws : sheets) {
            PageSetupModel ps = ws.getPageSetup();
            if (ps.getPrintArea() != null && !ps.getPrintArea().isEmpty()) hasDefinedNames = true;
            if (ps.getPrintTitleRows() != null && !ps.getPrintTitleRows().isEmpty()) hasDefinedNames = true;
            if (ps.getPrintTitleColumns() != null && !ps.getPrintTitleColumns().isEmpty()) hasDefinedNames = true;
        }
        if (!hasDefinedNames) return;

        sb.append("<definedNames>");
        for (int i = 0; i < sheets.size(); i++) {
            PageSetupModel ps = sheets.get(i).getPageSetup();
            String sheetName = sheets.get(i).getName();
            if (ps.getPrintArea() != null && !ps.getPrintArea().isEmpty()) {
                sb.append("<definedName name=\"_xlnm.Print_Area\" localSheetId=\"").append(i)
                  .append("\">").append(XlsxWorkbookSerializerCommon.xmlText(sheetName + "!" + ps.getPrintArea()))
                  .append("</definedName>");
            }
            String titleRows = ps.getPrintTitleRows();
            String titleCols = ps.getPrintTitleColumns();
            if ((titleRows != null && !titleRows.isEmpty()) || (titleCols != null && !titleCols.isEmpty())) {
                String titles = "";
                if (titleCols != null && !titleCols.isEmpty()) titles += sheetName + "!" + titleCols;
                if (titleRows != null && !titleRows.isEmpty()) {
                    if (!titles.isEmpty()) titles += ",";
                    titles += sheetName + "!" + titleRows;
                }
                sb.append("<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"").append(i)
                  .append("\">").append(XlsxWorkbookSerializerCommon.xmlText(titles)).append("</definedName>");
            }
        }
        sb.append("</definedNames>");
    }

    // =========================================================================
    // Load
    // =========================================================================

    /** Returns a map of localSheetId → [printArea, titleRows, titleCols]. */
    static Map<Integer, String[]> loadDefinedNames(Document wbDoc) {
        Map<Integer, String[]> result = new HashMap<>();
        NodeList nodes = wbDoc.getElementsByTagNameNS("*", "definedName");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < nodes.getLength(); i++) {
            Element el = (Element) nodes.item(i);
            String name = el.getAttribute("name");
            String localId = el.getAttribute("localSheetId");
            int sid = localId.isEmpty() ? -1 : Integer.parseInt(localId);
            String val = el.getTextContent().trim();
            if (sid < 0) continue;
            String[] arr = result.computeIfAbsent(sid, k -> new String[3]);
            if ("_xlnm.Print_Area".equals(name)) {
                arr[0] = val;
            } else if ("_xlnm.Print_Titles".equals(name)) {
                for (String part : val.split(",")) {
                    part = part.trim();
                    String stripped = XlsxWorkbookSerializerCommon.stripSheetPrefix(part);
                    if (stripped.matches("\\$[A-Z]+:\\$[A-Z]+")) arr[2] = stripped;
                    else if (stripped.matches("\\$\\d+:\\$\\d+")) arr[1] = stripped;
                }
            }
        }
        return result;
    }
}
