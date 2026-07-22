package org.aspose.cells_foss;

import org.aspose.cells_foss.core.*;
import org.w3c.dom.*;

/**
 * Package-private helpers for building and loading XLSX worksheet protection.
 */
final class XlsxWorkbookWorksheetProtection {

    /**
     * Initializes a new XlsxWorkbookWorksheetProtection instance.
     */
    private XlsxWorkbookWorksheetProtection() {}

    // =========================================================================
    // Build
    // =========================================================================

    /**
     * Builds the sheet protection section.
     * @param ws ws
     * @param sb sb
     */
    static void buildSheetProtectionSection(WorksheetModel ws, StringBuilder sb) {
        WorksheetProtectionModel prot = ws.getProtection();
        // Handle the relevant branch before the state changes.
        if (!prot.getIsProtected()) return;

        sb.append("<sheetProtection sheet=\"1\"");
        if (prot.getObjects()) sb.append(" objects=\"1\"");
        if (prot.getScenarios()) sb.append(" scenarios=\"1\"");
        if (!prot.getFormatCells()) sb.append(" formatCells=\"0\"");
        if (!prot.getFormatColumns()) sb.append(" formatColumns=\"0\"");
        if (!prot.getFormatRows()) sb.append(" formatRows=\"0\"");
        if (!prot.getInsertColumns()) sb.append(" insertColumns=\"0\"");
        if (!prot.getInsertRows()) sb.append(" insertRows=\"0\"");
        if (!prot.getDeleteColumns()) sb.append(" deleteColumns=\"0\"");
        if (!prot.getDeleteRows()) sb.append(" deleteRows=\"0\"");
        if (prot.getAutoFilter()) sb.append(" autoFilter=\"1\"");
        if (prot.getSelectLockedCells()) sb.append(" selectLockedCells=\"1\"");
        if (prot.getSelectUnlockedCells()) sb.append(" selectUnlockedCells=\"1\"");
        sb.append("/>");
    }

    // =========================================================================
    // Load
    // =========================================================================

    /**
     * Loads the sheet protection.
     * @param ws ws
     * @param doc doc
     */
    static void loadSheetProtection(WorksheetModel ws, Document doc) {
        NodeList spNodes = doc.getElementsByTagNameNS("*", "sheetProtection");
        // Handle the relevant branch before the state changes.
        if (spNodes.getLength() == 0) return;

        Element sp = (Element) spNodes.item(0);
        WorksheetProtectionModel prot = ws.getProtection();
        if ("1".equals(sp.getAttribute("sheet"))) prot.setIsProtected(true);
        if ("1".equals(sp.getAttribute("objects"))) prot.setObjects(true);
        if ("1".equals(sp.getAttribute("scenarios"))) prot.setScenarios(true);
        if ("1".equals(sp.getAttribute("autoFilter"))) prot.setAutoFilter(true);
        if ("1".equals(sp.getAttribute("selectLockedCells"))) prot.setSelectLockedCells(true);
        if ("1".equals(sp.getAttribute("selectUnlockedCells"))) prot.setSelectUnlockedCells(true);
        prot.setFormatCells(!sp.hasAttribute("formatCells") || "1".equals(sp.getAttribute("formatCells")));
        prot.setFormatColumns(!sp.hasAttribute("formatColumns") || "1".equals(sp.getAttribute("formatColumns")));
        prot.setFormatRows(!sp.hasAttribute("formatRows") || "1".equals(sp.getAttribute("formatRows")));
        prot.setInsertColumns(!sp.hasAttribute("insertColumns") || "1".equals(sp.getAttribute("insertColumns")));
        prot.setInsertRows(!sp.hasAttribute("insertRows") || "1".equals(sp.getAttribute("insertRows")));
        prot.setDeleteColumns(!sp.hasAttribute("deleteColumns") || "1".equals(sp.getAttribute("deleteColumns")));
        prot.setDeleteRows(!sp.hasAttribute("deleteRows") || "1".equals(sp.getAttribute("deleteRows")));
    }
}

