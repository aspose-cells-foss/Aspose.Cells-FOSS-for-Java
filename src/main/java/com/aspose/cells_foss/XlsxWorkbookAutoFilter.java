package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;

/**
 * Package-private helpers for building and loading XLSX autoFilter elements.
 */
final class XlsxWorkbookAutoFilter {

    /**
     * Initializes a new XlsxWorkbookAutoFilter instance.
     */
    private XlsxWorkbookAutoFilter() {}

    // =========================================================================
    // Build
    // =========================================================================

    /**
     * Builds the auto filter section.
     * @param ws ws
     * @param sb sb
     */
    static void buildAutoFilterSection(WorksheetModel ws, StringBuilder sb) {
        AutoFilterModel af = ws.getAutoFilter();
        // Handle the relevant branch before the state changes.
        if (af.getRange() == null || af.getRange().isEmpty()) return;

        sb.append("<autoFilter ref=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(af.getRange())).append("\">");

        for (AutoFilterModel.FilterColumnModel fc : af.getFilterColumns()) {
            sb.append("<filterColumn colId=\"").append(fc.getColumnIndex()).append("\"");
            if (fc.getHiddenButton()) sb.append(" hiddenButton=\"1\"");
            sb.append(">");
            if (!fc.getFilters().isEmpty()) {
                sb.append("<filters>");
                for (String val : fc.getFilters())
                    sb.append("<filter val=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(val)).append("\"/>");
                sb.append("</filters>");
            }
            if (!fc.getCustomFilters().isEmpty()) {
                sb.append("<customFilters");
                if (fc.isCustomFiltersAnd()) sb.append(" and=\"1\"");
                sb.append(">");
                for (AutoFilterModel.AutoFilterCustomFilterModel cf : fc.getCustomFilters())
                    sb.append("<customFilter operator=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(cf.getOperator()))
                      .append("\" val=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(cf.getValue())).append("\"/>");
                sb.append("</customFilters>");
            }
            if (fc.getColorFilter().isEnabled()) {
                AutoFilterModel.AutoFilterColorFilterModel cfm = fc.getColorFilter();
                sb.append("<colorFilter");
                if (cfm.getDifferentialStyleId() != null)
                    sb.append(" dxfId=\"").append(cfm.getDifferentialStyleId()).append("\"");
                sb.append(" cellColor=\"").append(cfm.isCellColor() ? "1" : "0").append("\"/>");
            }
            if (fc.getDynamicFilter().isEnabled()) {
                AutoFilterModel.AutoFilterDynamicFilterModel dfm = fc.getDynamicFilter();
                sb.append("<dynamicFilter type=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(dfm.getType())).append("\"");
                if (dfm.getValue() != null) sb.append(" val=\"").append(dfm.getValue()).append("\"");
                if (dfm.getMaxValue() != null) sb.append(" maxVal=\"").append(dfm.getMaxValue()).append("\"");
                sb.append("/>");
            }
            if (fc.getTop10().isEnabled()) {
                AutoFilterModel.AutoFilterTop10Model t10 = fc.getTop10();
                sb.append("<top10 top=\"").append(t10.isTop() ? "1" : "0")
                  .append("\" percent=\"").append(t10.isPercent() ? "1" : "0").append("\"");
                if (t10.getValue() != null) sb.append(" val=\"").append(t10.getValue()).append("\"");
                if (t10.getFilterValue() != null) sb.append(" filterVal=\"").append(t10.getFilterValue()).append("\"");
                sb.append("/>");
            }
            sb.append("</filterColumn>");
        }

        AutoFilterModel.AutoFilterSortStateModel ss = af.getSortState();
        if (ss.getRef() != null && !ss.getRef().isEmpty()) {
            sb.append("<sortState ref=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(ss.getRef())).append("\"");
            if (ss.isCaseSensitive()) sb.append(" caseSensitive=\"1\"");
            if (ss.getSortMethod() != null && !ss.getSortMethod().isEmpty())
                sb.append(" sortMethod=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(ss.getSortMethod())).append("\"");
            sb.append(">");
            for (AutoFilterModel.AutoFilterSortConditionModel sc : ss.getConditions()) {
                sb.append("<sortCondition ref=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(sc.getRef())).append("\"");
                if (sc.isDescending()) sb.append(" descending=\"1\"");
                if (sc.getSortBy() != null && !sc.getSortBy().isEmpty()) sb.append(" sortBy=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(sc.getSortBy())).append("\"");
                if (sc.getCustomList() != null && !sc.getCustomList().isEmpty()) sb.append(" customList=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(sc.getCustomList())).append("\"");
                if (sc.getDifferentialStyleId() != null) sb.append(" dxfId=\"").append(sc.getDifferentialStyleId()).append("\"");
                if (sc.getIconSet() != null && !sc.getIconSet().isEmpty()) sb.append(" iconSet=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(sc.getIconSet())).append("\"");
                if (sc.getIconId() != null) sb.append(" iconId=\"").append(sc.getIconId()).append("\"");
                sb.append("/>");
            }
            sb.append("</sortState>");
        }

        sb.append("</autoFilter>");
    }

    // =========================================================================
    // Load
    // =========================================================================

    /**
     * Loads the auto filter.
     * @param ws ws
     * @param doc doc
     */
    static void loadAutoFilter(WorksheetModel ws, Document doc) {
        NodeList afNodes = doc.getElementsByTagNameNS("*", "autoFilter");
        if (afNodes.getLength() == 0) return;

        Element af = (Element) afNodes.item(0);
        ws.getAutoFilter().setRange(af.getAttribute("ref"));

        NodeList fcNodes = af.getElementsByTagNameNS("*", "filterColumn");
        for (int i = 0; i < fcNodes.getLength(); i++) {
            Element fc = (Element) fcNodes.item(i);
            int colId = XlsxWorkbookSerializerCommon.parseInt(fc.getAttribute("colId"), 0);
            AutoFilterModel.FilterColumnModel fcm = new AutoFilterModel.FilterColumnModel();
            fcm.setColumnIndex(colId);
            if ("1".equals(fc.getAttribute("hiddenButton"))) fcm.setHiddenButton(true);

            // filters
            NodeList filterNodes = fc.getElementsByTagNameNS("*", "filter");
            for (int j = 0; j < filterNodes.getLength(); j++)
                fcm.getFilters().add(((Element) filterNodes.item(j)).getAttribute("val"));

            // customFilters
            NodeList cfParent = fc.getElementsByTagNameNS("*", "customFilters");
            if (cfParent.getLength() > 0) {
                Element cfp = (Element) cfParent.item(0);
                fcm.setCustomFiltersAnd("1".equals(cfp.getAttribute("and")));
                NodeList cfNodes = cfp.getElementsByTagNameNS("*", "customFilter");
                for (int j = 0; j < cfNodes.getLength(); j++) {
                    Element cf = (Element) cfNodes.item(j);
                    AutoFilterModel.AutoFilterCustomFilterModel cfm2 = new AutoFilterModel.AutoFilterCustomFilterModel();
                    cfm2.setOperator(cf.getAttribute("operator"));
                    cfm2.setValue(cf.getAttribute("val"));
                    fcm.getCustomFilters().add(cfm2);
                }
            }

            // top10
            NodeList t10Nodes = fc.getElementsByTagNameNS("*", "top10");
            if (t10Nodes.getLength() > 0) {
                Element t10 = (Element) t10Nodes.item(0);
                fcm.getTop10().setEnabled(true);
                fcm.getTop10().setTop(!"0".equals(t10.getAttribute("top")));
                fcm.getTop10().setPercent("1".equals(t10.getAttribute("percent")));
                String val = t10.getAttribute("val");
                if (!val.isEmpty()) fcm.getTop10().setValue(Double.parseDouble(val));
                String filterVal = t10.getAttribute("filterVal");
                if (!filterVal.isEmpty()) fcm.getTop10().setFilterValue(Double.parseDouble(filterVal));
            }

            // dynamicFilter
            NodeList dynNodes = fc.getElementsByTagNameNS("*", "dynamicFilter");
            if (dynNodes.getLength() > 0) {
                Element dyn = (Element) dynNodes.item(0);
                fcm.getDynamicFilter().setEnabled(true);
                fcm.getDynamicFilter().setType(dyn.getAttribute("type"));
                String val = dyn.getAttribute("val");
                if (!val.isEmpty()) fcm.getDynamicFilter().setValue(Double.parseDouble(val));
                String maxVal = dyn.getAttribute("maxVal");
                if (!maxVal.isEmpty()) fcm.getDynamicFilter().setMaxValue(Double.parseDouble(maxVal));
            }

            ws.getAutoFilter().getFilterColumns().add(fcm);
        }

        // sortState
        NodeList ssNodes = af.getElementsByTagNameNS("*", "sortState");
        if (ssNodes.getLength() > 0) {
            Element ss = (Element) ssNodes.item(0);
            ws.getAutoFilter().getSortState().setRef(ss.getAttribute("ref"));
            ws.getAutoFilter().getSortState().setCaseSensitive("1".equals(ss.getAttribute("caseSensitive")));
            ws.getAutoFilter().getSortState().setSortMethod(ss.getAttribute("sortMethod"));
            NodeList scNodes = ss.getElementsByTagNameNS("*", "sortCondition");
            for (int i = 0; i < scNodes.getLength(); i++) {
                Element sc = (Element) scNodes.item(i);
                AutoFilterModel.AutoFilterSortConditionModel scm = new AutoFilterModel.AutoFilterSortConditionModel();
                scm.setRef(sc.getAttribute("ref"));
                scm.setDescending("1".equals(sc.getAttribute("descending")));
                String sortBy = sc.getAttribute("sortBy");
                scm.setSortBy(sortBy.isEmpty() ? "value" : sortBy);
                scm.setCustomList(sc.getAttribute("customList"));
                String dxfId = sc.getAttribute("dxfId");
                if (!dxfId.isEmpty()) scm.setDifferentialStyleId(Integer.parseInt(dxfId));
                scm.setIconSet(sc.getAttribute("iconSet"));
                String iconId = sc.getAttribute("iconId");
                if (!iconId.isEmpty()) scm.setIconId(Integer.parseInt(iconId));
                ws.getAutoFilter().getSortState().getConditions().add(scm);
            }
        }
    }
}
