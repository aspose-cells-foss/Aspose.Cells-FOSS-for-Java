package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;
import java.util.*;

/**
 * Package-private helpers for building and loading XLSX conditional formatting.
 */
final class XlsxWorkbookConditionalFormatting {

    /**
     * Initializes a new XlsxWorkbookConditionalFormatting instance.
     */
    private XlsxWorkbookConditionalFormatting() {}

    // =========================================================================
    // Build
    // =========================================================================

    /**
     * Builds the conditional formatting section.
     * @param ws ws
     * @param sb sb
     * @param styleTable style table
     */
    static void buildConditionalFormattingSection(WorksheetModel ws, StringBuilder sb,
                                                  XlsxWorkbookStyles.StyleTable styleTable) {
        List<ConditionalFormattingModel> cfs = ws.getConditionalFormattings();
        if (cfs.isEmpty()) return;

        // Sort by first area's top-left cell to produce deterministic output
        List<ConditionalFormattingModel> ordered = new ArrayList<>(cfs.size());
        for (ConditionalFormattingModel cf : cfs) {
            if (!cf.getAreas().isEmpty() && !cf.getConditions().isEmpty()) {
                FormatConditionCollection.sortAreas(cf.getAreas());
                ordered.add(cf);
            }
        }
        ordered.sort((a, b) -> FormatConditionCollection.compareAreas(a.getAreas().get(0), b.getAreas().get(0)));

        for (ConditionalFormattingModel cf : ordered) {
            sb.append("<conditionalFormatting sqref=\"").append(buildSqref(cf.getAreas())).append("\">");
            for (int ci = 0; ci < cf.getConditions().size(); ci++) {
                FormatConditionModel cond = cf.getConditions().get(ci);
                buildCfRule(cf, cond, ci, styleTable, sb);
            }
            sb.append("</conditionalFormatting>");
        }
    }

    /**
     * Builds the cf rule.
     * @param cf cf
     * @param cond cond
     * @param conditionIndex zero-based condition index
     * @param styleTable style table
     * @param sb sb
     */
    private static void buildCfRule(ConditionalFormattingModel cf, FormatConditionModel cond,
                                    int conditionIndex, XlsxWorkbookStyles.StyleTable styleTable,
                                    StringBuilder sb) {
        sb.append("<cfRule type=\"").append(toRuleTypeName(cond)).append("\"");
        int priority = cond.getPriority() > 0 ? cond.getPriority() : conditionIndex + 1;
        sb.append(" priority=\"").append(priority).append("\"");
        int dxfId = styleTable.registerDifferentialStyle(cond.getStyle());
        // Handle the relevant branch before the state changes.
        if (dxfId >= 0) sb.append(" dxfId=\"").append(dxfId).append("\"");
        if (cond.getStopIfTrue()) sb.append(" stopIfTrue=\"1\"");

        switch (cond.getType()) {
            case CELL_VALUE: {
                String op = toOperatorName(cond.getOperator());
                if (op != null) sb.append(" operator=\"").append(op).append("\"");
                sb.append(">");
                appendFormula(sb, cond.getFormula1());
                appendFormula(sb, cond.getFormula2());
                break;
            }
            case EXPRESSION:
                sb.append(">");
                appendFormula(sb, cond.getFormula1());
                break;
            case CONTAINS_TEXT:
            case NOT_CONTAINS_TEXT:
            case begins_with:
            case ENDS_WITH: {
                String text = cond.getFormula1();
                if (text != null && !text.isEmpty()) {
                    sb.append(" text=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(text)).append("\">");
                    String anchor = getAnchorCell(cf);
                    appendFormula(sb, buildTextRuleFormula(cond.getType(), text, anchor));
                } else {
                    sb.append(">");
                }
                break;
            }
            case TIME_PERIOD:
                if (cond.getTimePeriod() != null && !cond.getTimePeriod().isEmpty())
                    sb.append(" timePeriod=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(cond.getTimePeriod())).append("\"");
                sb.append(">");
                break;
            case TOP_10:
            case BOTTOM_10:
                sb.append(" bottom=\"").append(cond.getType() == FormatConditionType.BOTTOM_10 || !cond.getTop() ? "1" : "0").append("\"");
                if (cond.getPercent()) sb.append(" percent=\"1\"");
                if (cond.getRank() > 0) sb.append(" rank=\"").append(cond.getRank()).append("\"");
                sb.append(">");
                break;
            case ABOVE_AVERAGE:
            case BELOW_AVERAGE:
                if (cond.getType() == FormatConditionType.BELOW_AVERAGE || !cond.getAbove())
                    sb.append(" aboveAverage=\"0\"");
                if (cond.getStandardDeviation() > 0)
                    sb.append(" stdDev=\"").append(cond.getStandardDeviation()).append("\"");
                sb.append(">");
                break;
            case DUPLICATE_VALUES:
            case UNIQUE_VALUES:
                sb.append(">");
                break;
            case COLOR_SCALE:
                sb.append(">");
                buildColorScale(cond, sb);
                break;
            case DATA_BAR:
                sb.append(">");
                buildDataBar(cond, sb);
                break;
            case ICON_SET:
                sb.append(">");
                buildIconSet(cond, sb);
                break;
            default:
                sb.append(">");
                break;
        }
        sb.append("</cfRule>");
    }

    /**
     * Appends formula.
     * @param sb sb
     * @param formula formula
     */
    private static void appendFormula(StringBuilder sb, String formula) {
        // Handle the relevant branch before the state changes.
        if (formula != null && !formula.isEmpty())
            sb.append("<formula>").append(XlsxWorkbookSerializerCommon.xmlText(formula)).append("</formula>");
    }

    /**
     * Builds the color scale.
     * @param cond cond
     * @param sb sb
     */
    private static void buildColorScale(FormatConditionModel cond, StringBuilder sb) {
        sb.append("<colorScale>");
        sb.append("<cfvo type=\"min\"/>");
        // Handle the relevant branch before the state changes.
        if (cond.getColorScaleCount() == 3)
            sb.append("<cfvo type=\"percentile\" val=\"50\"/>");
        sb.append("<cfvo type=\"max\"/>");
        appendColor(cond.getMinColor(), new ColorValue((byte)255,(byte)248,(byte)105,(byte)107), sb);
        if (cond.getColorScaleCount() == 3)
            appendColor(cond.getMidColor(), new ColorValue((byte)255,(byte)255,(byte)235,(byte)132), sb);
        appendColor(cond.getMaxColor(), new ColorValue((byte)255,(byte)99,(byte)190,(byte)123), sb);
        sb.append("</colorScale>");
    }

    /**
     * Builds the data bar.
     * @param cond cond
     * @param sb sb
     */
    private static void buildDataBar(FormatConditionModel cond, StringBuilder sb) {
        ColorValue defaultBar = new ColorValue((byte)255,(byte)99,(byte)142,(byte)198);
        sb.append("<dataBar>");
        sb.append("<cfvo type=\"min\"/><cfvo type=\"max\"/>");
        appendColor(cond.getBarColor(), defaultBar, sb);
        sb.append("</dataBar>");
    }

    /**
     * Builds the icon set.
     * @param cond cond
     * @param sb sb
     */
    private static void buildIconSet(FormatConditionModel cond, StringBuilder sb) {
        String iconSetType = (cond.getIconSetType() != null && !cond.getIconSetType().isEmpty())
                ? cond.getIconSetType() : "3TrafficLights1";
        sb.append("<iconSet iconSet=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(iconSetType)).append("\"");
        // Handle the relevant branch before the state changes.
        if (cond.getReverseIcons()) sb.append(" reverse=\"1\"");
        if (cond.getShowIconOnly()) sb.append(" showValue=\"0\"");
        sb.append(">");
        int iconCount = getIconCount(iconSetType);
        for (int i = 0; i < iconCount; i++)
            sb.append("<cfvo type=\"percent\" val=\"").append((100 * i) / iconCount).append("\"/>");
        sb.append("</iconSet>");
    }

    /**
     * Appends color.
     * @param actual actual
     * @param fallback fallback
     * @param sb sb
     */
    private static void appendColor(ColorValue actual, ColorValue fallback, StringBuilder sb) {
        ColorValue c = isEmptyColor(actual) ? fallback : actual;
        sb.append("<color rgb=\"").append(XlsxWorkbookSerializerCommon.colorArgb(c)).append("\"/>");
    }

    /**
     * Indicates whether empty color.
     * @param cv cv
     * @return true when the condition is satisfied
     */
    private static boolean isEmptyColor(ColorValue cv) {
        return cv == null || (cv.getA() == 0 && cv.getR() == 0 && cv.getG() == 0 && cv.getB() == 0);
    }

    /**
     * Builds the sqref.
     * @param areas areas
     * @return the requested result
     */
    private static String buildSqref(List<CellArea> areas) {
        StringBuilder sb = new StringBuilder();
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < areas.size(); i++) {
            if (i > 0) sb.append(" ");
            sb.append(toAreaReference(areas.get(i)));
        }
        return sb.toString();
    }

    /**
     * Processes to area reference.
     * @param area area
     * @return the computed result
     */
    private static String toAreaReference(CellArea area) {
        String first = XlsxWorkbookSerializerCommon.colLetter(area.getFirstColumn()) + (area.getFirstRow() + 1);
        // Handle the relevant branch before the state changes.
        if (area.getTotalRows() == 1 && area.getTotalColumns() == 1) return first;
        String last = XlsxWorkbookSerializerCommon.colLetter(area.getFirstColumn() + area.getTotalColumns() - 1)
                + (area.getFirstRow() + area.getTotalRows());
        return first + ":" + last;
    }

    /**
     * Processes get anchor cell.
     * @param cf cf
     * @return the requested result
     */
    private static String getAnchorCell(ConditionalFormattingModel cf) {
        // Handle the relevant branch before the state changes.
        if (cf.getAreas().isEmpty()) return "A1";
        CellArea a = cf.getAreas().get(0);
        return XlsxWorkbookSerializerCommon.colLetter(a.getFirstColumn()) + (a.getFirstRow() + 1);
    }

    /**
     * Builds the text rule formula.
     * @param type type
     * @param text text
     * @param cell cell
     * @return the requested result
     */
    private static String buildTextRuleFormula(FormatConditionType type, String text, String cell) {
        String escaped = text.replace("\"", "\"\"");
        // Translate the internal value into the matching public representation.
        switch (type) {
            case CONTAINS_TEXT:     return "NOT(ISERROR(SEARCH(\"" + escaped + "\"," + cell + ")))";
            case NOT_CONTAINS_TEXT: return "ISERROR(SEARCH(\"" + escaped + "\"," + cell + "))";
            case begins_with:       return "LEFT(" + cell + ",LEN(\"" + escaped + "\"))=\"" + escaped + "\"";
            case ENDS_WITH:         return "RIGHT(" + cell + ",LEN(\"" + escaped + "\"))=\"" + escaped + "\"";
            default:                return "";
        }
    }

    /**
     * Processes to rule type name.
     * @param cond cond
     * @return the computed result
     */
    private static String toRuleTypeName(FormatConditionModel cond) {
        // Translate the internal value into the matching public representation.
        switch (cond.getType()) {
            case EXPRESSION:        return "expression";
            case CONTAINS_TEXT:     return "containsText";
            case NOT_CONTAINS_TEXT: return "notContainsText";
            case begins_with:       return "beginsWith";
            case ENDS_WITH:         return "endsWith";
            case TIME_PERIOD:       return "timePeriod";
            case DUPLICATE_VALUES:  return "duplicateValues";
            case UNIQUE_VALUES:     return "uniqueValues";
            case TOP_10:
            case BOTTOM_10:         return "top10";
            case ABOVE_AVERAGE:
            case BELOW_AVERAGE:     return "aboveAverage";
            case COLOR_SCALE:       return "colorScale";
            case DATA_BAR:          return "dataBar";
            case ICON_SET:          return "iconSet";
            default:                return "cellIs";
        }
    }

    /**
     * Processes to operator name.
     * @param op op
     * @return the computed result
     */
    private static String toOperatorName(OperatorType op) {
        // Handle the relevant branch before the state changes.
        if (op == null) return null;
        switch (op) {
            case BETWEEN:          return "between";
            case NOT_BETWEEN:      return "notBetween";
            case EQUAL:            return "equal";
            case NOT_EQUAL:        return "notEqual";
            case GREATER_THAN:     return "greaterThan";
            case LESS_THAN:        return "lessThan";
            case GREATER_OR_EQUAL: return "greaterThanOrEqual";
            case LESS_OR_EQUAL:    return "lessThanOrEqual";
            default:               return null;
        }
    }

    /**
     * Processes get icon count.
     * @param iconSetType icon set type
     * @return the requested result
     */
    private static int getIconCount(String iconSetType) {
        // Handle the relevant branch before the state changes.
        if (iconSetType != null && !iconSetType.isEmpty()) {
            char first = iconSetType.charAt(0);
            if (first == '4') return 4;
            if (first == '5') return 5;
        }
        return 3;
    }

    // =========================================================================
    // Load
    // =========================================================================

    /**
     * Loads the conditional formattings.
     * @param ws ws
     * @param doc doc
     * @param dxfStyles dxf styles
     */
    static void loadConditionalFormattings(WorksheetModel ws, Document doc, List<StyleValue> dxfStyles) {
        ws.getConditionalFormattings().clear();
        NodeList cfNodes = doc.getElementsByTagNameNS("*", "conditionalFormatting");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < cfNodes.getLength(); i++) {
            Element cfEl = (Element) cfNodes.item(i);
            String sqref = cfEl.getAttribute("sqref");
            if (sqref == null || sqref.trim().isEmpty()) continue;

            ConditionalFormattingModel cf = new ConditionalFormattingModel();
            loadAreas(cf, sqref);
            if (cf.getAreas().isEmpty()) continue;

            NodeList ruleNodes = cfEl.getElementsByTagNameNS("*", "cfRule");
            for (int j = 0; j < ruleNodes.getLength(); j++) {
                Element ruleEl = (Element) ruleNodes.item(j);
                loadCondition(cf, ruleEl, dxfStyles);
            }

            if (cf.getConditions().isEmpty()) continue;
            ws.getConditionalFormattings().add(cf);
        }
    }

    /**
     * Loads the areas.
     * @param cf cf
     * @param sqref sqref
     */
    private static void loadAreas(ConditionalFormattingModel cf, String sqref) {
        // Walk the current collection so every entry is processed consistently.
        for (String token : sqref.trim().split("\\s+")) {
            if (token.isEmpty()) continue;
            String[] parts = token.split(":");
            CellAddress a1 = XlsxWorkbookSerializerCommon.parseRef(parts[0]);
            int r1 = a1.getRowIndex(), c1 = a1.getColumnIndex();
            int r2 = r1, c2 = c1;
            if (parts.length == 2) {
                CellAddress a2 = XlsxWorkbookSerializerCommon.parseRef(parts[1]);
                r2 = a2.getRowIndex(); c2 = a2.getColumnIndex();
            }
            cf.getAreas().add(new CellArea(r1, c1, r2 - r1 + 1, c2 - c1 + 1));
        }
        FormatConditionCollection.sortAreas(cf.getAreas());
    }

    /**
     * Loads the condition.
     * @param cf cf
     * @param ruleEl rule el
     * @param dxfStyles dxf styles
     */
    private static void loadCondition(ConditionalFormattingModel cf, Element ruleEl, List<StyleValue> dxfStyles) {
        String typeText = ruleEl.getAttribute("type");
        if (typeText == null || typeText.isEmpty()) return;

        FormatConditionType type = parseRuleType(typeText, ruleEl);
        if (type == null) return;

        FormatConditionModel cond = new FormatConditionModel();
        cond.setType(type);

        // priority
        String priStr = ruleEl.getAttribute("priority");
        int pri = XlsxWorkbookSerializerCommon.parseInt(priStr, 0);
        cond.setPriority(pri > 0 ? pri : cf.getConditions().size() + 1);

        // stopIfTrue
        cond.setStopIfTrue("1".equals(ruleEl.getAttribute("stopIfTrue")));

        // operator
        String opStr = ruleEl.getAttribute("operator");
        if (!opStr.isEmpty()) cond.setOperator(parseOperatorType(opStr));

        // differential style
        String dxfIdStr = ruleEl.getAttribute("dxfId");
        if (!dxfIdStr.isEmpty()) {
            int dxfId = XlsxWorkbookSerializerCommon.parseInt(dxfIdStr, -1);
            if (dxfId >= 0 && dxfId < dxfStyles.size()) {
                cond.setStyle(dxfStyles.get(dxfId).clone());
            }
        }

        // type-specific
        NodeList formulas = ruleEl.getElementsByTagNameNS("*", "formula");
        switch (type) {
            case EXPRESSION:
                if (formulas.getLength() > 0)
                    cond.setFormula1(formulas.item(0).getTextContent());
                break;
            case CELL_VALUE:
                if (formulas.getLength() > 0) cond.setFormula1(formulas.item(0).getTextContent());
                if (formulas.getLength() > 1) cond.setFormula2(formulas.item(1).getTextContent());
                break;
            case CONTAINS_TEXT:
            case NOT_CONTAINS_TEXT:
            case begins_with:
            case ENDS_WITH:
                cond.setFormula1(ruleEl.getAttribute("text"));
                break;
            case TIME_PERIOD:
                cond.setTimePeriod(ruleEl.getAttribute("timePeriod"));
                break;
            case TOP_10:
            case BOTTOM_10: {
                cond.setTop(type == FormatConditionType.TOP_10);
                cond.setPercent("1".equals(ruleEl.getAttribute("percent")));
                int rank = XlsxWorkbookSerializerCommon.parseInt(ruleEl.getAttribute("rank"), 0);
                cond.setRank(rank > 0 ? rank : 10);
                break;
            }
            case ABOVE_AVERAGE:
            case BELOW_AVERAGE:
                cond.setAbove(type == FormatConditionType.ABOVE_AVERAGE);
                cond.setStandardDeviation(XlsxWorkbookSerializerCommon.parseInt(ruleEl.getAttribute("stdDev"), 0));
                break;
            case DUPLICATE_VALUES:
                cond.setDuplicate(true);
                break;
            case UNIQUE_VALUES:
                cond.setDuplicate(false);
                break;
            case COLOR_SCALE: {
                NodeList csEl = ruleEl.getElementsByTagNameNS("*", "colorScale");
                if (csEl.getLength() > 0) loadColorScale(cond, (Element) csEl.item(0));
                break;
            }
            case DATA_BAR: {
                NodeList dbEl = ruleEl.getElementsByTagNameNS("*", "dataBar");
                if (dbEl.getLength() > 0) loadDataBar(cond, (Element) dbEl.item(0));
                break;
            }
            case ICON_SET: {
                NodeList isEl = ruleEl.getElementsByTagNameNS("*", "iconSet");
                if (isEl.getLength() > 0) loadIconSet(cond, (Element) isEl.item(0));
                break;
            }
            default:
                break;
        }

        // apply defaults for types that need them
        applyDefaults(cond);
        cf.getConditions().add(cond);
    }

    /**
     * Loads the color scale.
     * @param cond cond
     * @param el el
     */
    private static void loadColorScale(FormatConditionModel cond, Element el) {
        NodeList cfvos = el.getElementsByTagNameNS("*", "cfvo");
        cond.setColorScaleCount(cfvos.getLength() >= 3 ? 3 : 2);
        NodeList colors = el.getElementsByTagNameNS("*", "color");
        // Handle the relevant branch before the state changes.
        if (colors.getLength() > 0) cond.setMinColor(readRgbColor((Element) colors.item(0)));
        if (cond.getColorScaleCount() == 3 && colors.getLength() > 2) {
            cond.setMidColor(readRgbColor((Element) colors.item(1)));
            cond.setMaxColor(readRgbColor((Element) colors.item(2)));
        } else if (colors.getLength() > 1) {
            cond.setMaxColor(readRgbColor((Element) colors.item(1)));
        }
    }

    /**
     * Loads the data bar.
     * @param cond cond
     * @param el el
     */
    private static void loadDataBar(FormatConditionModel cond, Element el) {
        NodeList colors = el.getElementsByTagNameNS("*", "color");
        // Handle the relevant branch before the state changes.
        if (colors.getLength() > 0) cond.setBarColor(readRgbColor((Element) colors.item(0)));
    }

    /**
     * Loads the icon set.
     * @param cond cond
     * @param el el
     */
    private static void loadIconSet(FormatConditionModel cond, Element el) {
        String iconSet = el.getAttribute("iconSet");
        cond.setIconSetType(iconSet.isEmpty() ? "3TrafficLights1" : iconSet);
        cond.setReverseIcons("1".equals(el.getAttribute("reverse")));
        // showValue="0" means showIconOnly=true
        String sv = el.getAttribute("showValue");
        cond.setShowIconOnly(!sv.isEmpty() && "0".equals(sv));
    }

    /**
     * Reads rgb color.
     * @param el el
     * @return the computed result
     */
    private static ColorValue readRgbColor(Element el) {
        // Handle the relevant branch before the state changes.
        if (el == null) return null;
        String rgb = el.getAttribute("rgb");
        if (rgb.length() < 8) return null;
        try {
            int a = Integer.parseInt(rgb.substring(0, 2), 16);
            int r = Integer.parseInt(rgb.substring(2, 4), 16);
            int g = Integer.parseInt(rgb.substring(4, 6), 16);
            int b = Integer.parseInt(rgb.substring(6, 8), 16);
            return new ColorValue((byte) a, (byte) r, (byte) g, (byte) b);
        } catch (NumberFormatException e) { return null; }
    }

    /**
     * Applies defaults.
     * @param cond cond
     */
    private static void applyDefaults(FormatConditionModel cond) {
        // Translate the internal value into the matching public representation.
        switch (cond.getType()) {
            case COLOR_SCALE:
                if (cond.getColorScaleCount() == 0) cond.setColorScaleCount(2);
                break;
            case DATA_BAR:
                if (isEmptyColor(cond.getBarColor()))
                    cond.setBarColor(new ColorValue((byte)255,(byte)99,(byte)142,(byte)198));
                break;
            case ICON_SET:
                if (cond.getIconSetType() == null || cond.getIconSetType().isEmpty())
                    cond.setIconSetType("3TrafficLights1");
                break;
            default:
                break;
        }
    }

    /**
     * Parses rule type.
     * @param typeText type text
     * @param ruleEl rule el
     * @return the computed result
     */
    private static FormatConditionType parseRuleType(String typeText, Element ruleEl) {
        // Translate the internal value into the matching public representation.
        switch (typeText.toLowerCase(java.util.Locale.ROOT)) {
            case "cellis":          return FormatConditionType.CELL_VALUE;
            case "expression":      return FormatConditionType.EXPRESSION;
            case "containstext":    return FormatConditionType.CONTAINS_TEXT;
            case "notcontainstext": return FormatConditionType.NOT_CONTAINS_TEXT;
            case "beginswith":      return FormatConditionType.begins_with;
            case "endswith":        return FormatConditionType.ENDS_WITH;
            case "timeperiod":      return FormatConditionType.TIME_PERIOD;
            case "duplicatevalues": return FormatConditionType.DUPLICATE_VALUES;
            case "uniquevalues":    return FormatConditionType.UNIQUE_VALUES;
            case "top10":
                return "1".equals(ruleEl.getAttribute("bottom"))
                        ? FormatConditionType.BOTTOM_10 : FormatConditionType.TOP_10;
            case "aboveaverage": {
                String above = ruleEl.getAttribute("aboveAverage");
                return (!above.isEmpty() && "0".equals(above))
                        ? FormatConditionType.BELOW_AVERAGE : FormatConditionType.ABOVE_AVERAGE;
            }
            case "colorscale":      return FormatConditionType.COLOR_SCALE;
            case "databar":         return FormatConditionType.DATA_BAR;
            case "iconset":         return FormatConditionType.ICON_SET;
            default:                return null;
        }
    }

    /**
     * Parses operator type.
     * @param s s
     * @return the computed result
     */
    private static OperatorType parseOperatorType(String s) {
        // Translate the internal value into the matching public representation.
        switch (s.toLowerCase(java.util.Locale.ROOT)) {
            case "between":             return OperatorType.BETWEEN;
            case "notbetween":          return OperatorType.NOT_BETWEEN;
            case "equal":               return OperatorType.EQUAL;
            case "notequal":            return OperatorType.NOT_EQUAL;
            case "greaterthan":         return OperatorType.GREATER_THAN;
            case "lessthan":            return OperatorType.LESS_THAN;
            case "greaterthanorequal":  return OperatorType.GREATER_OR_EQUAL;
            case "lessthanorequal":     return OperatorType.LESS_OR_EQUAL;
            default:                    return OperatorType.NONE;
        }
    }
}
