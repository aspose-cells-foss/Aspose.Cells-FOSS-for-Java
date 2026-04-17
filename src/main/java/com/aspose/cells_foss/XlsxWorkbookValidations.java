package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;

/**
 * Package-private helpers for building and loading XLSX data validations.
 */
final class XlsxWorkbookValidations {

    /**
     * Initializes a new XlsxWorkbookValidations instance.
     */
    private XlsxWorkbookValidations() {}

    // =========================================================================
    // Build
    // =========================================================================

    /**
     * Builds the data validations section.
     * @param ws ws
     * @param sb sb
     */
    static void buildDataValidationsSection(WorksheetModel ws, StringBuilder sb) {
        if (ws.getValidations().isEmpty()) return;

        sb.append("<dataValidations count=\"").append(ws.getValidations().size()).append("\">");
        for (ValidationModel dv : ws.getValidations()) {
            sb.append("<dataValidation");
            // type
            String dvType = validationTypeToXml(dv.getType());
            if (!dvType.equals("none")) sb.append(" type=\"").append(dvType).append("\"");
            // errorStyle
            String errStyle = alertStyleToXml(dv.getAlertStyle());
            if (!errStyle.equals("stop")) sb.append(" errorStyle=\"").append(errStyle).append("\"");
            // operator (omit for list/custom/none)
            if (dv.getType() != ValidationType.LIST && dv.getType() != ValidationType.CUSTOM
                    && dv.getType() != null && dv.getType() != ValidationType.ANY_VALUE
                    && dv.getOperator() != null && dv.getOperator() != OperatorType.NONE) {
                sb.append(" operator=\"").append(operatorTypeToXml(dv.getOperator())).append("\"");
            }
            if (dv.getIgnoreBlank()) sb.append(" allowBlank=\"1\"");
            if (!dv.getInCellDropDown()) sb.append(" showDropDown=\"1\"");
            if (dv.getShowInput()) sb.append(" showInputMessage=\"1\"");
            if (dv.getShowError()) sb.append(" showErrorMessage=\"1\"");
            if (dv.getErrorTitle() != null && !dv.getErrorTitle().isEmpty())
                sb.append(" errorTitle=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(dv.getErrorTitle())).append("\"");
            if (dv.getErrorMessage() != null && !dv.getErrorMessage().isEmpty())
                sb.append(" error=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(dv.getErrorMessage())).append("\"");
            if (dv.getInputTitle() != null && !dv.getInputTitle().isEmpty())
                sb.append(" promptTitle=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(dv.getInputTitle())).append("\"");
            if (dv.getInputMessage() != null && !dv.getInputMessage().isEmpty())
                sb.append(" prompt=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(dv.getInputMessage())).append("\"");
            // sqref
            sb.append(" sqref=\"");
            boolean firstArea = true;
            for (CellArea area : dv.getAreas()) {
                if (!firstArea) sb.append(" ");
                firstArea = false;
                int r1 = area.getFirstRow(), c1 = area.getFirstColumn();
                int r2 = r1 + area.getTotalRows() - 1, c2 = c1 + area.getTotalColumns() - 1;
                sb.append(XlsxWorkbookSerializerCommon.colLetter(c1)).append(r1 + 1);
                if (r2 != r1 || c2 != c1)
                    sb.append(":").append(XlsxWorkbookSerializerCommon.colLetter(c2)).append(r2 + 1);
            }
            sb.append("\">");
            if (dv.getFormula1() != null && !dv.getFormula1().isEmpty())
                sb.append("<formula1>").append(XlsxWorkbookSerializerCommon.xmlText(dv.getFormula1())).append("</formula1>");
            if (dv.getFormula2() != null && !dv.getFormula2().isEmpty())
                sb.append("<formula2>").append(XlsxWorkbookSerializerCommon.xmlText(dv.getFormula2())).append("</formula2>");
            sb.append("</dataValidation>");
        }
        sb.append("</dataValidations>");
    }

    // =========================================================================
    // Load
    // =========================================================================

    /**
     * Loads the validations.
     * @param ws ws
     * @param doc doc
     */
    static void loadValidations(WorksheetModel ws, Document doc) {
        NodeList dvParent = doc.getElementsByTagNameNS("*", "dataValidations");
        if (dvParent.getLength() == 0) return;

        NodeList dvNodes = ((Element) dvParent.item(0)).getElementsByTagNameNS("*", "dataValidation");
        for (int i = 0; i < dvNodes.getLength(); i++) {
            Element dv = (Element) dvNodes.item(i);
            ValidationModel vm = new ValidationModel();
            vm.setType(xmlToValidationType(dv.getAttribute("type")));
            String errStyle = dv.getAttribute("errorStyle");
            if (!errStyle.isEmpty()) vm.setAlertStyle(xmlToAlertStyle(errStyle));
            String op = dv.getAttribute("operator");
            vm.setOperator(op.isEmpty() ? OperatorType.NONE : xmlToOperatorType(op));
            vm.setIgnoreBlank("1".equals(dv.getAttribute("allowBlank")));
            vm.setInCellDropDown(!"1".equals(dv.getAttribute("showDropDown")));
            vm.setShowInput("1".equals(dv.getAttribute("showInputMessage")));
            vm.setShowError("1".equals(dv.getAttribute("showErrorMessage")));
            String promptTitle = dv.getAttribute("promptTitle");
            if (!promptTitle.isEmpty()) vm.setInputTitle(promptTitle);
            String prompt = dv.getAttribute("prompt");
            if (!prompt.isEmpty()) vm.setInputMessage(prompt);
            String errorTitle = dv.getAttribute("errorTitle");
            if (!errorTitle.isEmpty()) vm.setErrorTitle(errorTitle);
            String error = dv.getAttribute("error");
            if (!error.isEmpty()) vm.setErrorMessage(error);
            NodeList f1 = dv.getElementsByTagNameNS("*", "formula1");
            if (f1.getLength() > 0) vm.setFormula1(f1.item(0).getTextContent());
            NodeList f2 = dv.getElementsByTagNameNS("*", "formula2");
            if (f2.getLength() > 0) vm.setFormula2(f2.item(0).getTextContent());
            // sqref
            String sqref = dv.getAttribute("sqref");
            for (String rangeStr : sqref.trim().split("\\s+")) {
                if (rangeStr.isEmpty()) continue;
                String[] parts = rangeStr.split(":");
                CellAddress ca1 = XlsxWorkbookSerializerCommon.parseRef(parts[0]);
                int r1 = ca1.getRowIndex(), c1 = ca1.getColumnIndex();
                int r2 = r1, c2 = c1;
                if (parts.length == 2) {
                    CellAddress ca2 = XlsxWorkbookSerializerCommon.parseRef(parts[1]);
                    r2 = ca2.getRowIndex();
                    c2 = ca2.getColumnIndex();
                }
                vm.getAreas().add(new CellArea(r1, c1, r2 - r1 + 1, c2 - c1 + 1));
            }
            ws.getValidations().add(vm);
        }
    }

    // =========================================================================
    // Enum converters
    // =========================================================================

    /**
     * Processes validation type to xml.
     * @param t t
     * @return the computed result
     */
    static String validationTypeToXml(ValidationType t) {
        // Handle the relevant branch before the state changes.
        if (t == null) return "none";
        switch (t) {
            case WHOLE_NUMBER: return "whole";
            case DECIMAL:      return "decimal";
            case LIST:         return "list";
            case DATE:         return "date";
            case TIME:         return "time";
            case TEXT_LENGTH:  return "textLength";
            case CUSTOM:       return "custom";
            default:           return "none";
        }
    }

    /**
     * Processes xml to validation type.
     * @param s s
     * @return the computed result
     */
    static ValidationType xmlToValidationType(String s) {
        // Handle the relevant branch before the state changes.
        if (s == null) return ValidationType.ANY_VALUE;
        switch (s) {
            case "whole":       return ValidationType.WHOLE_NUMBER;
            case "decimal":     return ValidationType.DECIMAL;
            case "list":        return ValidationType.LIST;
            case "date":        return ValidationType.DATE;
            case "time":        return ValidationType.TIME;
            case "textLength":  return ValidationType.TEXT_LENGTH;
            case "custom":      return ValidationType.CUSTOM;
            default:            return ValidationType.ANY_VALUE;
        }
    }

    /**
     * Processes alert style to xml.
     * @param a a
     * @return the computed result
     */
    static String alertStyleToXml(ValidationAlertType a) {
        // Handle the relevant branch before the state changes.
        if (a == null) return "stop";
        switch (a) {
            case WARNING:     return "warning";
            case INFORMATION: return "information";
            default:          return "stop";
        }
    }

    /**
     * Processes xml to alert style.
     * @param s s
     * @return the computed result
     */
    static ValidationAlertType xmlToAlertStyle(String s) {
        // Handle the relevant branch before the state changes.
        if (s == null) return ValidationAlertType.STOP;
        switch (s) {
            case "warning":     return ValidationAlertType.WARNING;
            case "information": return ValidationAlertType.INFORMATION;
            default:            return ValidationAlertType.STOP;
        }
    }

    /**
     * Processes operator type to xml.
     * @param op op
     * @return the computed result
     */
    static String operatorTypeToXml(OperatorType op) {
        // Handle the relevant branch before the state changes.
        if (op == null) return "between";
        switch (op) {
            case NOT_BETWEEN:      return "notBetween";
            case EQUAL:            return "equal";
            case NOT_EQUAL:        return "notEqual";
            case GREATER_THAN:     return "greaterThan";
            case LESS_THAN:        return "lessThan";
            case GREATER_OR_EQUAL: return "greaterThanOrEqual";
            case LESS_OR_EQUAL:    return "lessThanOrEqual";
            default:               return "between";
        }
    }

    /**
     * Processes xml to operator type.
     * @param s s
     * @return the computed result
     */
    static OperatorType xmlToOperatorType(String s) {
        // Handle the relevant branch before the state changes.
        if (s == null || s.isEmpty()) return OperatorType.BETWEEN;
        switch (s) {
            case "notBetween":         return OperatorType.NOT_BETWEEN;
            case "equal":              return OperatorType.EQUAL;
            case "notEqual":           return OperatorType.NOT_EQUAL;
            case "greaterThan":        return OperatorType.GREATER_THAN;
            case "lessThan":           return OperatorType.LESS_THAN;
            case "greaterThanOrEqual": return OperatorType.GREATER_OR_EQUAL;
            case "lessThanOrEqual":    return OperatorType.LESS_OR_EQUAL;
            default:                   return OperatorType.BETWEEN;
        }
    }
}
