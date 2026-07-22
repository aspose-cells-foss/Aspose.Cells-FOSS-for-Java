package org.aspose.cells_foss;

import org.aspose.cells_foss.core.CommentModel;
import org.aspose.cells_foss.core.WorksheetModel;
import org.w3c.dom.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * Builds and loads XLSX comment parts (xl/comments{N}.xml and vmlDrawing{N}.vml).
 */
final class XlsxWorkbookComments {

    private XlsxWorkbookComments() {}

    // =========================================================================
    // Build
    // =========================================================================

    /** Returns the XML bytes for xl/comments{N}.xml, or null if there are no comments. */
    static byte[] buildCommentsXml(WorksheetModel ws) {
        List<CommentModel> comments = ws.getComments();
        if (comments.isEmpty()) return null;

        List<String> authors = buildAuthorList(comments);
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        sb.append("<authors>");
        for (String a : authors) {
            sb.append("<author>").append(XlsxWorkbookSerializerCommon.xmlText(a)).append("</author>");
        }
        sb.append("</authors>");
        sb.append("<commentList>");
        for (CommentModel c : comments) {
            String ref = XlsxWorkbookSerializerCommon.colLetter(c.getColumn()) + (c.getRow() + 1);
            int authorIdx = authors.indexOf(c.getAuthor());
            sb.append("<comment ref=\"").append(ref).append("\" authorId=\"").append(authorIdx).append("\">");
            sb.append("<text><r><t xml:space=\"preserve\">")
              .append(XlsxWorkbookSerializerCommon.xmlText(c.getNote()))
              .append("</t></r></text>");
            sb.append("</comment>");
        }
        sb.append("</commentList>");
        sb.append("</comments>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    /** Returns the VML bytes for xl/drawings/vmlDrawing{N}.vml, or null if there are no comments. */
    static byte[] buildVmlDrawingXml(WorksheetModel ws) {
        List<CommentModel> comments = ws.getComments();
        if (comments.isEmpty()) return null;

        StringBuilder sb = new StringBuilder();
        sb.append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"");
        sb.append(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
        sb.append(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
        sb.append("<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>");
        sb.append("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"");
        sb.append(" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/>");
        sb.append("<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>");

        int shapeId = 1025;
        for (CommentModel c : comments) {
            if (c.getRawVmlShapeXml() != null) {
                sb.append(c.getRawVmlShapeXml());
            } else {
                sb.append(buildDefaultVmlShape(c, shapeId));
            }
            shapeId++;
        }
        sb.append("</xml>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    // =========================================================================
    // Load
    // =========================================================================

    static void loadComments(WorksheetModel ws, Document commentsDoc) {
        if (commentsDoc == null) return;
        NodeList authorNodes = commentsDoc.getElementsByTagNameNS("*", "author");
        List<String> authors = new ArrayList<>();
        for (int i = 0; i < authorNodes.getLength(); i++) {
            authors.add(authorNodes.item(i).getTextContent());
        }
        NodeList commentNodes = commentsDoc.getElementsByTagNameNS("*", "comment");
        for (int i = 0; i < commentNodes.getLength(); i++) {
            Element el = (Element) commentNodes.item(i);
            String ref = el.getAttribute("ref");
            if (ref.isBlank()) continue;
            CommentModel cm = new CommentModel();
            try {
                org.aspose.cells_foss.core.CellAddress addr = XlsxWorkbookSerializerCommon.parseRef(ref);
                cm.setRow(addr.getRowIndex());
                cm.setColumn(addr.getColumnIndex());
            } catch (Exception e) {
                continue;
            }
            String authorIdStr = el.getAttribute("authorId");
            if (!authorIdStr.isBlank()) {
                try {
                    int idx = Integer.parseInt(authorIdStr);
                    if (idx >= 0 && idx < authors.size()) cm.setAuthor(authors.get(idx));
                } catch (NumberFormatException ignored) {}
            }
            StringBuilder note = new StringBuilder();
            NodeList tNodes = el.getElementsByTagNameNS("*", "t");
            for (int j = 0; j < tNodes.getLength(); j++) note.append(tNodes.item(j).getTextContent());
            cm.setNote(note.toString());
            ws.getComments().add(cm);
        }
    }

    /** Parses the VML drawing to update isVisible, width, height on already-loaded comments. */
    static void loadVmlVisibility(WorksheetModel ws, byte[] vmlBytes) {
        if (vmlBytes == null) return;
        try {
            Document vml = XlsxWorkbookArchiveHelpers.parse(vmlBytes);
            if (vml == null) return;
            NodeList shapes = vml.getElementsByTagNameNS("*", "shape");
            for (int i = 0; i < shapes.getLength(); i++) {
                Element shape = (Element) shapes.item(i);
                NodeList clientDataNodes = shape.getElementsByTagNameNS("*", "ClientData");
                if (clientDataNodes.getLength() == 0) continue;
                Element cd = (Element) clientDataNodes.item(0);
                if (!"Note".equals(cd.getAttribute("ObjectType"))) continue;
                // Read row and column
                NodeList rowNodes = cd.getElementsByTagNameNS("*", "Row");
                NodeList colNodes = cd.getElementsByTagNameNS("*", "Column");
                if (rowNodes.getLength() == 0 || colNodes.getLength() == 0) continue;
                int row, col;
                try {
                    row = Integer.parseInt(rowNodes.item(0).getTextContent().trim());
                    col = Integer.parseInt(colNodes.item(0).getTextContent().trim());
                } catch (NumberFormatException e) { continue; }
                // Find matching comment
                CommentModel cm = null;
                for (CommentModel c : ws.getComments()) {
                    if (c.getRow() == row && c.getColumn() == col) { cm = c; break; }
                }
                if (cm == null) continue;
                // Visible flag: presence of <x:Visible/>
                boolean visible = cd.getElementsByTagNameNS("*", "Visible").getLength() > 0;
                cm.setVisible(visible);
                // Width and height from style attribute
                String style = shape.getAttribute("style");
                if (style != null && !style.isEmpty()) {
                    java.util.regex.Matcher wm = java.util.regex.Pattern.compile("width:(\\d+(?:\\.\\d+)?)pt").matcher(style);
                    java.util.regex.Matcher hm = java.util.regex.Pattern.compile("height:(\\d+(?:\\.\\d+)?)pt").matcher(style);
                    if (wm.find()) { try { cm.setWidth((int) Double.parseDouble(wm.group(1))); } catch (NumberFormatException ignore) {} }
                    if (hm.find()) { try { cm.setHeight((int) Double.parseDouble(hm.group(1))); } catch (NumberFormatException ignore) {} }
                }
            }
        } catch (Exception ignored) {}
    }

    // =========================================================================
    // Private helpers
    // =========================================================================

    private static List<String> buildAuthorList(List<CommentModel> comments) {
        List<String> authors = new ArrayList<>();
        for (CommentModel c : comments) {
            if (!authors.contains(c.getAuthor())) authors.add(c.getAuthor());
        }
        return authors;
    }

    private static String buildDefaultVmlShape(CommentModel c, int shapeId) {
        int leftPx = (c.getColumn() + 1) * 64 + 10;
        int topPx  = c.getRow() * 20 + 10;
        String visibility = c.isVisible() ? "visible" : "hidden";
        String col1 = String.valueOf(c.getColumn() + 1);
        String row1 = String.valueOf(c.getRow());
        String col2 = String.valueOf(c.getColumn() + 3);
        String row2 = String.valueOf(c.getRow() + 4);
        return "<v:shape id=\"_x0000_s" + shapeId + "\" type=\"#_x0000_t202\""
             + " style=\"position:absolute;margin-left:" + leftPx + "pt;margin-top:" + topPx + "pt;"
             + "width:" + c.getWidth() + "pt;height:" + c.getHeight() + "pt;z-index:1;"
             + "visibility:" + visibility + "\""
             + " fillcolor=\"#ffffe1\" o:insetmode=\"auto\">"
             + "<v:fill color2=\"#ffffe1\"/>"
             + "<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>"
             + "<v:path o:connecttype=\"none\"/>"
             + "<v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox>"
             + "<x:ClientData ObjectType=\"Note\">"
             + "<x:MoveWithCells/><x:SizeWithCells/>"
             + "<x:Anchor>" + col1 + ",15," + row1 + ",10," + col2 + ",15," + row2 + ",4</x:Anchor>"
             + "<x:AutoFill>False</x:AutoFill>"
             + "<x:Row>" + c.getRow() + "</x:Row>"
             + "<x:Column>" + c.getColumn() + "</x:Column>"
             + (c.isVisible() ? "<x:Visible/>" : "")
             + "</x:ClientData></v:shape>";
    }
}

