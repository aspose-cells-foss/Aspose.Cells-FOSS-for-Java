package org.aspose.cells_foss;

import java.nio.charset.StandardCharsets;

/** Generates minimal OOXML chart XML for each supported chart type. */
final class XlsxWorkbookChartTemplates {

    private XlsxWorkbookChartTemplates() {}

    private static final String C   = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static final String A   = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /** Returns the chart XML bytes for the given type, with {{DATA}} replaced by dataRange. */
    static byte[] build(ChartType type, String dataRange) {
        String inner = innerXml(type, dataRange);
        // CT_Chart requires view3D as a direct child of <c:chart>, before <c:plotArea>.
        String view3d = needsView3D(type) ? view3d() : "";
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                + "<c:chartSpace xmlns:c=\"" + C + "\" xmlns:a=\"" + A + "\" xmlns:r=\"" + R + "\">"
                + "<c:chart>"
                + "<c:autoTitleDeleted val=\"1\"/>"
                + view3d
                + "<c:plotArea>"
                + inner
                + "</c:plotArea>"
                + "<c:legend><c:legendPos val=\"b\"/></c:legend>"
                + "<c:plotVisOnly val=\"1\"/>"
                + "</c:chart>"
                + "</c:chartSpace>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    private static boolean needsView3D(ChartType type) {
        return switch (type) {
            case COLUMN_3D, BAR_3D, LINE_3D, AREA_3D, PIE_3D,
                 SURFACE_3D, SURFACE_WIREFRAME_3D -> true;
            default -> false;
        };
    }

    private static String innerXml(ChartType type, String data) {
        String s = ser1(data);
        String ref2  = axRef("1") + axRef("2");
        String ref3  = axRef("1") + axRef("2") + axRef("3");
        // Column/line/area: category axis at bottom, value axis at left
        String ax2d    = catValAxes("1", "2", "b", "l");
        // Horizontal bar: category axis at left, value axis at bottom
        String ax2dBar = catValAxes("1", "2", "l", "b");
        // 3D column/line/area: same as 2D column orientation
        String ax3d    = catValSerAxes("1", "2", "3", "b", "l");
        // 3D horizontal bar: category at left, value at bottom
        String ax3dBar = catValSerAxes("1", "2", "3", "l", "b");
        return switch (type) {
            case COLUMN  -> "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" + s + ref2 + "</c:barChart>" + ax2d;
            case BAR     -> "<c:barChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>" + s + ref2 + "</c:barChart>" + ax2dBar;
            case LINE    -> "<c:lineChart><c:grouping val=\"standard\"/>" + s + ref2 + "</c:lineChart>" + ax2d;
            case AREA    -> "<c:areaChart><c:grouping val=\"standard\"/>" + s + ref2 + "</c:areaChart>" + ax2d;
            case PIE     -> "<c:pieChart>" + s + "</c:pieChart>";
            case DOUGHNUT-> "<c:doughnutChart><c:holeSize val=\"50\"/>" + s + "</c:doughnutChart>";
            case SCATTER -> "<c:scatterChart><c:scatterStyle val=\"line\"/>" + serScatter(data) + ref2 + "</c:scatterChart>" + ax2d;
            case BUBBLE  -> "<c:bubbleChart>" + serBubble(data) + ref2 + "</c:bubbleChart>" + ax2d;
            case RADAR   -> "<c:radarChart><c:radarStyle val=\"standard\"/>" + s + ref2 + "</c:radarChart>" + ax2d;
            case STOCK   -> "<c:stockChart>" + serStock(data) + ref2 + "</c:stockChart>" + ax2d;
            case COLUMN_3D -> "<c:bar3DChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" + s + "<c:shape val=\"box\"/>" + ref3 + "</c:bar3DChart>" + ax3d;
            case BAR_3D    -> "<c:bar3DChart><c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>" + s + "<c:shape val=\"box\"/>" + ref3 + "</c:bar3DChart>" + ax3dBar;
            case LINE_3D   -> "<c:line3DChart><c:grouping val=\"standard\"/>" + s + ref3 + "</c:line3DChart>" + ax3d;
            case AREA_3D   -> "<c:area3DChart><c:grouping val=\"standard\"/>" + s + ref3 + "</c:area3DChart>" + ax3d;
            case PIE_3D    -> "<c:pie3DChart>" + s + "</c:pie3DChart>";
            case SURFACE_3D          -> "<c:surface3DChart><c:wireframe val=\"0\"/>" + s + ref3 + "</c:surface3DChart>" + ax3d;
            case SURFACE_WIREFRAME_3D-> "<c:surface3DChart><c:wireframe val=\"1\"/>" + s + ref3 + "</c:surface3DChart>" + ax3d;
            case CONTOUR   -> "<c:surfaceChart><c:wireframe val=\"0\"/>" + s + ref3 + "</c:surfaceChart>" + ax3d;
            default -> "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" + s + ref2 + "</c:barChart>" + ax2d;
        };
    }

    private static String axRef(String id) {
        return "<c:axId val=\"" + id + "\"/>";
    }

    private static String ser1(String data) {
        return "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
             + "<c:val><c:numRef><c:f>" + xmlEsc(data) + "</c:f></c:numRef></c:val>"
             + "</c:ser>";
    }

    private static String serScatter(String data) {
        return "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
             + "<c:xVal><c:numRef><c:f>" + xmlEsc(data) + "</c:f></c:numRef></c:xVal>"
             + "<c:yVal><c:numRef><c:f>" + xmlEsc(data) + "</c:f></c:numRef></c:yVal>"
             + "</c:ser>";
    }

    private static String serBubble(String data) {
        return "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
             + "<c:xVal><c:numRef><c:f>" + xmlEsc(data) + "</c:f></c:numRef></c:xVal>"
             + "<c:yVal><c:numRef><c:f>" + xmlEsc(data) + "</c:f></c:numRef></c:yVal>"
             + "<c:bubbleSize><c:numRef><c:f>" + xmlEsc(data) + "</c:f></c:numRef></c:bubbleSize>"
             + "</c:ser>";
    }

    private static String serStock(String data) {
        StringBuilder sb = new StringBuilder();
        String[] names = {"Open","High","Low","Close"};
        for (int i = 0; i < 4; i++) {
            sb.append("<c:ser><c:idx val=\"").append(i).append("\"/><c:order val=\"").append(i).append("\"/>")
              .append("<c:tx><c:strRef><c:f/><c:strCache><c:ptCount val=\"1\"/>")
              .append("<c:pt idx=\"0\"><c:v>").append(names[i]).append("</c:v></c:pt>")
              .append("</c:strCache></c:strRef></c:tx>")
              .append("<c:val><c:numRef><c:f>").append(xmlEsc(data)).append("</c:f></c:numRef></c:val>")
              .append("</c:ser>");
        }
        return sb.toString();
    }

    private static String catValAxes(String catId, String valId, String catPos, String valPos) {
        return ax("catAx", catId, valId, catPos)
             + ax("valAx", valId, catId, valPos);
    }

    private static String catValSerAxes(String catId, String valId, String serId, String catPos, String valPos) {
        return ax("catAx", catId, valId, catPos)
             + ax("valAx", valId, catId, valPos)
             + "<c:serAx><c:axId val=\"" + serId + "\"/>"
             + "<c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
             + "<c:delete val=\"1\"/><c:axPos val=\"b\"/><c:crossAx val=\"" + valId + "\"/></c:serAx>";
    }

    private static String ax(String tag, String id, String crossId, String pos) {
        return "<c:" + tag + "><c:axId val=\"" + id + "\"/>"
             + "<c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
             + "<c:delete val=\"0\"/><c:axPos val=\"" + pos + "\"/>"
             + "<c:crossAx val=\"" + crossId + "\"/></c:" + tag + ">";
    }

    private static String view3d() {
        return "<c:view3D><c:rotX val=\"15\"/><c:rotY val=\"20\"/><c:perspective val=\"30\"/></c:view3D>";
    }

    private static String xmlEsc(String s) {
        if (s == null) return "";
        return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\"","&quot;");
    }

    // ---- Type detection from raw chart XML ----

    static String detectChartType(String rawXml) {
        if (rawXml == null) return "unknown";
        if (rawXml.contains("c:bar3DChart")) {
            return rawXml.contains("barDir val=\"bar\"") ? "bar3d" : "column3d";
        }
        if (rawXml.contains("c:barChart")) {
            return rawXml.contains("barDir val=\"bar\"") ? "bar" : "column";
        }
        if (rawXml.contains("c:line3DChart")) return "line3d";
        if (rawXml.contains("c:lineChart")) return "line";
        if (rawXml.contains("c:area3DChart")) return "area3d";
        if (rawXml.contains("c:areaChart")) return "area";
        if (rawXml.contains("c:pie3DChart")) return "pie3d";
        if (rawXml.contains("c:pieChart")) return "pie";
        if (rawXml.contains("c:doughnutChart")) return "doughnut";
        if (rawXml.contains("c:scatterChart")) return "scatter";
        if (rawXml.contains("c:bubbleChart")) return "bubble";
        if (rawXml.contains("c:radarChart")) return "radar";
        if (rawXml.contains("c:stockChart")) return "stock";
        if (rawXml.contains("c:surface3DChart")) {
            return rawXml.contains("wireframe val=\"1\"") ? "surfacewireframe3d" : "surface3d";
        }
        if (rawXml.contains("c:surfaceChart")) return "contour";
        return "unknown";
    }
}

