package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ChartModel;
import com.aspose.cells_foss.core.PictureModel;
import com.aspose.cells_foss.core.ShapeModel;
import com.aspose.cells_foss.core.WorksheetModel;
import org.w3c.dom.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * Builds and loads XLSX drawing parts (xl/drawings/drawing{N}.xml) covering pictures and charts.
 */
final class XlsxWorkbookDrawings {

    private static final String XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static final String A   = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String C   = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static final String R_NS= "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    static final String CHART_REL_TYPE    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    static final String CHARTEX_REL_TYPE  = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
    static final String IMAGE_REL_TYPE    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    static final String DRAWING_REL_TYPE  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    private static final String CHARTSTYLE_REL_TYPE  = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
    private static final String CHARTCOLORS_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

    private XlsxWorkbookDrawings() {}

    // =========================================================================
    // Build drawing XML
    // =========================================================================

    /**
     * Returns true if this worksheet has any pictures or charts.
     */
    static boolean hasDrawings(WorksheetModel ws) {
        return !ws.getPictures().isEmpty() || !ws.getShapes().isEmpty() || !ws.getCharts().isEmpty();
    }

    /**
     * Builds the drawing XML bytes for xl/drawings/drawing{N}.xml.
     * Populates drawingRels with [rId, type, target] entries for pictures and charts.
     *
     * @param ws           worksheet model
     * @param chartOffset  zero-based index of the first chart (across all sheets)
     * @param imageOffset  zero-based index of the first media image (across all sheets)
     * @param drawingRels  output: relationship entries for this drawing's rels file
     * @param chartXmls    output: Map from "xl/charts/chart{N}.xml" → bytes to write
     * @param mediaFiles   output: Map from "xl/media/image{N}.{ext}" → bytes to write
     */
    static byte[] buildDrawingXml(WorksheetModel ws, int chartOffset, int imageOffset,
                                   List<String[]> drawingRels,
                                   Map<String, byte[]> chartXmls,
                                   Map<String, byte[]> mediaFiles) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        boolean needsRNs = !ws.getPictures().isEmpty() || !ws.getCharts().isEmpty()
            || ws.getShapes().stream().anyMatch(s -> !s.getExtraDrawingRels().isEmpty());
        sb.append("<xdr:wsDr xmlns:xdr=\"").append(XDR).append("\"")
          .append(" xmlns:a=\"").append(A).append("\"");
        if (needsRNs) sb.append(" xmlns:r=\"").append(R_NS).append("\"");
        sb.append(">");

        int rId = 1;
        int shapeId = 2;

        // Collect all rIds already spoken for by shapes (extra rels + embedded images)
        Set<String> usedRIds = new HashSet<>();
        for (ShapeModel shape : ws.getShapes()) {
            for (String[] rel : shape.getExtraDrawingRels()) usedRIds.add(rel[0]);
            usedRIds.addAll(shape.getEmbeddedImageData().keySet());
        }
        // Also reserve original rIds of pictures and charts
        for (PictureModel pic : ws.getPictures())
            if (pic.getOriginalRelId() != null) usedRIds.add(pic.getOriginalRelId());
        for (ChartModel chart : ws.getCharts())
            if (chart.getOriginalRelId() != null) usedRIds.add(chart.getOriginalRelId());

        // Pictures — use original rId when available; allocate a fresh one otherwise
        for (int i = 0; i < ws.getPictures().size(); i++) {
            PictureModel pic = ws.getPictures().get(i);
            String imgRId;
            if (pic.getOriginalRelId() != null && !pic.getOriginalRelId().isEmpty()) {
                imgRId = pic.getOriginalRelId();
            } else {
                while (usedRIds.contains("rId" + rId)) rId++;
                imgRId = "rId" + rId++;
                usedRIds.add(imgRId);
            }
            int imgIdx = imageOffset + i + 1;
            String ext = pic.getExtension();
            String mediaPath = "xl/media/image" + imgIdx + "." + ext;
            String relTarget = "../media/image" + imgIdx + "." + ext;
            mediaFiles.put(mediaPath, pic.getData());
            drawingRels.add(new String[]{imgRId, IMAGE_REL_TYPE, relTarget});

            sb.append(twoCellAnchor(pic.getUpperLeftRow(), pic.getUpperLeftColumn(),
                pic.getLowerRightRow(), pic.getLowerRightColumn()));
            sb.append(picXml(shapeId++, pic.getName(), imgRId));
            sb.append("</xdr:twoCellAnchor>");
        }

        // Shapes — emit embedded image rels + extra rels (hyperlinks), then draw the shape
        Set<String> emittedExtraRIds = new HashSet<>();
        // Shape-embedded images start right after standalone pictures
        int shapeImgIdx = imageOffset + ws.getPictures().size();
        for (int i = 0; i < ws.getShapes().size(); i++) {
            ShapeModel shape = ws.getShapes().get(i);
            // Embedded images (blip fills inside grpSp etc.)
            for (Map.Entry<String, byte[]> imgE : shape.getEmbeddedImageData().entrySet()) {
                String origImgRId = imgE.getKey();
                if (emittedExtraRIds.add(origImgRId)) {
                    shapeImgIdx++;
                    String ext = PictureCollection.extensionFromData(imgE.getValue());
                    mediaFiles.put("xl/media/image" + shapeImgIdx + "." + ext, imgE.getValue());
                    drawingRels.add(new String[]{origImgRId, IMAGE_REL_TYPE,
                        "../media/image" + shapeImgIdx + "." + ext});
                }
            }
            // Extra rels (hyperlinks etc.)
            for (String[] rel : shape.getExtraDrawingRels()) {
                if (emittedExtraRIds.add(rel[0]))
                    drawingRels.add(rel);
            }
            sb.append(twoCellAnchor(shape.getUpperLeftRow(), shape.getUpperLeftColumn(),
                shape.getLowerRightRow(), shape.getLowerRightColumn()));
            if (shape.getRawElementXml() != null && !shape.getRawElementXml().isBlank()) {
                sb.append(shape.getRawElementXml());
            } else {
                sb.append(shapeXml(shapeId, shape));
            }
            sb.append("<xdr:clientData/></xdr:twoCellAnchor>");
            shapeId++;
        }

        // Charts — images start after standalone pictures AND shape-embedded images
        int totalShapeImgs = ws.getShapes().stream()
            .mapToInt(s -> s.getEmbeddedImageData().size()).sum();
        int chartImageIdx = imageOffset + ws.getPictures().size() + totalShapeImgs;
        for (int i = 0; i < ws.getCharts().size(); i++) {
            ChartModel chart = ws.getCharts().get(i);
            String chartRId;
            if (chart.getOriginalRelId() != null && !chart.getOriginalRelId().isEmpty()) {
                chartRId = chart.getOriginalRelId();
            } else {
                while (usedRIds.contains("rId" + rId)) rId++;
                chartRId = "rId" + rId++;
                usedRIds.add(chartRId);
            }
            int chartIdx = chartOffset + i + 1;
            boolean isChartEx = chart.isChartEx() && chart.getRawGraphicFrameXml() != null;
            String chartFileName = isChartEx ? "chartEx" + chartIdx : "chart" + chartIdx;
            String chartRelType = isChartEx ? CHARTEX_REL_TYPE : CHART_REL_TYPE;

            // Write chart XML
            chartXmls.put("xl/charts/" + chartFileName + ".xml",
                chart.getRawChartXml() != null
                    ? chart.getRawChartXml().getBytes(StandardCharsets.UTF_8)
                    : new byte[0]);
            drawingRels.add(new String[]{chartRId, chartRelType, "../charts/" + chartFileName + ".xml"});

            // Write chart rels file (companion style/colors + embedded images)
            if (!chart.getChartRels().isEmpty()) {
                StringBuilder chartRelsXml = new StringBuilder(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (String[] rel : chart.getChartRels()) {
                    String relId    = rel[0];
                    String relType  = rel[1];
                    byte[] content  = chart.getChartRelContent().get(relId);
                    String newTarget;
                    if (relType.contains("chartStyle")) {
                        newTarget = "style" + chartIdx + ".xml";
                        if (content != null) chartXmls.put("xl/charts/" + newTarget, content);
                    } else if (relType.contains("chartColorStyle")) {
                        newTarget = "colors" + chartIdx + ".xml";
                        if (content != null) chartXmls.put("xl/charts/" + newTarget, content);
                    } else if (relType.contains("/relationships/image") && content != null) {
                        chartImageIdx++;
                        String ext = PictureCollection.extensionFromData(content);
                        newTarget = "../media/image" + chartImageIdx + "." + ext;
                        mediaFiles.put("xl/media/image" + chartImageIdx + "." + ext, content);
                    } else {
                        newTarget = rel[2]; // preserve original target for unknown rel types
                    }
                    chartRelsXml.append("<Relationship Id=\"").append(relId)
                        .append("\" Type=\"").append(relType)
                        .append("\" Target=\"").append(newTarget).append("\"/>");
                }
                chartRelsXml.append("</Relationships>");
                chartXmls.put("xl/charts/_rels/" + chartFileName + ".xml.rels",
                    chartRelsXml.toString().getBytes(StandardCharsets.UTF_8));
            }

            // Write drawing element
            sb.append(twoCellAnchor(chart.getUpperLeftRow(), chart.getUpperLeftColumn(),
                chart.getLowerRightRow(), chart.getLowerRightColumn()));
            if (isChartEx) {
                // Preserve AlternateContent wrapper with rId replacement
                String frameXml = chart.getRawGraphicFrameXml();
                String origRId  = chart.getOriginalRelId();
                if (origRId != null && !origRId.isEmpty()) {
                    frameXml = frameXml.replace("r:id=\"" + origRId + "\"", "r:id=\"" + chartRId + "\"");
                    frameXml = frameXml.replace("r:id='" + origRId + "'",   "r:id='" + chartRId + "'");
                }
                sb.append(frameXml).append("<xdr:clientData/>");
            } else {
                sb.append(graphicFrameXml(shapeId++, chart.getName(), chartRId));
            }
            sb.append("</xdr:twoCellAnchor>");
        }

        sb.append("</xdr:wsDr>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    // =========================================================================
    // Load drawing XML
    // =========================================================================

    static void loadDrawings(WorksheetModel ws, Map<String, byte[]> entries,
                              Map<String, String> sheetRels) {
        String drawingRelId = null;
        String drawingTarget = null;
        for (Map.Entry<String, String> e : sheetRels.entrySet()) {
            String target = e.getValue();
            if (target != null && target.toLowerCase().contains("drawing")) {
                drawingRelId = e.getKey();
                drawingTarget = target;
                break;
            }
        }
        if (drawingTarget == null) return;

        String drawingPath = resolveRelTarget("xl/worksheets/sheet1.xml", drawingTarget);
        byte[] drawBytes = entries.get(drawingPath);
        if (drawBytes == null) return;

        Document drawDoc = XlsxWorkbookArchiveHelpers.parse(drawBytes);
        if (drawDoc == null) return;

        // Load drawing rels
        String relsPath = drawingPath.substring(0, drawingPath.lastIndexOf('/') + 1)
                + "_rels/" + drawingPath.substring(drawingPath.lastIndexOf('/') + 1) + ".rels";
        Map<String, String> drawingRels = XlsxWorkbookArchiveHelpers.loadRels(entries, relsPath);
        Map<String, String[]> drawingRelsTyped = loadRelsWithTypes(entries, relsPath);

        // Parse pictures
        loadPictures(ws, drawDoc, drawingRels, entries, drawingPath);

        // Parse shapes (sp, cxnSp, grpSp) — pass full rels so extra refs can be stored
        loadShapes(ws, drawDoc, drawingRelsTyped, entries, drawingPath);

        // Parse charts
        loadCharts(ws, drawDoc, drawingRels, entries, drawingPath);
    }

    private static final java.util.regex.Pattern RID_ATTR_PATTERN =
        java.util.regex.Pattern.compile("r:[a-zA-Z]+=[\"|'](rId\\d+)[\"|']");

    private static void loadShapes(WorksheetModel ws, Document drawDoc,
                                   Map<String, String[]> drawingRelsTyped,
                                   Map<String, byte[]> entries, String drawingPath) {
        // Collect all anchor elements (both two-cell and one-cell)
        List<Element> anchors = new ArrayList<>();
        NodeList all = drawDoc.getElementsByTagNameNS("*", "twoCellAnchor");
        for (int i = 0; i < all.getLength(); i++) anchors.add((Element) all.item(i));
        NodeList oneCell = drawDoc.getElementsByTagNameNS("*", "oneCellAnchor");
        for (int i = 0; i < oneCell.getLength(); i++) anchors.add((Element) oneCell.item(i));

        for (Element anchor : anchors) {
            // Skip chart anchors (handled by loadCharts)
            if (anchor.getElementsByTagNameNS("*", "graphicFrame").getLength() > 0) continue;
            if (anchor.getElementsByTagNameNS("*", "AlternateContent").getLength() > 0) continue;
            // Note: anchors with embedded pics (inside grpSp) are intentionally NOT skipped here;
            // loadPictures() only extracts standalone pic anchors (not nested ones).

            // sp (regular shape), cxnSp (connector), grpSp (group shape)
            Element sp     = firstChildByLocalName(anchor, "sp");
            Element cxnSp  = firstChildByLocalName(anchor, "cxnSp");
            Element grpSp  = firstChildByLocalName(anchor, "grpSp");
            Element shapeEl = sp != null ? sp : (cxnSp != null ? cxnSp : grpSp);
            if (shapeEl == null) continue;

            ShapeModel model = new ShapeModel();
            model.setRawElementXml(elementToString(shapeEl));

            // Name from nvSpPr/cNvPr, nvCxnSpPr/cNvPr, or nvGrpSpPr/cNvPr
            String[] nvPrParents = {"nvSpPr", "nvCxnSpPr", "nvGrpSpPr"};
            for (String nvPr : nvPrParents) {
                NodeList nvNodes = shapeEl.getElementsByTagNameNS("*", nvPr);
                if (nvNodes.getLength() > 0) {
                    NodeList cNvPr = ((Element) nvNodes.item(0)).getElementsByTagNameNS("*", "cNvPr");
                    if (cNvPr.getLength() > 0) {
                        model.setName(((Element) cNvPr.item(0)).getAttribute("name"));
                        break;
                    }
                }
            }

            // Geometry type from spPr/prstGeom or cxnSpPr/prstGeom
            NodeList spPrNodes = shapeEl.getElementsByTagNameNS("*", "spPr");
            if (spPrNodes.getLength() == 0) spPrNodes = shapeEl.getElementsByTagNameNS("*", "cxnSpPr");
            if (spPrNodes.getLength() > 0) {
                NodeList prstGeom = ((Element) spPrNodes.item(0)).getElementsByTagNameNS("*", "prstGeom");
                if (prstGeom.getLength() > 0)
                    model.setGeometryType(((Element) prstGeom.item(0)).getAttribute("prst"));
            }

            // Preserve style and txBody for sp elements
            if (sp != null) {
                NodeList styleNodes = sp.getElementsByTagNameNS("*", "style");
                if (styleNodes.getLength() > 0) model.setRawStyleXml(elementToString((Element) styleNodes.item(0)));
                NodeList txBodyNodes = sp.getElementsByTagNameNS("*", "txBody");
                if (txBodyNodes.getLength() > 0) model.setRawTxBodyXml(elementToString((Element) txBodyNodes.item(0)));
            }

            // Collect embedded images (e.g. blip fills inside grpSp) and extra rels (hyperlinks)
            if (!drawingRelsTyped.isEmpty()) {
                // Embedded images: from <a:blip r:embed="rIdX"/> elements
                NodeList blips = shapeEl.getElementsByTagNameNS("*", "blip");
                for (int b = 0; b < blips.getLength(); b++) {
                    Element blip = (Element) blips.item(b);
                    String embedId = blip.getAttribute("r:embed");
                    if (embedId.isEmpty()) embedId = blip.getAttributeNS(R_NS, "embed");
                    if (!embedId.isEmpty() && !model.getEmbeddedImageData().containsKey(embedId)) {
                        String[] relEntry = drawingRelsTyped.get(embedId);
                        if (relEntry != null && relEntry[0].contains("/relationships/image")) {
                            String imgPath = resolveRelTarget(drawingPath, relEntry[1]);
                            byte[] imgBytes = entries.get(imgPath);
                            if (imgBytes != null) model.getEmbeddedImageData().put(embedId, imgBytes);
                        }
                    }
                }
                // Extra rels: non-image rId refs in rawElementXml (e.g. hyperlinks)
                if (model.getRawElementXml() != null) {
                    java.util.regex.Matcher m = RID_ATTR_PATTERN.matcher(model.getRawElementXml());
                    Set<String> seen = new HashSet<>();
                    while (m.find()) {
                        String refId = m.group(1);
                        if (seen.add(refId) && !model.getEmbeddedImageData().containsKey(refId)) {
                            String[] relEntry = drawingRelsTyped.get(refId);
                            if (relEntry != null && !relEntry[0].contains("/relationships/image")) {
                                // 4th element is TargetMode (e.g. "External" for hyperlinks)
                                model.getExtraDrawingRels().add(new String[]{refId, relEntry[0], relEntry[1], relEntry[2]});
                            }
                        }
                    }
                }
            }

            // Anchor coordinates
            int[] coords = extractAnchorFromElement(anchor);
            model.setUpperLeftRow(coords[0]);
            model.setUpperLeftColumn(coords[1]);
            model.setLowerRightRow(coords[2]);
            model.setLowerRightColumn(coords[3]);

            ws.getShapes().add(model);
        }
    }

    private static Element firstChildByLocalName(Element parent, String localName) {
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node n = children.item(i);
            if (n.getNodeType() == Node.ELEMENT_NODE && localName.equals(n.getLocalName()))
                return (Element) n;
        }
        return null;
    }

    private static String elementToString(Element el) {
        try {
            javax.xml.transform.TransformerFactory tf = javax.xml.transform.TransformerFactory.newInstance();
            javax.xml.transform.Transformer t = tf.newTransformer();
            t.setOutputProperty(javax.xml.transform.OutputKeys.OMIT_XML_DECLARATION, "yes");
            java.io.StringWriter sw = new java.io.StringWriter();
            t.transform(new javax.xml.transform.dom.DOMSource(el), new javax.xml.transform.stream.StreamResult(sw));
            return sw.toString();
        } catch (Exception e) {
            return "";
        }
    }

    private static int[] extractAnchorFromElement(Element anchor) {
        int ulr = 0, ulc = 0, lrr = 0, lrc = 0;
        NodeList froms = anchor.getElementsByTagNameNS("*", "from");
        NodeList tos   = anchor.getElementsByTagNameNS("*", "to");
        if (froms.getLength() > 0) { Element f = (Element) froms.item(0); ulr = parseInt(f, "row"); ulc = parseInt(f, "col"); }
        if (tos.getLength() > 0)   { Element t = (Element) tos.item(0);   lrr = parseInt(t, "row"); lrc = parseInt(t, "col"); }
        return new int[]{ulr, ulc, lrr, lrc};
    }

    private static final java.util.Set<String> CONNECTOR_GEOMS = new java.util.HashSet<>(java.util.Arrays.asList(
        "straightConnector1",
        "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5",
        "curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5"
    ));

    private static String shapeXml(int shapeId, ShapeModel shape) {
        long cx = shape.getExtentCx() > 0 ? shape.getExtentCx()
                : Math.max(609600L, (long)(shape.getLowerRightColumn() - shape.getUpperLeftColumn()) * 609600L);
        long cy = shape.getExtentCy() > 0 ? shape.getExtentCy()
                : Math.max(190500L, (long)(shape.getLowerRightRow() - shape.getUpperLeftRow()) * 190500L);
        String geom = shape.getGeometryType() != null && !shape.getGeometryType().isEmpty()
                ? shape.getGeometryType() : "rect";
        String name = shape.getName() != null && !shape.getName().isEmpty()
                ? shape.getName() : "Shape " + shapeId;

        if (CONNECTOR_GEOMS.contains(geom)) {
            return connectorXml(shapeId, name, geom, cx, cy);
        }

        StringBuilder sb = new StringBuilder();
        sb.append("<xdr:sp macro=\"\" textlink=\"\">");
        sb.append("<xdr:nvSpPr>");
        sb.append("<xdr:cNvPr id=\"").append(shapeId).append("\" name=\"").append(xmlAttr(name)).append("\"/>");
        sb.append("<xdr:cNvSpPr/>");
        sb.append("</xdr:nvSpPr>");
        sb.append("<xdr:spPr>");
        sb.append("<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"").append(cx).append("\" cy=\"").append(cy).append("\"/></a:xfrm>");
        sb.append("<a:prstGeom prst=\"").append(xmlAttr(geom)).append("\"><a:avLst/></a:prstGeom>");
        sb.append("</xdr:spPr>");
        sb.append("<xdr:style>");
        sb.append("<a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef>");
        sb.append("<a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef>");
        sb.append("<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>");
        sb.append("<a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef>");
        sb.append("</xdr:style>");
        sb.append("<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></xdr:txBody>");
        sb.append("</xdr:sp>");
        return sb.toString();
    }

    private static String connectorXml(int shapeId, String name, String geom, long cx, long cy) {
        StringBuilder sb = new StringBuilder();
        sb.append("<xdr:cxnSp macro=\"\">");
        sb.append("<xdr:nvCxnSpPr>");
        sb.append("<xdr:cNvPr id=\"").append(shapeId).append("\" name=\"").append(xmlAttr(name)).append("\"/>");
        sb.append("<xdr:cNvCxnSpPr/>");
        sb.append("</xdr:nvCxnSpPr>");
        sb.append("<xdr:spPr>");
        sb.append("<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"").append(cx).append("\" cy=\"").append(cy).append("\"/></a:xfrm>");
        sb.append("<a:prstGeom prst=\"").append(xmlAttr(geom)).append("\"><a:avLst/></a:prstGeom>");
        sb.append("</xdr:spPr>");
        sb.append("<xdr:style>");
        sb.append("<a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:lnRef>");
        sb.append("<a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef>");
        sb.append("<a:effectRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:effectRef>");
        sb.append("<a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef>");
        sb.append("</xdr:style>");
        sb.append("</xdr:cxnSp>");
        return sb.toString();
    }

    private static void loadPictures(WorksheetModel ws, Document drawDoc,
                                      Map<String, String> drawingRels,
                                      Map<String, byte[]> entries, String drawingPath) {
        NodeList picNodes = drawDoc.getElementsByTagNameNS("*", "pic");
        for (int i = 0; i < picNodes.getLength(); i++) {
            Element picEl = (Element) picNodes.item(i);
            // Skip pics nested inside grpSp — those are preserved as part of the group shape's rawElementXml
            if (isInsideGrpSp(picEl)) continue;
            PictureModel pic = new PictureModel();

            // Get name from cNvPr
            NodeList cNvPr = picEl.getElementsByTagNameNS("*", "cNvPr");
            if (cNvPr.getLength() > 0) {
                pic.setName(((Element)cNvPr.item(0)).getAttribute("name"));
            }

            // Get embed rId
            String embedRId = null;
            NodeList blips = picEl.getElementsByTagNameNS("*", "blip");
            if (blips.getLength() > 0) {
                Element blip = (Element)blips.item(0);
                embedRId = blip.getAttribute("r:embed");
                if (embedRId.isEmpty()) embedRId = blip.getAttributeNS(R_NS, "embed");
            }

            // Load image data
            if (embedRId != null && !embedRId.isEmpty()) {
                String imgTarget = drawingRels.get(embedRId);
                if (imgTarget != null) {
                    String imgPath = resolveRelTarget(drawingPath, imgTarget);
                    byte[] imgData = entries.get(imgPath);
                    if (imgData != null) {
                        pic.setData(imgData);
                        pic.setExtension(PictureCollection.extensionFromData(imgData));
                    }
                }
                pic.setOriginalRelId(embedRId);
            }

            // Anchor — find the twoCellAnchor or oneCellAnchor parent
            int[] anchors = extractAnchor(picEl);
            pic.setUpperLeftRow(anchors[0]);
            pic.setUpperLeftColumn(anchors[1]);
            pic.setLowerRightRow(anchors[2]);
            pic.setLowerRightColumn(anchors[3]);

            if (pic.getData() != null) ws.getPictures().add(pic);
        }
    }

    private static void loadCharts(WorksheetModel ws, Document drawDoc,
                                    Map<String, String> drawingRels,
                                    Map<String, byte[]> entries, String drawingPath) {
        NodeList frameNodes = drawDoc.getElementsByTagNameNS("*", "graphicFrame");
        for (int i = 0; i < frameNodes.getLength(); i++) {
            Element frame = (Element)frameNodes.item(i);
            ChartModel chart = new ChartModel();

            // Detect ChartEx: graphicFrame inside mc:AlternateContent
            Element alternateContent = findAlternateContentAncestor(frame);

            // Name
            NodeList cNvPr = frame.getElementsByTagNameNS("*", "cNvPr");
            if (cNvPr.getLength() > 0) {
                chart.setName(((Element)cNvPr.item(0)).getAttribute("name"));
            }

            // Chart rId — works for both c:chart and cx:chart since we match any namespace
            String chartRId = null;
            NodeList chartNodes = frame.getElementsByTagNameNS("*", "chart");
            for (int j = 0; j < chartNodes.getLength(); j++) {
                Element cEl = (Element)chartNodes.item(j);
                chartRId = cEl.getAttribute("r:id");
                if (chartRId.isEmpty()) chartRId = cEl.getAttributeNS(R_NS, "id");
                if (!chartRId.isEmpty()) break;
            }

            if (chartRId == null || chartRId.isEmpty()) continue;

            // Load chart XML
            String chartTarget = drawingRels.get(chartRId);
            if (chartTarget != null) {
                String chartPath = resolveRelTarget(drawingPath, chartTarget);
                byte[] chartBytes = entries.get(chartPath);
                if (chartBytes != null) {
                    String rawXml = new String(chartBytes, StandardCharsets.UTF_8);
                    chart.setRawChartXml(rawXml);
                    chart.setChartType(XlsxWorkbookChartTemplates.detectChartType(rawXml));
                }

                if (alternateContent != null) {
                    chart.setChartEx(true);
                    chart.setRawGraphicFrameXml(elementToString(alternateContent));
                }
                // Load chart rels for ALL charts (standard and ChartEx)
                String chartRelsPath = chartPath.substring(0, chartPath.lastIndexOf('/') + 1)
                    + "_rels/" + chartPath.substring(chartPath.lastIndexOf('/') + 1) + ".rels";
                Map<String, String[]> chartRelsMap = loadRelsWithTypes(entries, chartRelsPath);
                for (Map.Entry<String, String[]> re : chartRelsMap.entrySet()) {
                    String relId      = re.getKey();
                    String relType    = re.getValue()[0];
                    String relTarget2 = re.getValue()[1];
                    chart.getChartRels().add(new String[]{relId, relType, relTarget2});
                    String contentPath = resolveRelTarget(chartPath, relTarget2);
                    byte[] contentBytes = entries.get(contentPath);
                    if (contentBytes != null) chart.getChartRelContent().put(relId, contentBytes);
                }
            }
            chart.setOriginalRelId(chartRId);

            // Anchor — extractAnchor walks up from frame through mc:AlternateContent to the anchor
            int[] anchors = extractAnchor(frame);
            chart.setUpperLeftRow(anchors[0]);
            chart.setUpperLeftColumn(anchors[1]);
            chart.setLowerRightRow(anchors[2]);
            chart.setLowerRightColumn(anchors[3]);

            ws.getCharts().add(chart);
        }
    }

    // =========================================================================
    // Private helpers
    // =========================================================================

    private static String twoCellAnchor(int r1, int c1, int r2, int c2) {
        return "<xdr:twoCellAnchor editAs=\"oneCell\">"
             + "<xdr:from><xdr:col>" + c1 + "</xdr:col><xdr:colOff>0</xdr:colOff>"
             + "<xdr:row>" + r1 + "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
             + "<xdr:to><xdr:col>" + c2 + "</xdr:col><xdr:colOff>0</xdr:colOff>"
             + "<xdr:row>" + r2 + "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>";
    }

    private static String picXml(int shapeId, String name, String rId) {
        return "<xdr:pic>"
             + "<xdr:nvPicPr>"
             + "<xdr:cNvPr id=\"" + shapeId + "\" name=\"" + xmlAttr(name) + "\"/>"
             + "<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr>"
             + "</xdr:nvPicPr>"
             + "<xdr:blipFill>"
             + "<a:blip xmlns:r=\"" + R_NS + "\" r:embed=\"" + rId + "\"/>"
             + "<a:stretch><a:fillRect/></a:stretch>"
             + "</xdr:blipFill>"
             + "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1\" cy=\"1\"/></a:xfrm>"
             + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>"
             + "</xdr:pic><xdr:clientData/>";
    }

    private static String graphicFrameXml(int shapeId, String name, String rId) {
        return "<xdr:graphicFrame macro=\"\">"
             + "<xdr:nvGraphicFramePr>"
             + "<xdr:cNvPr id=\"" + shapeId + "\" name=\"" + xmlAttr(name) + "\"/>"
             + "<xdr:cNvGraphicFramePr/>"
             + "</xdr:nvGraphicFramePr>"
             + "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>"
             + "<a:graphic>"
             + "<a:graphicData uri=\"" + C + "\">"
             + "<c:chart xmlns:c=\"" + C + "\" xmlns:r=\"" + R_NS + "\" r:id=\"" + rId + "\"/>"
             + "</a:graphicData></a:graphic>"
             + "</xdr:graphicFrame><xdr:clientData/>";
    }

    private static int[] extractAnchor(Element child) {
        // Walk up to the anchor element
        Node p = child.getParentNode();
        while (p != null && p.getNodeType() == Node.ELEMENT_NODE) {
            String ln = ((Element)p).getLocalName();
            if ("twoCellAnchor".equals(ln) || "oneCellAnchor".equals(ln)) break;
            p = p.getParentNode();
        }
        if (p == null || p.getNodeType() != Node.ELEMENT_NODE) return new int[4];

        Element anchor = (Element)p;
        NodeList froms = anchor.getElementsByTagNameNS("*", "from");
        NodeList tos   = anchor.getElementsByTagNameNS("*", "to");

        int ulr = 0, ulc = 0, lrr = 0, lrc = 0;
        if (froms.getLength() > 0) {
            Element from = (Element)froms.item(0);
            ulr = parseInt(from, "row");
            ulc = parseInt(from, "col");
        }
        if (tos.getLength() > 0) {
            Element to = (Element)tos.item(0);
            lrr = parseInt(to, "row");
            lrc = parseInt(to, "col");
        }
        return new int[]{ulr, ulc, lrr, lrc};
    }

    private static int parseInt(Element parent, String localName) {
        NodeList nl = parent.getElementsByTagNameNS("*", localName);
        if (nl.getLength() == 0) return 0;
        try { return Integer.parseInt(nl.item(0).getTextContent().trim()); }
        catch (NumberFormatException e) { return 0; }
    }

    static String resolveRelTarget(String baseEntry, String target) {
        if (target == null) return "";
        if (target.startsWith("/")) return target.substring(1);
        String base = baseEntry.contains("/") ? baseEntry.substring(0, baseEntry.lastIndexOf('/') + 1) : "";
        String[] parts = (base + target).split("/");
        java.util.Deque<String> stack = new java.util.ArrayDeque<>();
        for (String part : parts) {
            if ("..".equals(part)) { if (!stack.isEmpty()) stack.pollLast(); }
            else if (!part.isEmpty() && !".".equals(part)) stack.addLast(part);
        }
        return String.join("/", stack);
    }

    /** Returns true if el is nested inside a grpSp (before reaching the anchor). */
    private static boolean isInsideGrpSp(Element el) {
        Node n = el.getParentNode();
        while (n != null && n.getNodeType() == Node.ELEMENT_NODE) {
            String ln = ((Element) n).getLocalName();
            if ("grpSp".equals(ln)) return true;
            if ("twoCellAnchor".equals(ln) || "oneCellAnchor".equals(ln)) break;
            n = n.getParentNode();
        }
        return false;
    }

    /** Returns the nearest mc:AlternateContent ancestor within the same anchor, or null. */
    private static Element findAlternateContentAncestor(Element el) {
        Node n = el.getParentNode();
        while (n != null && n.getNodeType() == Node.ELEMENT_NODE) {
            String ln = ((Element) n).getLocalName();
            if ("AlternateContent".equals(ln)) return (Element) n;
            if ("twoCellAnchor".equals(ln) || "oneCellAnchor".equals(ln)) break;
            n = n.getParentNode();
        }
        return null;
    }

    /** Loads a rels file returning id → [type, target]. */
    private static Map<String, String[]> loadRelsWithTypes(Map<String, byte[]> entries, String path) {
        byte[] bytes = entries.get(path);
        if (bytes == null) return new LinkedHashMap<>();
        try {
            Document doc = XlsxWorkbookArchiveHelpers.parse(bytes);
            Map<String, String[]> result = new LinkedHashMap<>();
            NodeList nodes = doc.getElementsByTagNameNS("*", "Relationship");
            for (int i = 0; i < nodes.getLength(); i++) {
                Element el = (Element) nodes.item(i);
                String tm = el.getAttribute("TargetMode");
                result.put(el.getAttribute("Id"),
                    new String[]{el.getAttribute("Type"), el.getAttribute("Target"),
                                 tm.isEmpty() ? null : tm});
            }
            return result;
        } catch (Exception e) {
            return new LinkedHashMap<>();
        }
    }

    private static String xmlAttr(String s) {
        if (s == null) return "";
        return s.replace("&","&amp;").replace("<","&lt;").replace("\"","&quot;");
    }
}
