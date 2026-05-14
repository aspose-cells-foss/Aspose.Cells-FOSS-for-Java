package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.zip.*;

/**
 * Serializer for XLSX workbook files — thin coordinator that delegates to helper classes.
 */
public final class XlsxWorkbookSerializer {

    /**
     * Initializes a new XlsxWorkbookSerializer instance.
     */
    private XlsxWorkbookSerializer() {}

    // =========================================================================
    // SAVE
    // =========================================================================

    /**
     * Saves the current content.
     * @param model model
     * @param stream stream to apply
     * @param options options to apply
     */
    public static void save(WorkbookModel model, OutputStream stream, SaveOptions options) throws IOException {
        if (options.getSaveFormat() != SaveFormat.XLSX) {
            throw new UnsupportedFeatureException("Only XLSX save is supported.");
        }
        List<WorksheetModel> sheets = model.getWorksheets();
        if (sheets.isEmpty()) throw new WorkbookSaveException("A workbook must contain at least one worksheet.");

        SharedStringRepository sst = new SharedStringRepository();
        if (options.getUseSharedStrings()) {
            for (WorksheetModel ws : sheets) {
                for (Map.Entry<CellAddress, CellRecord> e : ws.getCells().entrySet()) {
                    if (XlsxWorkbookSerializerCommon.shouldPersist(model.getDefaultStyle(), e.getValue())
                            && e.getValue().getKind() == CellValueKind.STRING
                            && e.getValue().getValue() instanceof String s) {
                        sst.intern(s);
                    }
                }
            }
        }
        boolean hasSst = !sst.getValues().isEmpty();
        boolean hasDocProps = true; // always include docProps per OPC "should" requirement

        // Pre-build style table so styles.xml and cell s= indices are consistent
        XlsxWorkbookStyles.StyleTable styleTable = XlsxWorkbookStyles.buildStyleTable(model);

        int[] tableCounter = {1};
        int[] chartCounter = {1};
        int[] imageCounter = {1};
        int[] drawingCounter = {1};

        // Collect all drawing parts first (charts, pictures) so Content_Types can be built
        Map<String, byte[]> pendingChartXmls = new LinkedHashMap<>();
        Map<String, byte[]> pendingMediaFiles = new LinkedHashMap<>();

        // Pre-pass: intentional reset — counters are accumulated per-sheet in the build loop below
        chartCounter[0] = 1;
        imageCounter[0] = 1;

        // Build drawings data before writing the ZIP
        Map<Integer, byte[]> drawingXmls = new LinkedHashMap<>();
        Map<Integer, List<String[]>> drawingRelsMaps = new LinkedHashMap<>();
        Map<Integer, Integer> sheetToDrawingIndex = new LinkedHashMap<>();

        for (int i = 0; i < sheets.size(); i++) {
            WorksheetModel ws = sheets.get(i);
            if (!XlsxWorkbookDrawings.hasDrawings(ws)) continue;
            int dIdx = drawingCounter[0]++;
            List<String[]> dRels = new ArrayList<>();
            byte[] dXml = XlsxWorkbookDrawings.buildDrawingXml(
                ws, chartCounter[0] - 1, imageCounter[0] - 1, dRels, pendingChartXmls, pendingMediaFiles);
            chartCounter[0] += ws.getCharts().size();
            imageCounter[0] += ws.getPictures().size();
            for (ChartModel chart : ws.getCharts()) imageCounter[0] += chart.getChartImageCount();
            for (ShapeModel s : ws.getShapes()) imageCounter[0] += s.getEmbeddedImageData().size();
            drawingXmls.put(dIdx, dXml);
            drawingRelsMaps.put(dIdx, dRels);
            sheetToDrawingIndex.put(i, dIdx);
        }

        // Reset counters for actual writing
        chartCounter[0] = 1;
        imageCounter[0] = 1;
        drawingCounter[0] = 1;

        boolean hasDrawings = !drawingXmls.isEmpty();
        boolean hasCharts = !pendingChartXmls.isEmpty();
        boolean hasVml = sheets.stream().anyMatch(ws -> XlsxWorkbookComments.buildVmlDrawingXml(ws) != null);
        boolean hasTheme = hasDrawings;

        int totalTableCount = sheets.stream().mapToInt(ws -> ws.getListObjects().size()).sum();
        try (ZipOutputStream zip = new ZipOutputStream(stream, StandardCharsets.UTF_8)) {
            zip.setMethod(ZipOutputStream.DEFLATED);
            byte[] ctXml = XlsxWorkbookArchiveHelpers.contentTypesXml(sheets.size(), hasSst, hasDocProps,
                    hasDrawings, hasCharts, pendingChartXmls.keySet(), pendingMediaFiles,
                    drawingXmls.size(), hasVml, hasTheme);
            ctXml = XlsxWorkbookArchiveHelpers.addTableContentTypes(ctXml, totalTableCount);
            XlsxWorkbookArchiveHelpers.write(zip, "[Content_Types].xml", ctXml);
            XlsxWorkbookArchiveHelpers.write(zip, "_rels/.rels", XlsxWorkbookArchiveHelpers.packageRelsXml(hasDocProps));
            XlsxWorkbookArchiveHelpers.write(zip, "xl/workbook.xml", workbookXml(model));
            if (hasDocProps) {
                XlsxWorkbookArchiveHelpers.write(zip, "docProps/core.xml", buildCorePropertiesXml(model.getDocumentProperties().getCore()));
                XlsxWorkbookArchiveHelpers.write(zip, "docProps/app.xml", buildAppPropertiesXml(model.getDocumentProperties().getExtended()));
            }
            XlsxWorkbookArchiveHelpers.write(zip, "xl/_rels/workbook.xml.rels", XlsxWorkbookArchiveHelpers.workbookRelsXml(sheets.size(), hasSst, hasTheme));
            if (hasTheme) {
                byte[] themeBytes = model.getRawThemeXml() != null
                    ? model.getRawThemeXml()
                    : XlsxWorkbookArchiveHelpers.defaultThemeXml();
                XlsxWorkbookArchiveHelpers.write(zip, "xl/theme/theme1.xml", themeBytes);
            }

            for (int i = 0; i < sheets.size(); i++) {
                WorksheetModel ws = sheets.get(i);
                List<String[]> sheetRels = new ArrayList<>();

                Integer dIdxForSheet = sheetToDrawingIndex.get(i);
                String drawingRelId = dIdxForSheet != null ? "rIdDrw" + dIdxForSheet : null;
                XlsxWorkbookArchiveHelpers.write(zip, "xl/worksheets/sheet" + (i + 1) + ".xml",
                        worksheetXml(ws, sst, model, styleTable, sheetRels, tableCounter, drawingRelId));

                // Comments
                byte[] commentsXml = XlsxWorkbookComments.buildCommentsXml(ws);
                if (commentsXml != null) {
                    int sn = i + 1;
                    XlsxWorkbookArchiveHelpers.write(zip, "xl/comments" + sn + ".xml", commentsXml);
                    byte[] vmlXml = XlsxWorkbookComments.buildVmlDrawingXml(ws);
                    if (vmlXml != null)
                        XlsxWorkbookArchiveHelpers.write(zip, "xl/drawings/vmlDrawing" + sn + ".vml", vmlXml);
                    sheetRels.add(new String[]{"rIdComm" + sn,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                        "../comments" + sn + ".xml", null});
                    if (vmlXml != null)
                        sheetRels.add(new String[]{"rIdVml" + sn,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                            "../drawings/vmlDrawing" + sn + ".vml", null});
                }

                // Tables — indices were already assigned by buildTablePartsSnippet
                List<ListObjectModel> tables = ws.getListObjects();
                int firstTableIdx = tableCounter[0] - tables.size();
                for (int t = 0; t < tables.size(); t++) {
                    XlsxWorkbookArchiveHelpers.write(zip, "xl/tables/table" + (firstTableIdx + t) + ".xml",
                        XlsxWorkbookTables.buildTableXml(tables.get(t)));
                }

                // Drawing (pictures + charts)
                Integer dIdx = sheetToDrawingIndex.get(i);
                if (dIdx != null) {
                    String dPath = "xl/drawings/drawing" + dIdx + ".xml";
                    XlsxWorkbookArchiveHelpers.write(zip, dPath, drawingXmls.get(dIdx));
                    // Write drawing rels
                    List<String[]> dRels = drawingRelsMaps.get(dIdx);
                    if (!dRels.isEmpty()) {
                        XlsxWorkbookArchiveHelpers.write(zip,
                            "xl/drawings/_rels/drawing" + dIdx + ".xml.rels",
                            XlsxWorkbookHyperlinks.buildSheetRelsXml(dRels));
                    }
                    sheetRels.add(new String[]{"rIdDrw" + dIdx,
                        XlsxWorkbookDrawings.DRAWING_REL_TYPE,
                        "../drawings/drawing" + dIdx + ".xml", null});
                }

                if (!sheetRels.isEmpty()) {
                    XlsxWorkbookArchiveHelpers.write(zip, "xl/worksheets/_rels/sheet" + (i + 1) + ".xml.rels",
                            XlsxWorkbookHyperlinks.buildSheetRelsXml(sheetRels));
                }
            }

            // Media files (images)
            for (Map.Entry<String, byte[]> e : pendingMediaFiles.entrySet()) {
                XlsxWorkbookArchiveHelpers.write(zip, e.getKey(), e.getValue());
            }

            // Chart XML files
            for (Map.Entry<String, byte[]> e : pendingChartXmls.entrySet()) {
                XlsxWorkbookArchiveHelpers.write(zip, e.getKey(), e.getValue());
            }

            if (hasSst) XlsxWorkbookArchiveHelpers.write(zip, "xl/sharedStrings.xml", XlsxWorkbookArchiveHelpers.sharedStringsXml(sst));
            XlsxWorkbookArchiveHelpers.write(zip, "xl/styles.xml", styleTable.buildStylesXmlBytes());
        }
    }

    // --- XML builders ---------------------------------------------------------

    /**
     * Processes workbook xml.
     * @param model model
     * @return the computed result
     */
    private static byte[] workbookXml(WorkbookModel model) {
        List<WorksheetModel> sheets = model.getWorksheets();
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
        sb.append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        boolean date1904 = model.getSettings().getDateSystem() == DateSystem.MAC_1904;
        WorkbookPropertiesModel wp = model.getProperties();
        sb.append("<workbookPr date1904=\"").append(date1904 ? "1" : "0").append("\"");
        // Handle the relevant branch before the state changes.
        if (!wp.getCodeName().isBlank()) sb.append(" codeName=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(wp.getCodeName())).append("\"");
        if (!wp.getShowObjects().isBlank()) sb.append(" showObjects=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(wp.getShowObjects())).append("\"");
        if (wp.getFilterPrivacy()) sb.append(" filterPrivacy=\"1\"");
        if (!wp.getShowBorderUnselectedTables()) sb.append(" showBorderUnselectedTables=\"0\"");
        if (!wp.getShowInkAnnotation()) sb.append(" showInkAnnotation=\"0\"");
        if (wp.getBackupFile()) sb.append(" backupFile=\"1\"");
        if (!wp.getSaveExternalLinkValues()) sb.append(" saveExternalLinkValues=\"0\"");
        if (!wp.getUpdateLinks().isBlank()) sb.append(" updateLinks=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(wp.getUpdateLinks())).append("\"");
        if (wp.getHidePivotFieldList()) sb.append(" hidePivotFieldList=\"1\"");
        if (wp.getDefaultThemeVersion() != null) sb.append(" defaultThemeVersion=\"").append(wp.getDefaultThemeVersion()).append("\"");
        sb.append("/>");
        sb.append("<bookViews><workbookView activeTab=\"").append(model.getActiveSheetIndex()).append("\"/></bookViews>");
        sb.append("<sheets>");
        for (int i = 0; i < sheets.size(); i++) {
            WorksheetModel ws = sheets.get(i);
            String state = ws.getVisibility() == SheetVisibility.HIDDEN ? "hidden"
                    : ws.getVisibility() == SheetVisibility.VERY_HIDDEN ? "veryHidden" : "visible";
            sb.append("<sheet name=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(ws.getName()))
              .append("\" sheetId=\"").append(i + 1)
              .append("\" state=\"").append(state)
              .append("\" r:id=\"rId").append(i + 1).append("\"/>");
        }
        sb.append("</sheets>");
        XlsxWorkbookDefinedNames.buildDefinedNamesXml(model, sb);
        sb.append("</workbook>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    /**
     * Processes worksheet xml.
     * @param ws ws
     * @param sst sst
     * @param wb wb
     * @param styleTable style table
     * @param sheetRels accumulated sheet relationships (hyperlinks, tables, comments)
     * @param tableCounter global table index counter (incremented per table written)
     * @return the computed result
     */
    private static byte[] worksheetXml(WorksheetModel ws, SharedStringRepository sst, WorkbookModel wb,
                                       XlsxWorkbookStyles.StyleTable styleTable, List<String[]> sheetRels,
                                       int[] tableCounter, String drawingRelId) {
        List<String[]> externalRels = sheetRels;
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
        sb.append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

        XlsxWorkbookWorksheetViews.buildSheetPr(ws, sb);
        XlsxWorkbookWorksheetViews.buildSheetViews(ws, sb);

        // sheetFormatPr
        sb.append("<sheetFormatPr defaultRowHeight=\"15\"/>");

        XlsxWorkbookWorksheetViews.buildCols(ws, sb);

        // sheetData
        sb.append("<sheetData>");
        Map<Integer, RowModel> rowModels = ws.getRows();
        Map<CellAddress, CellRecord> cells = ws.getCells();
        TreeMap<Integer, List<Map.Entry<CellAddress, CellRecord>>> byRow = new TreeMap<>();
        for (Map.Entry<CellAddress, CellRecord> e : cells.entrySet()) {
            if (!XlsxWorkbookSerializerCommon.shouldPersist(wb.getDefaultStyle(), e.getValue())) continue;
            byRow.computeIfAbsent(e.getKey().getRowIndex(), k -> new ArrayList<>()).add(e);
        }
        // add rows with metadata but no cells
        for (Map.Entry<Integer, RowModel> rm : rowModels.entrySet()) {
            if (!byRow.containsKey(rm.getKey())) byRow.put(rm.getKey(), new ArrayList<>());
        }
        for (Map.Entry<Integer, List<Map.Entry<CellAddress, CellRecord>>> rowEntry : byRow.entrySet()) {
            int rowIdx = rowEntry.getKey();
            sb.append("<row r=\"").append(rowIdx + 1).append("\"");
            RowModel rm = rowModels.get(rowIdx);
            if (rm != null) {
                if (rm.getHeight() != null) sb.append(" ht=\"").append(XlsxWorkbookSerializerCommon.fmt(rm.getHeight())).append("\" customHeight=\"1\"");
                if (rm.getHidden()) sb.append(" hidden=\"1\"");
                if (rm.getOutlineLevel() > 0) sb.append(" outlineLevel=\"").append(rm.getOutlineLevel()).append("\"");
                if (rm.getCollapsed()) sb.append(" collapsed=\"1\"");
            }
            sb.append(">");
            List<Map.Entry<CellAddress, CellRecord>> rowCells = rowEntry.getValue();
            rowCells.sort(Comparator.comparingInt(e -> e.getKey().getColumnIndex()));
            for (Map.Entry<CellAddress, CellRecord> ce : rowCells) {
                CellAddress addr = ce.getKey();
                CellRecord rec = ce.getValue();
                String ref = XlsxWorkbookSerializerCommon.colLetter(addr.getColumnIndex()) + (addr.getRowIndex() + 1);
                boolean isDate = rec.getKind() == CellValueKind.DATE_TIME;
                int sIdx = styleTable.register(rec.getStyle(), isDate);
                String sAttr = sIdx > 0 ? " s=\"" + sIdx + "\"" : "";
                if (rec.getFormula() != null && !rec.getFormula().isEmpty()) {
                    sb.append("<c r=\"").append(ref).append("\"").append(sAttr).append("><f>")
                      .append(XlsxWorkbookSerializerCommon.xmlText(rec.getFormula())).append("</f>");
                    if (rec.getValue() instanceof Number n) {
                        sb.append("<v>").append(n).append("</v>");
                    }
                    sb.append("</c>");
                } else if (rec.getKind() == CellValueKind.STRING && rec.getValue() instanceof String s) {
                    int idx = sst.intern(s);
                    sb.append("<c r=\"").append(ref).append("\"").append(sAttr).append(" t=\"s\"><v>").append(idx).append("</v></c>");
                } else if (rec.getKind() == CellValueKind.BOOLEAN && rec.getValue() instanceof Boolean b) {
                    sb.append("<c r=\"").append(ref).append("\"").append(sAttr).append(" t=\"b\"><v>").append(b ? "1" : "0").append("</v></c>");
                } else if (isDate && rec.getValue() instanceof LocalDateTime dt) {
                    double serial = DateSerialConverter.toSerial(dt, wb.getSettings().getDateSystem());
                    sb.append("<c r=\"").append(ref).append("\"").append(sAttr).append("><v>")
                      .append(XlsxWorkbookSerializerCommon.fmt(serial)).append("</v></c>");
                } else if (rec.getKind() == CellValueKind.NUMBER && rec.getValue() instanceof Number n) {
                    sb.append("<c r=\"").append(ref).append("\"").append(sAttr).append("><v>").append(n).append("</v></c>");
                } else if (!sAttr.isEmpty()) {
                    // blank cell with non-default style — emit as styled empty cell
                    sb.append("<c r=\"").append(ref).append("\"").append(sAttr).append("/>");
                }
            }
            sb.append("</row>");
        }
        sb.append("</sheetData>");

        // mergeCells
        if (!ws.getMergeRegions().isEmpty()) {
            sb.append("<mergeCells count=\"").append(ws.getMergeRegions().size()).append("\">");
            for (MergeRegion mr : ws.getMergeRegions()) {
                int r1 = mr.getFirstRow(), c1 = mr.getFirstColumn();
                int r2 = r1 + mr.getTotalRows() - 1, c2 = c1 + mr.getTotalColumns() - 1;
                sb.append("<mergeCell ref=\"")
                  .append(XlsxWorkbookSerializerCommon.colLetter(c1)).append(r1 + 1).append(":")
                  .append(XlsxWorkbookSerializerCommon.colLetter(c2)).append(r2 + 1).append("\"/>");
            }
            sb.append("</mergeCells>");
        }

        XlsxWorkbookConditionalFormatting.buildConditionalFormattingSection(ws, sb, styleTable);
        XlsxWorkbookValidations.buildDataValidationsSection(ws, sb);
        XlsxWorkbookAutoFilter.buildAutoFilterSection(ws, sb);
        XlsxWorkbookWorksheetProtection.buildSheetProtectionSection(ws, sb);
        XlsxWorkbookHyperlinks.buildHyperlinksSection(ws, sb, externalRels);
        XlsxWorkbookPageSetup.buildPageSetupSection(ws, sb);
        if (drawingRelId != null)
            sb.append("<drawing r:id=\"").append(drawingRelId).append("\"/>");
        sb.append(XlsxWorkbookTables.buildTablePartsSnippet(ws, 0, sheetRels, tableCounter));

        sb.append("</worksheet>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    /**
     * Builds the core properties xml.
     * @param core core
     * @return the requested result
     */
    private static byte[] buildCorePropertiesXml(DocumentPropertiesModel.CoreDocumentPropertiesModel core) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<cp:coreProperties");
        sb.append(" xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"");
        sb.append(" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"");
        sb.append(" xmlns:dcterms=\"http://purl.org/dc/terms/\"");
        sb.append(" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        // Handle the relevant branch before the state changes.
        if (!core.getTitle().isEmpty()) sb.append("<dc:title>").append(XlsxWorkbookSerializerCommon.xmlText(core.getTitle())).append("</dc:title>");
        if (!core.getSubject().isEmpty()) sb.append("<dc:subject>").append(XlsxWorkbookSerializerCommon.xmlText(core.getSubject())).append("</dc:subject>");
        if (!core.getCreator().isEmpty()) sb.append("<dc:creator>").append(XlsxWorkbookSerializerCommon.xmlText(core.getCreator())).append("</dc:creator>");
        if (!core.getKeywords().isEmpty()) sb.append("<cp:keywords>").append(XlsxWorkbookSerializerCommon.xmlText(core.getKeywords())).append("</cp:keywords>");
        if (!core.getDescription().isEmpty()) sb.append("<dc:description>").append(XlsxWorkbookSerializerCommon.xmlText(core.getDescription())).append("</dc:description>");
        if (!core.getLastModifiedBy().isEmpty()) sb.append("<cp:lastModifiedBy>").append(XlsxWorkbookSerializerCommon.xmlText(core.getLastModifiedBy())).append("</cp:lastModifiedBy>");
        if (!core.getRevision().isEmpty()) sb.append("<cp:revision>").append(XlsxWorkbookSerializerCommon.xmlText(core.getRevision())).append("</cp:revision>");
        if (!core.getCategory().isEmpty()) sb.append("<cp:category>").append(XlsxWorkbookSerializerCommon.xmlText(core.getCategory())).append("</cp:category>");
        if (!core.getContentStatus().isEmpty()) sb.append("<cp:contentStatus>").append(XlsxWorkbookSerializerCommon.xmlText(core.getContentStatus())).append("</cp:contentStatus>");
        if (core.getCreated() != null) sb.append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">").append(formatIso(core.getCreated())).append("</dcterms:created>");
        if (core.getModified() != null) sb.append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">").append(formatIso(core.getModified())).append("</dcterms:modified>");
        sb.append("</cp:coreProperties>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    /**
     * Builds the app properties xml.
     * @param ext ext
     * @return the requested result
     */
    private static byte[] buildAppPropertiesXml(DocumentPropertiesModel.ExtendedDocumentPropertiesModel ext) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        sb.append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">");
        // Handle the relevant branch before the state changes.
        if (!ext.getApplication().isEmpty()) sb.append("<Application>").append(XlsxWorkbookSerializerCommon.xmlText(ext.getApplication())).append("</Application>");
        if (!ext.getAppVersion().isEmpty()) sb.append("<AppVersion>").append(XlsxWorkbookSerializerCommon.xmlText(ext.getAppVersion())).append("</AppVersion>");
        if (!ext.getCompany().isEmpty()) sb.append("<Company>").append(XlsxWorkbookSerializerCommon.xmlText(ext.getCompany())).append("</Company>");
        if (!ext.getManager().isEmpty()) sb.append("<Manager>").append(XlsxWorkbookSerializerCommon.xmlText(ext.getManager())).append("</Manager>");
        if (ext.getDocSecurity() != null) sb.append("<DocSecurity>").append(ext.getDocSecurity()).append("</DocSecurity>");
        if (!ext.getHyperlinkBase().isEmpty()) sb.append("<HyperlinkBase>").append(XlsxWorkbookSerializerCommon.xmlText(ext.getHyperlinkBase())).append("</HyperlinkBase>");
        if (ext.getScaleCrop() != null) sb.append("<ScaleCrop>").append(ext.getScaleCrop() ? "true" : "false").append("</ScaleCrop>");
        if (ext.getLinksUpToDate() != null) sb.append("<LinksUpToDate>").append(ext.getLinksUpToDate() ? "true" : "false").append("</LinksUpToDate>");
        if (ext.getSharedDoc() != null) sb.append("<SharedDoc>").append(ext.getSharedDoc() ? "true" : "false").append("</SharedDoc>");
        sb.append("</Properties>");
        return sb.toString().getBytes(StandardCharsets.UTF_8);
    }

    private static final DateTimeFormatter ISO_DT = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss'Z'");

    /**
     * Formats the iso.
     * @param dt dt
     * @return the computed result
     */
    private static String formatIso(LocalDateTime dt) {
        return dt.format(ISO_DT);
    }

    // =========================================================================
    // LOAD
    // =========================================================================

    /**
     * Loads the current content.
     * @param stream stream to apply
     * @param options options to apply
     * @param diagnostics diagnostics
     * @return the requested result
     */
    public static WorkbookModel load(InputStream stream, LoadOptions options, LoadDiagnostics diagnostics) throws IOException {
        byte[] data = stream.readAllBytes();

        Map<String, byte[]> entries = new HashMap<>();
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(data), StandardCharsets.UTF_8)) {
            ZipEntry ze;
            while ((ze = zis.getNextEntry()) != null) {
                entries.put(ze.getName(), zis.readAllBytes());
            }
        } catch (IOException e) {
            throw new InvalidFileFormatException("The workbook is not a valid XLSX zip package.");
        }

        byte[] wbBytes = entries.get("xl/workbook.xml");
        if (wbBytes == null) throw new InvalidFileFormatException("The package does not contain /xl/workbook.xml.");

        Document wbDoc = XlsxWorkbookArchiveHelpers.parse(wbBytes);
        WorkbookModel model = new WorkbookModel();

        // workbookPr — date1904 + workbook properties
        NodeList prList = wbDoc.getElementsByTagNameNS("*", "workbookPr");
        if (prList.getLength() > 0) {
            Element pr = (Element) prList.item(0);
            model.getSettings().setDateSystem("1".equals(pr.getAttribute("date1904"))
                    ? DateSystem.MAC_1904 : DateSystem.WINDOWS_1900);
            WorkbookPropertiesModel wp = model.getProperties();
            String codeName = pr.getAttribute("codeName");
            if (!codeName.isEmpty()) wp.setCodeName(codeName);
            String showObjects = pr.getAttribute("showObjects");
            if (!showObjects.isEmpty()) wp.setShowObjects(showObjects);
            if ("1".equals(pr.getAttribute("filterPrivacy"))) wp.setFilterPrivacy(true);
            String sbut = pr.getAttribute("showBorderUnselectedTables");
            if ("0".equals(sbut)) wp.setShowBorderUnselectedTables(false);
            String sia = pr.getAttribute("showInkAnnotation");
            if ("0".equals(sia)) wp.setShowInkAnnotation(false);
            if ("1".equals(pr.getAttribute("backupFile"))) wp.setBackupFile(true);
            String selv = pr.getAttribute("saveExternalLinkValues");
            if ("0".equals(selv)) wp.setSaveExternalLinkValues(false);
            String updateLinks = pr.getAttribute("updateLinks");
            if (!updateLinks.isEmpty()) wp.setUpdateLinks(updateLinks);
            if ("1".equals(pr.getAttribute("hidePivotFieldList"))) wp.setHidePivotFieldList(true);
            String dtv = pr.getAttribute("defaultThemeVersion");
            if (!dtv.isEmpty()) { try { wp.setDefaultThemeVersion(Integer.parseInt(dtv)); } catch (NumberFormatException ignored) {} }
        }

        // active tab
        NodeList bvList = wbDoc.getElementsByTagNameNS("*", "workbookView");
        if (bvList.getLength() > 0) {
            String at = ((Element) bvList.item(0)).getAttribute("activeTab");
            if (!at.isEmpty()) model.setActiveSheetIndex(Integer.parseInt(at));
        }

        // workbook.xml.rels — rId → target
        Map<String, String> wbRels = XlsxWorkbookArchiveHelpers.loadRels(entries, "xl/_rels/workbook.xml.rels");

        // sheets from workbook.xml
        NodeList sheetNodes = wbDoc.getElementsByTagNameNS("*", "sheet");
        List<String[]> sheetDefs = new ArrayList<>(); // [name, rId, state]
        for (int i = 0; i < sheetNodes.getLength(); i++) {
            Element s = (Element) sheetNodes.item(i);
            String rId = s.getAttribute("r:id");
            if (rId.isEmpty()) rId = s.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
            sheetDefs.add(new String[]{s.getAttribute("name"), rId, s.getAttribute("state")});
        }

        // shared strings
        String[] sstArr = loadSharedStrings(entries);

        // style values (indexed by cellXfs position)
        XlsxWorkbookStyles.StyleLoadResult styleLoadResult = XlsxWorkbookStyles.loadStyleValues(entries);

        // defined names
        Map<Integer, String[]> definedNames = XlsxWorkbookDefinedNames.loadDefinedNames(wbDoc);
        XlsxWorkbookDefinedNames.loadUserDefinedNames(wbDoc, model);

        // parse each worksheet
        model.getWorksheets().clear();
        for (int i = 0; i < sheetDefs.size(); i++) {
            String[] def = sheetDefs.get(i);
            String target = wbRels.getOrDefault(def[1], "worksheets/sheet" + (i + 1) + ".xml");
            String entryPath = "xl/" + target;
            byte[] wsBytes = entries.get(entryPath);
            WorksheetModel ws = new WorksheetModel(def[0]);
            if ("hidden".equals(def[2])) ws.setVisibility(SheetVisibility.HIDDEN);
            else if ("veryHidden".equals(def[2])) ws.setVisibility(SheetVisibility.VERY_HIDDEN);
            int lastSlash = entryPath.lastIndexOf('/');
            String relsPath = entryPath.substring(0, lastSlash + 1) + "_rels/"
                    + entryPath.substring(lastSlash + 1) + ".rels";
            Map<String, String> sheetRels = XlsxWorkbookArchiveHelpers.loadRels(entries, relsPath);
            if (wsBytes != null) parseWorksheet(ws, wsBytes, sstArr, styleLoadResult, model, sheetRels, entries);
            String[] dn = definedNames.get(i);
            if (dn != null) {
                if (dn[0] != null) ws.getPageSetup().setPrintArea(XlsxWorkbookSerializerCommon.stripSheetPrefix(dn[0]));
                if (dn[1] != null) ws.getPageSetup().setPrintTitleRows(XlsxWorkbookSerializerCommon.stripSheetPrefix(dn[1]));
                if (dn[2] != null) ws.getPageSetup().setPrintTitleColumns(XlsxWorkbookSerializerCommon.stripSheetPrefix(dn[2]));
            }
            model.getWorksheets().add(ws);
        }

        // xl/theme/theme1.xml — preserve verbatim so theme-based fill/font colors round-trip correctly
        byte[] themeBytes = entries.get("xl/theme/theme1.xml");
        if (themeBytes != null) model.setRawThemeXml(themeBytes);

        // preserve fonts[0] raw XML so the default font (family/scheme/color theme) round-trips correctly
        if (styleLoadResult.rawDefaultFontXml != null)
            model.setRawDefaultFontXml(styleLoadResult.rawDefaultFontXml);

        // docProps/core.xml
        byte[] coreBytes = entries.get("docProps/core.xml");
        if (coreBytes != null) loadCoreProperties(model.getDocumentProperties().getCore(), coreBytes);

        // docProps/app.xml
        byte[] appBytes = entries.get("docProps/app.xml");
        if (appBytes != null) loadAppProperties(model.getDocumentProperties().getExtended(), appBytes);

        return model;
    }

    /**
     * Loads the core properties.
     * @param core core
     * @param bytes bytes
     */
    private static void loadCoreProperties(DocumentPropertiesModel.CoreDocumentPropertiesModel core, byte[] bytes) {
        Document doc = XlsxWorkbookArchiveHelpers.parse(bytes);
        core.setTitle(text(doc, "title"));
        core.setSubject(text(doc, "subject"));
        core.setCreator(text(doc, "creator"));
        core.setKeywords(text(doc, "keywords"));
        core.setDescription(text(doc, "description"));
        core.setLastModifiedBy(text(doc, "lastModifiedBy"));
        core.setRevision(text(doc, "revision"));
        core.setCategory(text(doc, "category"));
        core.setContentStatus(text(doc, "contentStatus"));
        String created = text(doc, "created");
        // Handle the relevant branch before the state changes.
        if (!created.isEmpty()) core.setCreated(parseIso(created));
        String modified = text(doc, "modified");
        if (!modified.isEmpty()) core.setModified(parseIso(modified));
    }

    /**
     * Loads the app properties.
     * @param ext ext
     * @param bytes bytes
     */
    private static void loadAppProperties(DocumentPropertiesModel.ExtendedDocumentPropertiesModel ext, byte[] bytes) {
        Document doc = XlsxWorkbookArchiveHelpers.parse(bytes);
        ext.setApplication(text(doc, "Application"));
        ext.setAppVersion(text(doc, "AppVersion"));
        ext.setCompany(text(doc, "Company"));
        ext.setManager(text(doc, "Manager"));
        ext.setHyperlinkBase(text(doc, "HyperlinkBase"));
        String ds = text(doc, "DocSecurity");
        // Handle the relevant branch before the state changes.
        if (!ds.isEmpty()) { try { ext.setDocSecurity(Integer.parseInt(ds)); } catch (NumberFormatException ignored) {} }
        String sc = text(doc, "ScaleCrop");
        if (!sc.isEmpty()) ext.setScaleCrop("true".equalsIgnoreCase(sc));
        String lu = text(doc, "LinksUpToDate");
        if (!lu.isEmpty()) ext.setLinksUpToDate("true".equalsIgnoreCase(lu));
        String sd = text(doc, "SharedDoc");
        if (!sd.isEmpty()) ext.setSharedDoc("true".equalsIgnoreCase(sd));
    }

    /** Returns the text content of the first element matching localName (any namespace), or "". */
    private static String text(Document doc, String localName) {
        NodeList nl = doc.getElementsByTagNameNS("*", localName);
        // Handle the relevant branch before the state changes.
        if (nl.getLength() == 0) return "";
        String s = nl.item(0).getTextContent();
        return s != null ? s.trim() : "";
    }

    /**
     * Parses iso.
     * @param s s
     * @return the computed result
     */
    private static LocalDateTime parseIso(String s) {
        try {
            // Normalise: remove trailing Z/z, truncate to seconds
            String n = s.trim().replace("Z", "").replace("z", "");
            if (n.length() > 19) n = n.substring(0, 19);
            return LocalDateTime.parse(n);
        } catch (Exception e) { return null; }
    }

    /**
     * Loads the shared strings.
     * @param entries entries
     * @return the requested result
     */
    private static String[] loadSharedStrings(Map<String, byte[]> entries) {
        byte[] bytes = entries.get("xl/sharedStrings.xml");
        // Handle the relevant branch before the state changes.
        if (bytes == null) return new String[0];
        Document doc = XlsxWorkbookArchiveHelpers.parse(bytes);
        NodeList siNodes = doc.getElementsByTagNameNS("*", "si");
        String[] arr = new String[siNodes.getLength()];
        for (int i = 0; i < siNodes.getLength(); i++) {
            Element si = (Element) siNodes.item(i);
            StringBuilder sb = new StringBuilder();
            NodeList tNodes = si.getElementsByTagNameNS("*", "t");
            for (int j = 0; j < tNodes.getLength(); j++) sb.append(tNodes.item(j).getTextContent());
            arr[i] = sb.toString();
        }
        return arr;
    }

    /**
     * Parses worksheet.
     * @param ws ws
     * @param bytes bytes
     * @param sst sst
     * @param styleLoadResult style load result
     * @param wb wb
     * @param sheetRels sheet rels
     */
    private static void parseWorksheet(WorksheetModel ws, byte[] bytes, String[] sst,
                                       XlsxWorkbookStyles.StyleLoadResult styleLoadResult, WorkbookModel wb,
                                       Map<String, String> sheetRels,
                                       Map<String, byte[]> allEntries) {
        Set<Integer> dateStyleIndices = styleLoadResult.dateStyleIndices;
        Document doc = XlsxWorkbookArchiveHelpers.parse(bytes);

        XlsxWorkbookWorksheetViews.loadTabColor(ws, doc);
        XlsxWorkbookWorksheetViews.loadSheetView(ws, doc);
        XlsxWorkbookWorksheetViews.loadCols(ws, doc);

        // sheetData — rows and cells
        NodeList rowNodes = doc.getElementsByTagNameNS("*", "row");
        for (int i = 0; i < rowNodes.getLength(); i++) {
            Element row = (Element) rowNodes.item(i);
            int rowIdx = XlsxWorkbookSerializerCommon.parseInt(row.getAttribute("r"), 1) - 1;
            String ht = row.getAttribute("ht");
            String hidden = row.getAttribute("hidden");
            String rowOl = row.getAttribute("outlineLevel");
            String rowCollapsed = row.getAttribute("collapsed");
            if (!ht.isEmpty() || "1".equals(hidden) || !rowOl.isEmpty() || "1".equals(rowCollapsed)) {
                RowModel rm = ws.getRows().computeIfAbsent(rowIdx, k -> new RowModel());
                if (!ht.isEmpty()) rm.setHeight(Double.parseDouble(ht));
                if ("1".equals(hidden)) rm.setHidden(true);
                if (!rowOl.isEmpty()) { try { rm.setOutlineLevel(Integer.parseInt(rowOl)); } catch (NumberFormatException ignored) {} }
                if ("1".equals(rowCollapsed)) rm.setCollapsed(true);
            }
            NodeList cNodes = row.getElementsByTagNameNS("*", "c");
            for (int j = 0; j < cNodes.getLength(); j++) {
                Element c = (Element) cNodes.item(j);
                String ref = c.getAttribute("r");
                if (ref.isEmpty()) continue;
                CellAddress addr = XlsxWorkbookSerializerCommon.parseRef(ref);
                String type = c.getAttribute("t");
                int styleIdx = XlsxWorkbookSerializerCommon.parseInt(c.getAttribute("s"), 0);
                NodeList vNodes = c.getElementsByTagNameNS("*", "v");
                NodeList fNodes = c.getElementsByTagNameNS("*", "f");
                String vText = vNodes.getLength() > 0 ? vNodes.item(0).getTextContent() : "";
                String fText = fNodes.getLength() > 0 ? fNodes.item(0).getTextContent() : "";

                CellRecord rec = ws.getCells().computeIfAbsent(addr, k -> new CellRecord());

                if (styleIdx < styleLoadResult.cellStyles.size()) {
                    rec.setStyle(styleLoadResult.cellStyles.get(styleIdx));
                }

                if (!fText.isEmpty()) {
                    rec.setFormula(fText);
                    rec.setKind(CellValueKind.FORMULA);
                    if (!vText.isEmpty()) {
                        try {
                            double numVal = Double.parseDouble(vText);
                            if (!vText.contains(".") && !vText.contains("e") && !vText.contains("E")
                                    && numVal >= Integer.MIN_VALUE && numVal <= Integer.MAX_VALUE) {
                                rec.setValue((int) numVal);
                            } else {
                                rec.setValue(numVal);
                            }
                        } catch (NumberFormatException ignored) {}
                    }
                } else if ("s".equals(type)) {
                    int idx = XlsxWorkbookSerializerCommon.parseInt(vText, -1);
                    String sv = idx >= 0 && idx < sst.length ? sst[idx] : "";
                    rec.setValue(sv);
                    rec.setKind(CellValueKind.STRING);
                } else if ("b".equals(type)) {
                    rec.setValue("1".equals(vText));
                    rec.setKind(CellValueKind.BOOLEAN);
                } else if ("str".equals(type) || "inlineStr".equals(type)) {
                    NodeList tNodes = c.getElementsByTagNameNS("*", "t");
                    String s = tNodes.getLength() > 0 ? tNodes.item(0).getTextContent() : vText;
                    rec.setValue(s);
                    rec.setKind(CellValueKind.STRING);
                } else if (!vText.isEmpty()) {
                    double num;
                    try { num = Double.parseDouble(vText); }
                    catch (NumberFormatException ignored) { rec.setValue(vText); rec.setKind(CellValueKind.STRING); continue; }
                    if (dateStyleIndices.contains(styleIdx)) {
                        LocalDateTime dt = DateSerialConverter.fromSerial(num, wb.getSettings().getDateSystem());
                        rec.setValue(dt);
                        rec.setKind(CellValueKind.DATE_TIME);
                    } else {
                        if (!vText.contains(".") && !vText.contains("e") && !vText.contains("E")) {
                            if (num >= Integer.MIN_VALUE && num <= Integer.MAX_VALUE) {
                                rec.setValue((int) num);
                            } else {
                                rec.setValue((long) num);
                            }
                        } else {
                            rec.setValue(num);
                        }
                        rec.setKind(CellValueKind.NUMBER);
                    }
                }
            }
        }

        // mergeCells
        NodeList mcNodes = doc.getElementsByTagNameNS("*", "mergeCell");
        for (int i = 0; i < mcNodes.getLength(); i++) {
            String mref = ((Element) mcNodes.item(i)).getAttribute("ref");
            String[] parts = mref.split(":");
            if (parts.length == 2) {
                CellAddress a1 = XlsxWorkbookSerializerCommon.parseRef(parts[0]);
                CellAddress a2 = XlsxWorkbookSerializerCommon.parseRef(parts[1]);
                ws.getMergeRegions().add(new MergeRegion(a1.getRowIndex(), a1.getColumnIndex(),
                        a2.getRowIndex() - a1.getRowIndex() + 1, a2.getColumnIndex() - a1.getColumnIndex() + 1));
            }
        }

        XlsxWorkbookConditionalFormatting.loadConditionalFormattings(ws, doc, styleLoadResult.dxfStyles);
        XlsxWorkbookAutoFilter.loadAutoFilter(ws, doc);
        XlsxWorkbookWorksheetProtection.loadSheetProtection(ws, doc);
        XlsxWorkbookPageSetup.loadPageSetup(ws, doc);
        XlsxWorkbookValidations.loadValidations(ws, doc);
        XlsxWorkbookHyperlinks.loadHyperlinks(ws, doc, sheetRels);
        XlsxWorkbookDrawings.loadDrawings(ws, allEntries, sheetRels);

        // Comments — two passes: first load comment text, then update visibility from VML
        byte[] vmlBytes = null;
        for (Map.Entry<String, String> rel : sheetRels.entrySet()) {
            String target = rel.getValue();
            if (target != null && target.contains("comments")) {
                String partPath = resolveRelTarget("xl/worksheets/sheet1.xml", target);
                byte[] cBytes = allEntries.get(partPath);
                if (cBytes != null) XlsxWorkbookComments.loadComments(ws, XlsxWorkbookArchiveHelpers.parse(cBytes));
            }
            if (target != null && target.contains("vmlDrawing")) {
                String partPath = resolveRelTarget("xl/worksheets/sheet1.xml", target);
                vmlBytes = allEntries.get(partPath);
            }
        }
        if (vmlBytes != null) XlsxWorkbookComments.loadVmlVisibility(ws, vmlBytes);

        // Tables — find via <tableParts> elements then resolve rels
        NodeList tpNodes = doc.getElementsByTagNameNS("*", "tablePart");
        for (int i = 0; i < tpNodes.getLength(); i++) {
            Element tpe = (Element) tpNodes.item(i);
            String rId = tpe.getAttribute("r:id");
            if (rId.isEmpty()) rId = tpe.getAttributeNS(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
            String target = sheetRels.get(rId);
            if (target == null) continue;
            String partPath = resolveRelTarget("xl/worksheets/sheet1.xml", target);
            byte[] tBytes = allEntries.get(partPath);
            if (tBytes != null) XlsxWorkbookTables.loadTable(ws, XlsxWorkbookArchiveHelpers.parse(tBytes));
        }
    }

    private static String resolveRelTarget(String baseEntry, String target) {
        if (target.startsWith("/")) return target.substring(1);
        String base = baseEntry.contains("/") ? baseEntry.substring(0, baseEntry.lastIndexOf('/') + 1) : "";
        String[] parts = (base + target).split("/");
        java.util.Deque<String> stack = new java.util.ArrayDeque<>();
        for (String p : parts) {
            if ("..".equals(p)) { if (!stack.isEmpty()) stack.pollLast(); }
            else if (!p.isEmpty() && !".".equals(p)) stack.addLast(p);
        }
        return String.join("/", stack);
    }
}
