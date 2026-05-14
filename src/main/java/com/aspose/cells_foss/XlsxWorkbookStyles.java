package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * Package-private helpers for building and loading XLSX styles.
 */
final class XlsxWorkbookStyles {

    // Custom date numFmtId stored in styles.xml
    static final int DATE_NUM_FMT_ID = 164;

    /**
     * Initializes a new XlsxWorkbookStyles instance.
     */
    private XlsxWorkbookStyles() {}

    // =========================================================================
    // StyleTable
    // =========================================================================

    /** Accumulates unique style components and assigns cellXfs indices. */
    static final class StyleTable {
        private final LinkedHashMap<String, Integer> fontIdx   = new LinkedHashMap<>();
        private final LinkedHashMap<String, Integer> fillIdx   = new LinkedHashMap<>();
        private final LinkedHashMap<String, Integer> borderIdx = new LinkedHashMap<>();
        private final LinkedHashMap<String, Integer> xfIdx     = new LinkedHashMap<>();

        private final List<FontValue>    fonts   = new ArrayList<>();
        private final List<Object[]>     fills   = new ArrayList<>();  // [FillPatternKind, ColorValue fg, ColorValue bg]
        private final List<BordersValue> borders = new ArrayList<>();
        private final List<int[]>        xfs     = new ArrayList<>();  // [numFmtId, fontId, fillId, borderId]
        private final List<AlignmentValue>  xfAligns = new ArrayList<>();
        private final List<ProtectionValue> xfProts  = new ArrayList<>();

        private final List<int[]>    customFmtIds   = new ArrayList<>();
        private final List<String>   customFmtCodes = new ArrayList<>();
        private final Map<String, Integer> customFmtByCode = new LinkedHashMap<>();
        private int nextFmtId = DATE_NUM_FMT_ID + 1;

        // Differential formats (dxf) for conditional formatting styles
        private final List<StyleValue>     dxfStyles = new ArrayList<>();
        private final Map<String, Integer> dxfIdx    = new LinkedHashMap<>();

        /**
         * Initializes a new StyleTable instance.
         * @param rawDefaultFontXml raw XML of the source file's fonts[0], or null to use plain Calibri 11pt
         */
        StyleTable(String rawDefaultFontXml) {
            // Default font: use source file's fonts[0] when available (preserves family/scheme/color theme)
            FontValue defFont = new FontValue();
            defFont.setName("Calibri");
            defFont.setSize(11.0);
            if (rawDefaultFontXml != null) defFont.setRawFontXml(rawDefaultFontXml);
            fonts.add(defFont);
            fontIdx.put(fontKey(defFont), 0);

            // Mandatory fills: index 0 = none, index 1 = gray125
            fills.add(new Object[]{FillPatternKind.NONE, null, null});
            fillIdx.put(fillKey(FillPatternKind.NONE, null, null), 0);
            fills.add(new Object[]{FillPatternKind.GRAY_125, null, null});
            fillIdx.put(fillKey(FillPatternKind.GRAY_125, null, null), 1);

            // Default border: empty
            borders.add(new BordersValue());
            borderIdx.put(bordersKey(new BordersValue()), 0);

            // Always include the date custom numFmt at DATE_NUM_FMT_ID
            customFmtIds.add(new int[]{DATE_NUM_FMT_ID});
            customFmtCodes.add("m/d/yyyy h:mm");
            customFmtByCode.put("m/d/yyyy h:mm", DATE_NUM_FMT_ID);

            // Default cellXf: everything zero
            xfs.add(new int[]{0, 0, 0, 0});
            xfAligns.add(new AlignmentValue());
            xfProts.add(new ProtectionValue());
            xfIdx.put(xfKey(0, 0, 0, 0, new AlignmentValue(), new ProtectionValue()), 0);
        }

        /** Returns (or adds) the cellXfs index for the given StyleValue. */
        int register(StyleValue sv, boolean isDate) {
            FontValue fv = sv.getFont() != null ? sv.getFont() : new FontValue();
            // Lazily create the backing record the first time this path is used.
            int fId = fontIdx.computeIfAbsent(fontKey(fv), k -> {
                fonts.add(fv); return fonts.size() - 1;
            });

            FillPatternKind pat = sv.getPattern() != null ? sv.getPattern() : FillPatternKind.NONE;
            ColorValue fg = sv.getForegroundColor(), bg = sv.getBackgroundColor();
            String rawFillXml = sv.getRawFillXml();
            int fillId = fillIdx.computeIfAbsent(fillKey(pat, fg, bg, rawFillXml), k -> {
                fills.add(new Object[]{pat, fg, bg, rawFillXml}); return fills.size() - 1;
            });

            BordersValue bv = sv.getBorders() != null ? sv.getBorders() : new BordersValue();
            int bId = borderIdx.computeIfAbsent(bordersKey(bv), k -> {
                borders.add(bv); return borders.size() - 1;
            });

            int numFmtId;
            NumberFormatValue nfv = sv.getNumberFormat();
            String custom = nfv != null ? nfv.getCustom() : null;
            int builtIn  = nfv != null ? nfv.getNumber() : 0;
            if (custom != null && !custom.isEmpty()) {
                numFmtId = customFmtByCode.computeIfAbsent(custom, c -> {
                    int id = nextFmtId++;
                    customFmtIds.add(new int[]{id});
                    customFmtCodes.add(c);
                    return id;
                });
            } else if (builtIn != 0) {
                numFmtId = builtIn;
            } else if (isDate) {
                numFmtId = DATE_NUM_FMT_ID;
            } else {
                numFmtId = 0;
            }

            AlignmentValue av = sv.getAlignment() != null ? sv.getAlignment() : new AlignmentValue();
            ProtectionValue pv = sv.getProtection() != null ? sv.getProtection() : new ProtectionValue();

            String key = xfKey(numFmtId, fId, fillId, bId, av, pv);
            return xfIdx.computeIfAbsent(key, k -> {
                xfs.add(new int[]{numFmtId, fId, fillId, bId});
                xfAligns.add(av);
                xfProts.add(pv);
                return xfs.size() - 1;
            });
        }

        /** Registers a differential style (for conditional formatting). Returns 0-based dxf index, or -1 if style is default. */
        int registerDifferentialStyle(StyleValue sv) {
            // Handle the relevant branch before the state changes.
            if (sv == null || isDefaultStyle(sv)) return -1;
            String key = dxfStyleKey(sv);
            return dxfIdx.computeIfAbsent(key, k -> {
                dxfStyles.add(sv);
                return dxfStyles.size() - 1;
            });
        }

        /**
         * Indicates whether default style.
         * @param sv sv
         * @return true when the condition is satisfied
         */
        private static boolean isDefaultStyle(StyleValue sv) {
            FontValue fv = sv.getFont();
            boolean fontDefault = fv == null || (!fv.getBold() && !fv.getItalic() && !fv.getUnderline()
                    && !fv.getStrikeThrough() && fv.getColor() == null);
            boolean fillDefault = sv.getPattern() == null || sv.getPattern() == FillPatternKind.NONE;
            BordersValue bv = sv.getBorders();
            boolean borderDefault = bv == null || (isNullBorder(bv.getLeft()) && isNullBorder(bv.getRight())
                    && isNullBorder(bv.getTop()) && isNullBorder(bv.getBottom()) && isNullBorder(bv.getDiagonal()));
            NumberFormatValue nf = sv.getNumberFormat();
            boolean numFmtDefault = nf == null || (nf.getNumber() == 0 && (nf.getCustom() == null || nf.getCustom().isEmpty()));
            AlignmentValue av = sv.getAlignment();
            boolean alignDefault = av == null || !hasNonDefaultAlignment(av);
            return fontDefault && fillDefault && borderDefault && numFmtDefault && alignDefault;
        }

        /**
         * Indicates whether null border.
         * @param bsv bsv
         * @return true when the condition is satisfied
         */
        private static boolean isNullBorder(BorderSideValue bsv) {
            return bsv == null || bsv.getStyle() == null;
        }

        /**
         * Processes dxf style key.
         * @param sv sv
         * @return the computed result
         */
        private static String dxfStyleKey(StyleValue sv) {
            FontValue fv = sv.getFont() != null ? sv.getFont() : new FontValue();
            FillPatternKind p = sv.getPattern() != null ? sv.getPattern() : FillPatternKind.NONE;
            ColorValue fg = sv.getForegroundColor(), bg = sv.getBackgroundColor();
            BordersValue bv = sv.getBorders() != null ? sv.getBorders() : new BordersValue();
            NumberFormatValue nf = sv.getNumberFormat();
            int numFmtId = nf != null ? nf.getNumber() : 0;
            String custom = nf != null ? nf.getCustom() : null;
            AlignmentValue av = sv.getAlignment() != null ? sv.getAlignment() : new AlignmentValue();
            ProtectionValue pv = sv.getProtection() != null ? sv.getProtection() : new ProtectionValue();
            return "DXF|" + fontKey(fv) + "|" + fillKey(p, fg, bg) + "|" + bordersKey(bv)
                 + "|" + numFmtId + "|" + (custom != null ? custom : "")
                 + "|" + av.getHorizontal() + "|" + av.getVertical() + "|" + av.getWrapText()
                 + "|" + pv.getIsLocked() + "|" + pv.getIsHidden();
        }

        /**
         * Builds the styles xml bytes.
         * @return the requested result
         */
        byte[] buildStylesXmlBytes() {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
            sb.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

            // numFmts
            sb.append("<numFmts count=\"").append(customFmtIds.size()).append("\">");
            for (int i = 0; i < customFmtIds.size(); i++)
                sb.append("<numFmt numFmtId=\"").append(customFmtIds.get(i)[0])
                  .append("\" formatCode=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(customFmtCodes.get(i))).append("\"/>");
            sb.append("</numFmts>");

            // fonts
            sb.append("<fonts count=\"").append(fonts.size()).append("\">");
            for (FontValue fv : fonts) {
                if (fv.getRawFontXml() != null) {
                    sb.append(fv.getRawFontXml());
                } else {
                    sb.append("<font>");
                    if (fv.getBold())          sb.append("<b/>");
                    if (fv.getItalic())        sb.append("<i/>");
                    if (fv.getStrikeThrough()) sb.append("<strike/>");
                    if (fv.getUnderline())     sb.append("<u/>");
                    sb.append("<sz val=\"").append(XlsxWorkbookSerializerCommon.fmt(fv.getSize())).append("\"/>");
                    if (fv.getColor() != null)
                        sb.append("<color rgb=\"").append(XlsxWorkbookSerializerCommon.colorArgb(fv.getColor())).append("\"/>");
                    sb.append("<name val=\"").append(XlsxWorkbookSerializerCommon.xmlAttr(fv.getName() != null ? fv.getName() : "Calibri")).append("\"/>");
                    sb.append("</font>");
                }
            }
            sb.append("</fonts>");

            // fills
            sb.append("<fills count=\"").append(fills.size()).append("\">");
            for (Object[] fill : fills) {
                FillPatternKind p = (FillPatternKind) fill[0];
                ColorValue fgc = (ColorValue) fill[1], bgc = (ColorValue) fill[2];
                String rawFillXml = fill.length > 3 ? (String) fill[3] : null;
                if (rawFillXml != null) {
                    sb.append(rawFillXml);
                } else {
                    sb.append("<fill><patternFill patternType=\"").append(fillPatternName(p)).append("\">");
                    if (fgc != null) sb.append("<fgColor rgb=\"").append(XlsxWorkbookSerializerCommon.colorArgb(fgc)).append("\"/>");
                    if (bgc != null) sb.append("<bgColor rgb=\"").append(XlsxWorkbookSerializerCommon.colorArgb(bgc)).append("\"/>");
                    sb.append("</patternFill></fill>");
                }
            }
            sb.append("</fills>");

            // borders
            sb.append("<borders count=\"").append(borders.size()).append("\">");
            for (BordersValue bv : borders) {
                sb.append("<border");
                if (bv.getDiagonalUp())   sb.append(" diagonalUp=\"1\"");
                if (bv.getDiagonalDown()) sb.append(" diagonalDown=\"1\"");
                sb.append(">");
                appendBorderSide(sb, "left",     bv.getLeft());
                appendBorderSide(sb, "right",    bv.getRight());
                appendBorderSide(sb, "top",      bv.getTop());
                appendBorderSide(sb, "bottom",   bv.getBottom());
                appendBorderSide(sb, "diagonal", bv.getDiagonal());
                sb.append("</border>");
            }
            sb.append("</borders>");

            // dxfs (differential formats for conditional formatting)
            if (!dxfStyles.isEmpty()) {
                sb.append("<dxfs count=\"").append(dxfStyles.size()).append("\">");
                for (StyleValue dsv : dxfStyles) {
                    sb.append("<dxf>");
                    FontValue dfv = dsv.getFont();
                    if (dfv != null && (dfv.getBold() || dfv.getItalic() || dfv.getUnderline()
                            || dfv.getStrikeThrough() || dfv.getColor() != null)) {
                        sb.append("<font>");
                        if (dfv.getBold())          sb.append("<b/>");
                        if (dfv.getItalic())        sb.append("<i/>");
                        if (dfv.getStrikeThrough()) sb.append("<strike/>");
                        if (dfv.getUnderline())     sb.append("<u/>");
                        if (dfv.getColor() != null)
                            sb.append("<color rgb=\"").append(XlsxWorkbookSerializerCommon.colorArgb(dfv.getColor())).append("\"/>");
                        sb.append("</font>");
                    }
                    FillPatternKind dpat = dsv.getPattern();
                    ColorValue dfg = dsv.getForegroundColor(), dbg = dsv.getBackgroundColor();
                    if (dpat != null && dpat != FillPatternKind.NONE) {
                        sb.append("<fill><patternFill patternType=\"").append(fillPatternName(dpat)).append("\">");
                        if (dfg != null) sb.append("<fgColor rgb=\"").append(XlsxWorkbookSerializerCommon.colorArgb(dfg)).append("\"/>");
                        if (dbg != null) sb.append("<bgColor rgb=\"").append(XlsxWorkbookSerializerCommon.colorArgb(dbg)).append("\"/>");
                        sb.append("</patternFill></fill>");
                    }
                    BordersValue dbv = dsv.getBorders();
                    if (dbv != null && (!isNullBorder(dbv.getLeft()) || !isNullBorder(dbv.getRight())
                            || !isNullBorder(dbv.getTop()) || !isNullBorder(dbv.getBottom())
                            || !isNullBorder(dbv.getDiagonal()))) {
                        sb.append("<border>");
                        appendBorderSide(sb, "left",     dbv.getLeft());
                        appendBorderSide(sb, "right",    dbv.getRight());
                        appendBorderSide(sb, "top",      dbv.getTop());
                        appendBorderSide(sb, "bottom",   dbv.getBottom());
                        appendBorderSide(sb, "diagonal", dbv.getDiagonal());
                        sb.append("</border>");
                    }
                    sb.append("</dxf>");
                }
                sb.append("</dxfs>");
            }

            // cellStyleXfs
            sb.append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

            // cellXfs
            sb.append("<cellXfs count=\"").append(xfs.size()).append("\">");
            for (int i = 0; i < xfs.size(); i++) {
                int[] xf = xfs.get(i);
                AlignmentValue av  = xfAligns.get(i);
                ProtectionValue pv = xfProts.get(i);
                boolean hasAlign = hasNonDefaultAlignment(av);
                boolean hasProt  = !pv.getIsLocked() || pv.getIsHidden();
                sb.append("<xf numFmtId=\"").append(xf[0])
                  .append("\" fontId=\"").append(xf[1])
                  .append("\" fillId=\"").append(xf[2])
                  .append("\" borderId=\"").append(xf[3])
                  .append("\" xfId=\"0\"");
                if (xf[0] != 0) sb.append(" applyNumberFormat=\"1\"");
                if (xf[1] != 0) sb.append(" applyFont=\"1\"");
                if (xf[2] != 0) sb.append(" applyFill=\"1\"");
                if (xf[3] != 0) sb.append(" applyBorder=\"1\"");
                if (hasAlign)   sb.append(" applyAlignment=\"1\"");
                if (hasProt)    sb.append(" applyProtection=\"1\"");
                if (hasAlign || hasProt) {
                    sb.append(">");
                    if (hasAlign) appendAlignmentElement(sb, av);
                    if (hasProt)  appendProtectionElement(sb, pv);
                    sb.append("</xf>");
                } else {
                    sb.append("/>");
                }
            }
            sb.append("</cellXfs>");

            sb.append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
            sb.append("</styleSheet>");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        }
    }

    // =========================================================================
    // StyleLoadResult
    // =========================================================================

    /**
     * Represents the StyleLoadResult component.
     */
    static final class StyleLoadResult {
        final Set<Integer> dateStyleIndices;
        final List<StyleValue> cellStyles;
        final List<StyleValue> dxfStyles;
        final String rawDefaultFontXml;
        /**
         * Initializes a new StyleLoadResult instance.
         * @param d d
         * @param s s
         * @param dxf dxf
         * @param rawDefaultFontXml raw XML of fonts[0] from source styles.xml
         */
        StyleLoadResult(Set<Integer> d, List<StyleValue> s, List<StyleValue> dxf, String rawDefaultFontXml) {
            dateStyleIndices = d; cellStyles = s; dxfStyles = dxf;
            this.rawDefaultFontXml = rawDefaultFontXml;
        }
    }

    // =========================================================================
    // Build helpers
    // =========================================================================

    /**
     * Builds the style table.
     * @param model model
     * @return the requested result
     */
    static StyleTable buildStyleTable(WorkbookModel model) {
        StyleTable t = new StyleTable(model.getRawDefaultFontXml());
        // Walk the current collection so every entry is processed consistently.
        for (WorksheetModel ws : model.getWorksheets()) {
            for (Map.Entry<CellAddress, CellRecord> e : ws.getCells().entrySet()) {
                if (XlsxWorkbookSerializerCommon.shouldPersist(model.getDefaultStyle(), e.getValue())) {
                    boolean isDate = e.getValue().getKind() == CellValueKind.DATE_TIME;
                    t.register(e.getValue().getStyle(), isDate);
                }
            }
            for (ConditionalFormattingModel cf : ws.getConditionalFormattings()) {
                for (FormatConditionModel cond : cf.getConditions()) {
                    t.registerDifferentialStyle(cond.getStyle());
                }
            }
        }
        return t;
    }

    // =========================================================================
    // Load helpers
    // =========================================================================

    /**
     * Loads the style values.
     * @param entries entries
     * @return the requested result
     */
    static StyleLoadResult loadStyleValues(Map<String, byte[]> entries) {
        Set<Integer> dateIndices = new HashSet<>();
        List<StyleValue> cellStyles = new ArrayList<>();

        byte[] bytes = entries.get("xl/styles.xml");
        if (bytes == null) return new StyleLoadResult(dateIndices, cellStyles, new ArrayList<>(), null);

        Document doc = XlsxWorkbookArchiveHelpers.parse(bytes);

        // Gather custom numFmt codes and identify date fmtIds
        Set<Integer> dateFmtIds = new HashSet<>(Arrays.asList(14,15,16,17,18,19,20,21,22,45,46,47));
        Map<Integer, String> numFmtCodeMap = new HashMap<>();
        NodeList numFmtsEl = doc.getElementsByTagNameNS("*", "numFmts");
        if (numFmtsEl.getLength() > 0) {
            NodeList nfNodes = ((Element) numFmtsEl.item(0)).getElementsByTagNameNS("*", "numFmt");
            for (int i = 0; i < nfNodes.getLength(); i++) {
                Element nf = (Element) nfNodes.item(i);
                int id = XlsxWorkbookSerializerCommon.parseInt(nf.getAttribute("numFmtId"), 0);
                String code = nf.getAttribute("formatCode");
                numFmtCodeMap.put(id, code);
                if (XlsxWorkbookSerializerCommon.isDateFmt(code.toLowerCase())) dateFmtIds.add(id);
            }
        }

        // Read fonts
        List<FontValue> fontList = new ArrayList<>();
        NodeList fontsEl = doc.getElementsByTagNameNS("*", "fonts");
        if (fontsEl.getLength() > 0) {
            NodeList kids = ((Element) fontsEl.item(0)).getChildNodes();
            for (int i = 0; i < kids.getLength(); i++) {
                if (!(kids.item(i) instanceof Element)) continue;
                Element fontEl = (Element) kids.item(i);
                FontValue fv = new FontValue();
                Element nameEl = findChildEl(fontEl, "name");
                fv.setName(nameEl != null ? nameEl.getAttribute("val") : "Calibri");
                Element szEl = findChildEl(fontEl, "sz");
                fv.setSize(szEl != null ? XlsxWorkbookSerializerCommon.parseDouble(szEl.getAttribute("val"), 11.0) : 11.0);
                fv.setBold(findChildEl(fontEl, "b") != null);
                fv.setItalic(findChildEl(fontEl, "i") != null);
                fv.setUnderline(findChildEl(fontEl, "u") != null);
                fv.setStrikeThrough(findChildEl(fontEl, "strike") != null);
                Element colorEl = findChildEl(fontEl, "color");
                boolean hasNonRgbColor = false;
                if (colorEl != null) {
                    String rgb = colorEl.getAttribute("rgb");
                    if (rgb.length() == 8) {
                        fv.setColor(parseArgbColor(rgb));
                    } else {
                        hasNonRgbColor = true; // theme/tint/indexed color
                    }
                }
                boolean hasExtraAttrs = findChildEl(fontEl, "family") != null
                    || findChildEl(fontEl, "scheme") != null
                    || findChildEl(fontEl, "charset") != null
                    || findChildEl(fontEl, "vertAlign") != null;
                if (hasNonRgbColor || hasExtraAttrs)
                    fv.setRawFontXml(fontElementToXml(fontEl));
                fontList.add(fv);
            }
        }
        if (fontList.isEmpty()) { FontValue d = new FontValue(); d.setName("Calibri"); d.setSize(11.0); fontList.add(d); }

        // Read fills — 4th element is raw XML for fills using theme/tint/indexed colors
        List<Object[]> fillList = new ArrayList<>();
        NodeList fillsEl = doc.getElementsByTagNameNS("*", "fills");
        if (fillsEl.getLength() > 0) {
            NodeList kids = ((Element) fillsEl.item(0)).getChildNodes();
            for (int i = 0; i < kids.getLength(); i++) {
                if (!(kids.item(i) instanceof Element)) continue;
                Element fillEl = (Element) kids.item(i);
                Element pf = findChildEl(fillEl, "patternFill");
                if (pf == null) { fillList.add(new Object[]{FillPatternKind.NONE, null, null, null}); continue; }
                FillPatternKind pat = parseFillPatternType(pf.getAttribute("patternType"));
                Element fgEl = findChildEl(pf, "fgColor"), bgEl = findChildEl(pf, "bgColor");
                ColorValue fg = fgEl != null ? parseArgbColor(fgEl.getAttribute("rgb")) : null;
                ColorValue bg = bgEl != null ? parseArgbColor(bgEl.getAttribute("rgb")) : null;
                // If any color element exists but couldn't be parsed as RGB, preserve raw fill XML
                boolean hasNonRgb = (fgEl != null && fg == null) || (bgEl != null && bg == null);
                String rawFillXml = hasNonRgb ? fillElementToXml(pf) : null;
                fillList.add(new Object[]{pat, fg, bg, rawFillXml});
            }
        }
        if (fillList.isEmpty()) {
            fillList.add(new Object[]{FillPatternKind.NONE, null, null, null});
            fillList.add(new Object[]{FillPatternKind.GRAY_125, null, null, null});
        }

        // Read borders
        List<BordersValue> borderList = new ArrayList<>();
        NodeList bordersEl = doc.getElementsByTagNameNS("*", "borders");
        if (bordersEl.getLength() > 0) {
            NodeList kids = ((Element) bordersEl.item(0)).getChildNodes();
            for (int i = 0; i < kids.getLength(); i++) {
                if (!(kids.item(i) instanceof Element)) continue;
                Element borderEl = (Element) kids.item(i);
                BordersValue bv = new BordersValue();
                bv.setLeft(readBorderSideEl(findChildEl(borderEl, "left")));
                bv.setRight(readBorderSideEl(findChildEl(borderEl, "right")));
                bv.setTop(readBorderSideEl(findChildEl(borderEl, "top")));
                bv.setBottom(readBorderSideEl(findChildEl(borderEl, "bottom")));
                bv.setDiagonal(readBorderSideEl(findChildEl(borderEl, "diagonal")));
                bv.setDiagonalUp("1".equals(borderEl.getAttribute("diagonalUp")));
                bv.setDiagonalDown("1".equals(borderEl.getAttribute("diagonalDown")));
                borderList.add(bv);
            }
        }
        if (borderList.isEmpty()) borderList.add(new BordersValue());

        // Read cellXfs
        NodeList cellXfsEl = doc.getElementsByTagNameNS("*", "cellXfs");
        if (cellXfsEl.getLength() > 0) {
            NodeList xfNodes = ((Element) cellXfsEl.item(0)).getChildNodes();
            for (int i = 0; i < xfNodes.getLength(); i++) {
                if (!(xfNodes.item(i) instanceof Element)) continue;
                Element xf = (Element) xfNodes.item(i);
                int numFmtId = XlsxWorkbookSerializerCommon.parseInt(xf.getAttribute("numFmtId"), 0);
                int fontId   = XlsxWorkbookSerializerCommon.parseInt(xf.getAttribute("fontId"),   0);
                int fillId   = XlsxWorkbookSerializerCommon.parseInt(xf.getAttribute("fillId"),   0);
                int borderId = XlsxWorkbookSerializerCommon.parseInt(xf.getAttribute("borderId"), 0);

                if (dateFmtIds.contains(numFmtId)) dateIndices.add(i);

                StyleValue sv = new StyleValue();

                if (fontId   < fontList.size())   sv.setFont(fontList.get(fontId));
                if (fillId   < fillList.size()) {
                    Object[] fl = fillList.get(fillId);
                    sv.setPattern((FillPatternKind) fl[0]);
                    sv.setForegroundColor((ColorValue) fl[1]);
                    sv.setBackgroundColor((ColorValue) fl[2]);
                    if (fl.length > 3) sv.setRawFillXml((String) fl[3]);
                }
                if (borderId < borderList.size()) sv.setBorders(borderList.get(borderId));

                NumberFormatValue nf = new NumberFormatValue();
                nf.setNumber(numFmtId);
                if (numFmtCodeMap.containsKey(numFmtId)) nf.setCustom(numFmtCodeMap.get(numFmtId));
                sv.setNumberFormat(nf);

                Element alignEl = findChildEl(xf, "alignment");
                if (alignEl != null) {
                    AlignmentValue av = new AlignmentValue();
                    av.setHorizontal(parseHAlignStr(alignEl.getAttribute("horizontal")));
                    av.setVertical(parseVAlignStr(alignEl.getAttribute("vertical")));
                    av.setWrapText("1".equals(alignEl.getAttribute("wrapText")));
                    av.setIndentLevel(XlsxWorkbookSerializerCommon.parseInt(alignEl.getAttribute("indent"), 0));
                    av.setTextRotation(XlsxWorkbookSerializerCommon.parseInt(alignEl.getAttribute("textRotation"), 0));
                    av.setShrinkToFit("1".equals(alignEl.getAttribute("shrinkToFit")));
                    av.setReadingOrder(XlsxWorkbookSerializerCommon.parseInt(alignEl.getAttribute("readingOrder"), 0));
                    av.setRelativeIndent(XlsxWorkbookSerializerCommon.parseInt(alignEl.getAttribute("relativeIndent"), 0));
                    sv.setAlignment(av);
                }

                Element protEl = findChildEl(xf, "protection");
                if (protEl != null) {
                    ProtectionValue pv = new ProtectionValue();
                    String locked = protEl.getAttribute("locked");
                    pv.setIsLocked(locked.isEmpty() || "1".equals(locked));
                    pv.setIsHidden("1".equals(protEl.getAttribute("hidden")));
                    sv.setProtection(pv);
                }

                cellStyles.add(sv);
            }
        }

        // Read differential formats (dxfs) — used by conditional formatting
        List<StyleValue> dxfStyles = new ArrayList<>();
        NodeList dxfsEl = doc.getElementsByTagNameNS("*", "dxfs");
        if (dxfsEl.getLength() > 0) {
            NodeList dxfNodes = ((Element) dxfsEl.item(0)).getChildNodes();
            for (int i = 0; i < dxfNodes.getLength(); i++) {
                if (!(dxfNodes.item(i) instanceof Element)) continue;
                Element dxf = (Element) dxfNodes.item(i);
                StyleValue dsv = new StyleValue();
                // font
                Element dFontEl = findChildEl(dxf, "font");
                if (dFontEl != null) {
                    FontValue dfv = new FontValue();
                    dfv.setBold(findChildEl(dFontEl, "b") != null);
                    dfv.setItalic(findChildEl(dFontEl, "i") != null);
                    dfv.setUnderline(findChildEl(dFontEl, "u") != null);
                    dfv.setStrikeThrough(findChildEl(dFontEl, "strike") != null);
                    Element dColorEl = findChildEl(dFontEl, "color");
                    if (dColorEl != null) dfv.setColor(parseArgbColor(dColorEl.getAttribute("rgb")));
                    Element dSzEl = findChildEl(dFontEl, "sz");
                    dfv.setSize(dSzEl != null ? XlsxWorkbookSerializerCommon.parseDouble(dSzEl.getAttribute("val"), 11.0) : 11.0);
                    Element dNameEl = findChildEl(dFontEl, "name");
                    dfv.setName(dNameEl != null ? dNameEl.getAttribute("val") : "Calibri");
                    dsv.setFont(dfv);
                }
                // fill
                Element dFillEl = findChildEl(dxf, "fill");
                if (dFillEl != null) {
                    Element dPf = findChildEl(dFillEl, "patternFill");
                    if (dPf != null) {
                        dsv.setPattern(parseFillPatternType(dPf.getAttribute("patternType")));
                        Element dFgEl = findChildEl(dPf, "fgColor");
                        Element dBgEl = findChildEl(dPf, "bgColor");
                        if (dFgEl != null) dsv.setForegroundColor(parseArgbColor(dFgEl.getAttribute("rgb")));
                        if (dBgEl != null) dsv.setBackgroundColor(parseArgbColor(dBgEl.getAttribute("rgb")));
                    }
                }
                // border
                Element dBorderEl = findChildEl(dxf, "border");
                if (dBorderEl != null) {
                    BordersValue dbv = new BordersValue();
                    dbv.setLeft(readBorderSideEl(findChildEl(dBorderEl, "left")));
                    dbv.setRight(readBorderSideEl(findChildEl(dBorderEl, "right")));
                    dbv.setTop(readBorderSideEl(findChildEl(dBorderEl, "top")));
                    dbv.setBottom(readBorderSideEl(findChildEl(dBorderEl, "bottom")));
                    dbv.setDiagonal(readBorderSideEl(findChildEl(dBorderEl, "diagonal")));
                    dsv.setBorders(dbv);
                }
                dxfStyles.add(dsv);
            }
        }

        String rawDefaultFontXml = !fontList.isEmpty() ? fontList.get(0).getRawFontXml() : null;
        return new StyleLoadResult(dateIndices, cellStyles, dxfStyles, rawDefaultFontXml);
    }

    // =========================================================================
    // Private helpers: keys
    // =========================================================================

    /**
     * Processes font key.
     * @param fv fv
     * @return the computed result
     */
    private static String fontKey(FontValue fv) {
        if (fv.getRawFontXml() != null) return "raw|" + fv.getRawFontXml();
        return fv.getName() + "|" + fv.getSize() + "|" + fv.getBold() + "|" + fv.getItalic()
             + "|" + fv.getUnderline() + "|" + fv.getStrikeThrough()
             + "|" + (fv.getColor() == null ? "null" : XlsxWorkbookSerializerCommon.colorArgb(fv.getColor()));
    }

    /**
     * Processes fill key.
     * @param p p
     * @param fg fg
     * @param bg bg
     * @return the computed result
     */
    private static String fillKey(FillPatternKind p, ColorValue fg, ColorValue bg) {
        return fillKey(p, fg, bg, null);
    }

    private static String fillKey(FillPatternKind p, ColorValue fg, ColorValue bg, String rawFillXml) {
        if (rawFillXml != null) return "raw|" + rawFillXml;
        return p + "|" + (fg == null ? "null" : XlsxWorkbookSerializerCommon.colorArgb(fg))
             + "|" + (bg == null ? "null" : XlsxWorkbookSerializerCommon.colorArgb(bg));
    }

    /**
     * Processes bs key.
     * @param bsv bsv
     * @return the computed result
     */
    private static String bsKey(BorderSideValue bsv) {
        // Handle the relevant branch before the state changes.
        if (bsv == null) return "null";
        return (bsv.getStyle() == null ? "null" : bsv.getStyle().name())
             + ":" + (bsv.getColor() == null ? "null" : XlsxWorkbookSerializerCommon.colorArgb(bsv.getColor()));
    }

    /**
     * Processes borders key.
     * @param bv bv
     * @return the computed result
     */
    private static String bordersKey(BordersValue bv) {
        // Handle the relevant branch before the state changes.
        if (bv == null) return "empty";
        return bsKey(bv.getLeft()) + "|" + bsKey(bv.getRight()) + "|"
             + bsKey(bv.getTop()) + "|" + bsKey(bv.getBottom()) + "|"
             + bsKey(bv.getDiagonal()) + "|" + bv.getDiagonalUp() + "|" + bv.getDiagonalDown();
    }

    /**
     * Processes xf key.
     * @param nf nf
     * @param fi fi
     * @param ll ll
     * @param bi bi
     * @param av av
     * @param pv pv
     * @return the computed result
     */
    private static String xfKey(int nf, int fi, int ll, int bi, AlignmentValue av, ProtectionValue pv) {
        return nf + "|" + fi + "|" + ll + "|" + bi + "|"
             + av.getHorizontal() + "|" + av.getVertical() + "|" + av.getWrapText() + "|"
             + av.getIndentLevel() + "|" + av.getTextRotation() + "|" + av.getShrinkToFit() + "|"
             + av.getReadingOrder() + "|" + av.getRelativeIndent() + "|"
             + pv.getIsLocked() + "|" + pv.getIsHidden();
    }

    // =========================================================================
    // Private helpers: name-to-string for OOXML output
    // =========================================================================

    /**
     * Processes fill pattern name.
     * @param p p
     * @return the computed result
     */
    private static String fillPatternName(FillPatternKind p) {
        // Handle the relevant branch before the state changes.
        if (p == null) return "none";
        switch (p) {
            case SOLID:            return "solid";
            case MEDIUM_GRAY:      return "mediumGray";
            case DARK_GRAY:        return "darkGray";
            case GRAY_125:         return "gray125";
            case GRAY_0625:        return "gray0625";
            case DARK_HORIZONTAL:  return "darkHorizontal";
            case DARK_VERTICAL:    return "darkVertical";
            case DARK_DOWN:        return "darkDown";
            case DARK_UP:          return "darkUp";
            case DARK_GRID:        return "darkGrid";
            case DARK_TRELLIS:     return "darkTrellis";
            case LIGHT_HORIZONTAL: return "lightHorizontal";
            case LIGHT_VERTICAL:   return "lightVertical";
            case LIGHT_DOWN:       return "lightDown";
            case LIGHT_UP:         return "lightUp";
            case LIGHT_GRID:       return "lightGrid";
            case LIGHT_TRELLIS:    return "lightTrellis";
            default:               return "none";
        }
    }

    /**
     * Processes border style name.
     * @param bs bs
     * @return the computed result
     */
    private static String borderStyleName(BorderStyle bs) {
        // Handle the relevant branch before the state changes.
        if (bs == null) return "none";
        switch (bs) {
            case THIN:                return "thin";
            case MEDIUM:              return "medium";
            case THICK:               return "thick";
            case DOTTED:              return "dotted";
            case DASHED:              return "dashed";
            case DOUBLE:              return "double";
            case HAIR:                return "hair";
            case MEDIUM_DASHED:       return "mediumDashed";
            case DASH_DOT:            return "dashDot";
            case MEDIUM_DASH_DOT:     return "mediumDashDot";
            case DASH_DOT_DOT:        return "dashDotDot";
            case MEDIUM_DASH_DOT_DOT: return "mediumDashDotDot";
            case SLANTED_DASH_DOT:    return "slantedDashDot";
            default:                  return "none";
        }
    }

    /**
     * Processes h align name.
     * @param h h
     * @return the computed result
     */
    private static String hAlignName(HorizontalAlignment h) {
        // Handle the relevant branch before the state changes.
        if (h == null) return "general";
        switch (h) {
            case LEFT:              return "left";
            case CENTER:            return "center";
            case RIGHT:             return "right";
            case FILL:              return "fill";
            case JUSTIFY:           return "justify";
            case CENTER_CONTINUOUS: return "centerContinuous";
            case DISTRIBUTED:       return "distributed";
            default:                return "general";
        }
    }

    /**
     * Processes v align name.
     * @param v v
     * @return the computed result
     */
    private static String vAlignName(VerticalAlignment v) {
        // Handle the relevant branch before the state changes.
        if (v == null) return "bottom";
        switch (v) {
            case TOP:         return "top";
            case CENTER:      return "center";
            case JUSTIFY:     return "justify";
            case DISTRIBUTED: return "distributed";
            default:          return "bottom";
        }
    }

    /**
     * Indicates whether has non default alignment.
     * @param av av
     * @return true when the condition is satisfied
     */
    private static boolean hasNonDefaultAlignment(AlignmentValue av) {
        return (av.getHorizontal() != null && av.getHorizontal() != HorizontalAlignment.GENERAL)
            || (av.getVertical() != null && av.getVertical() != VerticalAlignment.BOTTOM)
            || av.getWrapText() || av.getIndentLevel() != 0 || av.getTextRotation() != 0
            || av.getShrinkToFit() || av.getReadingOrder() != 0 || av.getRelativeIndent() != 0;
    }

    /**
     * Appends alignment element.
     * @param sb sb
     * @param av av
     */
    private static void appendAlignmentElement(StringBuilder sb, AlignmentValue av) {
        sb.append("<alignment");
        // Handle the relevant branch before the state changes.
        if (av.getHorizontal() != null && av.getHorizontal() != HorizontalAlignment.GENERAL)
            sb.append(" horizontal=\"").append(hAlignName(av.getHorizontal())).append("\"");
        if (av.getVertical() != null && av.getVertical() != VerticalAlignment.BOTTOM)
            sb.append(" vertical=\"").append(vAlignName(av.getVertical())).append("\"");
        if (av.getWrapText())        sb.append(" wrapText=\"1\"");
        if (av.getIndentLevel() != 0) sb.append(" indent=\"").append(av.getIndentLevel()).append("\"");
        if (av.getTextRotation() != 0) sb.append(" textRotation=\"").append(av.getTextRotation()).append("\"");
        if (av.getShrinkToFit())     sb.append(" shrinkToFit=\"1\"");
        if (av.getReadingOrder() != 0) sb.append(" readingOrder=\"").append(av.getReadingOrder()).append("\"");
        if (av.getRelativeIndent() != 0) sb.append(" relativeIndent=\"").append(av.getRelativeIndent()).append("\"");
        sb.append("/>");
    }

    /**
     * Appends protection element.
     * @param sb sb
     * @param pv pv
     */
    private static void appendProtectionElement(StringBuilder sb, ProtectionValue pv) {
        sb.append("<protection");
        // Handle the relevant branch before the state changes.
        if (!pv.getIsLocked()) sb.append(" locked=\"0\"");
        if (pv.getIsHidden())  sb.append(" hidden=\"1\"");
        sb.append("/>");
    }

    /**
     * Appends border side.
     * @param sb sb
     * @param tag tag
     * @param bsv bsv
     */
    private static void appendBorderSide(StringBuilder sb, String tag, BorderSideValue bsv) {
        // Handle the relevant branch before the state changes.
        if (bsv == null || bsv.getStyle() == null) {
            sb.append("<").append(tag).append("/>");
            return;
        }
        sb.append("<").append(tag).append(" style=\"").append(borderStyleName(bsv.getStyle())).append("\">");
        if (bsv.getColor() != null)
            sb.append("<color rgb=\"").append(XlsxWorkbookSerializerCommon.colorArgb(bsv.getColor())).append("\"/>");
        sb.append("</").append(tag).append(">");
    }

    // =========================================================================
    // Private helpers: XML parsing
    // =========================================================================

    /**
     * Finds child el.
     * @param parent parent
     * @param localName name to use
     * @return the requested result
     */
    static Element findChildEl(Element parent, String localName) {
        NodeList kids = parent.getChildNodes();
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < kids.getLength(); i++) {
            if (kids.item(i) instanceof Element && localName.equals(kids.item(i).getLocalName()))
                return (Element) kids.item(i);
        }
        return null;
    }

    /**
     * Parses argb color.
     * @param rgb rgb
     * @return the computed result
     */
    /** Serializes a font element to a complete &lt;font&gt;…&lt;/font&gt; XML string, preserving all child elements verbatim. */
    private static String fontElementToXml(Element fontEl) {
        StringBuilder sb = new StringBuilder("<font>");
        org.w3c.dom.NodeList children = fontEl.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            org.w3c.dom.Node n = children.item(i);
            if (!(n instanceof Element)) continue;
            Element child = (Element) n;
            String localName = child.getLocalName();
            sb.append("<").append(localName);
            org.w3c.dom.NamedNodeMap attrs = child.getAttributes();
            for (int j = 0; j < attrs.getLength(); j++) {
                org.w3c.dom.Attr a = (org.w3c.dom.Attr) attrs.item(j);
                sb.append(" ").append(a.getName()).append("=\"")
                  .append(a.getValue().replace("&", "&amp;").replace("\"", "&quot;"))
                  .append("\"");
            }
            sb.append("/>");
        }
        sb.append("</font>");
        return sb.toString();
    }

    /** Serializes a patternFill element to a complete <fill>…</fill> XML string, preserving all color attributes verbatim. */
    private static String fillElementToXml(Element pf) {
        StringBuilder sb = new StringBuilder("<fill><patternFill");
        String pt = pf.getAttribute("patternType");
        if (!pt.isEmpty()) sb.append(" patternType=\"").append(pt).append("\"");
        sb.append(">");
        appendColorEl(sb, "fgColor", findChildEl(pf, "fgColor"));
        appendColorEl(sb, "bgColor", findChildEl(pf, "bgColor"));
        sb.append("</patternFill></fill>");
        return sb.toString();
    }

    private static void appendColorEl(StringBuilder sb, String tag, Element el) {
        if (el == null) return;
        sb.append("<").append(tag);
        org.w3c.dom.NamedNodeMap attrs = el.getAttributes();
        for (int i = 0; i < attrs.getLength(); i++) {
            org.w3c.dom.Attr a = (org.w3c.dom.Attr) attrs.item(i);
            sb.append(" ").append(a.getName()).append("=\"")
              .append(a.getValue().replace("&", "&amp;").replace("\"", "&quot;"))
              .append("\"");
        }
        sb.append("/>");
    }

    private static ColorValue parseArgbColor(String rgb) {
        // Handle the relevant branch before the state changes.
        if (rgb == null || rgb.length() < 8) return null;
        try {
            int a = Integer.parseInt(rgb.substring(0, 2), 16);
            int r = Integer.parseInt(rgb.substring(2, 4), 16);
            int g = Integer.parseInt(rgb.substring(4, 6), 16);
            int b = Integer.parseInt(rgb.substring(6, 8), 16);
            return new ColorValue((byte) a, (byte) r, (byte) g, (byte) b);
        } catch (NumberFormatException e) { return null; }
    }

    /**
     * Parses fill pattern type.
     * @param s s
     * @return the computed result
     */
    private static FillPatternKind parseFillPatternType(String s) {
        // Handle the relevant branch before the state changes.
        if (s == null) return FillPatternKind.NONE;
        switch (s.toLowerCase()) {
            case "solid":            return FillPatternKind.SOLID;
            case "mediumgray":       return FillPatternKind.MEDIUM_GRAY;
            case "darkgray":         return FillPatternKind.DARK_GRAY;
            case "gray125":          return FillPatternKind.GRAY_125;
            case "gray0625":         return FillPatternKind.GRAY_0625;
            case "darkhorizontal":   return FillPatternKind.DARK_HORIZONTAL;
            case "darkvertical":     return FillPatternKind.DARK_VERTICAL;
            case "darkdown":         return FillPatternKind.DARK_DOWN;
            case "darkup":           return FillPatternKind.DARK_UP;
            case "darkgrid":         return FillPatternKind.DARK_GRID;
            case "darktrellis":      return FillPatternKind.DARK_TRELLIS;
            case "lighthorizontal":  return FillPatternKind.LIGHT_HORIZONTAL;
            case "lightvertical":    return FillPatternKind.LIGHT_VERTICAL;
            case "lightdown":        return FillPatternKind.LIGHT_DOWN;
            case "lightup":          return FillPatternKind.LIGHT_UP;
            case "lightgrid":        return FillPatternKind.LIGHT_GRID;
            case "lighttrellis":     return FillPatternKind.LIGHT_TRELLIS;
            default:                 return FillPatternKind.NONE;
        }
    }

    /**
     * Reads border side el.
     * @param el el
     * @return the computed result
     */
    private static BorderSideValue readBorderSideEl(Element el) {
        BorderSideValue bsv = new BorderSideValue();
        // Handle the relevant branch before the state changes.
        if (el == null) return bsv;
        String style = el.getAttribute("style");
        if (!style.isEmpty()) bsv.setStyle(parseBorderStyleStr(style));
        Element colorEl = findChildEl(el, "color");
        if (colorEl != null) bsv.setColor(parseArgbColor(colorEl.getAttribute("rgb")));
        return bsv;
    }

    /**
     * Parses border style str.
     * @param s s
     * @return the computed result
     */
    private static BorderStyle parseBorderStyleStr(String s) {
        // Handle the relevant branch before the state changes.
        if (s == null) return null;
        switch (s.toLowerCase()) {
            case "thin":               return BorderStyle.THIN;
            case "medium":             return BorderStyle.MEDIUM;
            case "thick":              return BorderStyle.THICK;
            case "dotted":             return BorderStyle.DOTTED;
            case "dashed":             return BorderStyle.DASHED;
            case "double":             return BorderStyle.DOUBLE;
            case "hair":               return BorderStyle.HAIR;
            case "mediumdashed":       return BorderStyle.MEDIUM_DASHED;
            case "dashdot":            return BorderStyle.DASH_DOT;
            case "mediumdashdot":      return BorderStyle.MEDIUM_DASH_DOT;
            case "dashdotdot":         return BorderStyle.DASH_DOT_DOT;
            case "mediumdashdotdot":   return BorderStyle.MEDIUM_DASH_DOT_DOT;
            case "slanteddashdot":     return BorderStyle.SLANTED_DASH_DOT;
            default:                   return null;
        }
    }

    /**
     * Parses h align str.
     * @param s s
     * @return the computed result
     */
    private static HorizontalAlignment parseHAlignStr(String s) {
        // Handle the relevant branch before the state changes.
        if (s == null || s.isEmpty()) return HorizontalAlignment.GENERAL;
        switch (s.toLowerCase()) {
            case "left":              return HorizontalAlignment.LEFT;
            case "center":            return HorizontalAlignment.CENTER;
            case "right":             return HorizontalAlignment.RIGHT;
            case "fill":              return HorizontalAlignment.FILL;
            case "justify":           return HorizontalAlignment.JUSTIFY;
            case "centercontinuous":  return HorizontalAlignment.CENTER_CONTINUOUS;
            case "distributed":       return HorizontalAlignment.DISTRIBUTED;
            default:                  return HorizontalAlignment.GENERAL;
        }
    }

    /**
     * Parses v align str.
     * @param s s
     * @return the computed result
     */
    private static VerticalAlignment parseVAlignStr(String s) {
        // Handle the relevant branch before the state changes.
        if (s == null || s.isEmpty()) return VerticalAlignment.BOTTOM;
        switch (s.toLowerCase()) {
            case "top":         return VerticalAlignment.TOP;
            case "center":      return VerticalAlignment.CENTER;
            case "justify":     return VerticalAlignment.JUSTIFY;
            case "distributed": return VerticalAlignment.DISTRIBUTED;
            default:            return VerticalAlignment.BOTTOM;
        }
    }
}
