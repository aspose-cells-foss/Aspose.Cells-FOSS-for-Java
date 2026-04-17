package com.aspose.cells_foss;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import com.aspose.cells_foss.core.ColorValue;
import com.aspose.cells_foss.core.ProtectionValue;
import com.aspose.cells_foss.core.StyleValue;
import com.aspose.cells_foss.core.HorizontalAlignment;
import com.aspose.cells_foss.core.FontValue;
import com.aspose.cells_foss.core.VerticalAlignment;
import com.aspose.cells_foss.core.AlignmentValue;
import com.aspose.cells_foss.core.NumberFormatValue;
import com.aspose.cells_foss.core.BordersValue;
import com.aspose.cells_foss.core.BorderSideValue;
import com.aspose.cells_foss.core.FillPatternKind;

/**
 * Provides methods for reading and writing XLSX workbook styles XML.
 */
public final class XlsxWorkbookStylesXml {
    
    private static final String MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    
    /**
     * Initializes a new XlsxWorkbookStylesXml instance.
     */
    private XlsxWorkbookStylesXml() {}
    
    /**
     * Inner class representing fill value.
     */
    public static final class FillValue {
        private FillPatternKind pattern;
        private ColorValue foregroundColor;
        private ColorValue backgroundColor;
        
        /**
         * Initializes a new FillValue instance.
         */
        public FillValue() {
            this.pattern = FillPatternKind.NONE;
            this.foregroundColor = new ColorValue((byte)0, (byte)0, (byte)0, (byte)0);
            this.backgroundColor = new ColorValue((byte)0, (byte)0, (byte)0, (byte)0);
        }
        
        /**
         * Returns the pattern.
         * @return the requested result
         */
        public FillPatternKind getPattern() {
            return pattern;
        }
        
        /**
         * Sets the pattern.
         * @param pattern pattern
         */
        public void setPattern(FillPatternKind pattern) {
            this.pattern = pattern;
        }
        
        /**
         * Returns the foreground color.
         * @return the requested result
         */
        public ColorValue getForegroundColor() {
            return foregroundColor;
        }
        
        /**
         * Sets the foreground color.
         * @param foregroundColor foreground color
         */
        public void setForegroundColor(ColorValue foregroundColor) {
            this.foregroundColor = foregroundColor;
        }
        
        /**
         * Returns the background color.
         * @return the requested result
         */
        public ColorValue getBackgroundColor() {
            return backgroundColor;
        }
        
        /**
         * Sets the background color.
         * @param backgroundColor background color
         */
        public void setBackgroundColor(ColorValue backgroundColor) {
            this.backgroundColor = backgroundColor;
        }
    }
    
    /**
     * Inner class representing cell format value.
     */
    public static final class CellFormatValue {
        private int numFmtId;
        private int fontId;
        private int fillId;
        private int borderId;
        private AlignmentValue alignment;
        private ProtectionValue protection;
        
        /**
         * Initializes a new CellFormatValue instance.
         */
        public CellFormatValue() {
            this.alignment = new AlignmentValue();
            this.protection = new ProtectionValue();
        }
        
        /**
         * Returns the num fmt id.
         * @return the requested result
         */
        public int getNumFmtId() {
            return numFmtId;
        }
        
        /**
         * Sets the num fmt id.
         * @param numFmtId num fmt id
         */
        public void setNumFmtId(int numFmtId) {
            this.numFmtId = numFmtId;
        }
        
        /**
         * Returns the font id.
         * @return the requested result
         */
        public int getFontId() {
            return fontId;
        }
        
        /**
         * Sets the font id.
         * @param fontId font id
         */
        public void setFontId(int fontId) {
            this.fontId = fontId;
        }
        
        /**
         * Returns the fill id.
         * @return the requested result
         */
        public int getFillId() {
            return fillId;
        }
        
        /**
         * Sets the fill id.
         * @param fillId fill id
         */
        public void setFillId(int fillId) {
            this.fillId = fillId;
        }
        
        /**
         * Returns the border id.
         * @return the requested result
         */
        public int getBorderId() {
            return borderId;
        }
        
        /**
         * Sets the border id.
         * @param borderId border id
         */
        public void setBorderId(int borderId) {
            this.borderId = borderId;
        }
        
        /**
         * Returns the alignment.
         * @return the requested result
         */
        public AlignmentValue getAlignment() {
            return alignment;
        }
        
        /**
         * Sets the alignment.
         * @param alignment alignment
         */
        public void setAlignment(AlignmentValue alignment) {
            this.alignment = alignment;
        }
        
        /**
         * Returns the protection.
         * @return the requested result
         */
        public ProtectionValue getProtection() {
            return protection;
        }
        
        /**
         * Sets the protection.
         * @param protection protection
         */
        public void setProtection(ProtectionValue protection) {
            this.protection = protection;
        }
    }
    
    /**
     * Builds a stylesheet document from the provided style values.
     */
    public static Document buildStylesheetDocument(
            List<FontValue> fonts,
            List<FillValue> fills,
            List<BordersValue> borders,
            CellFormatValue normalCellFormat,
            List<CellFormatValue> cellFormats,
            List<java.util.Map.Entry<Integer, String>> customNumberFormats,
            List<StyleValue> differentialFormats) {
        
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element stylesheet = doc.createElementNS(MAIN_NS, "styleSheet");
            doc.appendChild(stylesheet);
            
            if (customNumberFormats != null && !customNumberFormats.isEmpty()) {
                Element numFmts = doc.createElementNS(MAIN_NS, "numFmts");
                numFmts.setAttribute("count", String.valueOf(customNumberFormats.size()));
                
                for (java.util.Map.Entry<Integer, String> pair : customNumberFormats) {
                    Element numFmt = doc.createElementNS(MAIN_NS, "numFmt");
                    numFmt.setAttribute("numFmtId", String.valueOf(pair.getKey()));
                    numFmt.setAttribute("formatCode", pair.getValue());
                    numFmts.appendChild(numFmt);
                }
                
                stylesheet.appendChild(numFmts);
            }
            
            // Fonts
            Element fontsElement = doc.createElementNS(MAIN_NS, "fonts");
            fontsElement.setAttribute("count", String.valueOf(fonts.size()));
            List<Element> fontElements = buildFontElements(fonts);
            for (Element fontElement : fontElements) {
                fontsElement.appendChild(fontElement);
            }
            stylesheet.appendChild(fontsElement);
            
            // Fills
            Element fillsElement = doc.createElementNS(MAIN_NS, "fills");
            fillsElement.setAttribute("count", String.valueOf(fills.size()));
            List<Element> fillElements = buildFillElements(fills);
            for (Element fillElement : fillElements) {
                fillsElement.appendChild(fillElement);
            }
            stylesheet.appendChild(fillsElement);
            
            // Borders
            Element bordersElement = doc.createElementNS(MAIN_NS, "borders");
            bordersElement.setAttribute("count", String.valueOf(borders.size()));
            List<Element> borderElements = buildBorderElements(borders);
            for (Element borderElement : borderElements) {
                bordersElement.appendChild(borderElement);
            }
            stylesheet.appendChild(bordersElement);
            
            // CellStyleXfs
            Element cellStyleXfs = doc.createElementNS(MAIN_NS, "cellStyleXfs");
            cellStyleXfs.setAttribute("count", "1");
            Element normalCellFormatElement = buildCellStyleFormatElement(normalCellFormat);
            cellStyleXfs.appendChild(normalCellFormatElement);
            stylesheet.appendChild(cellStyleXfs);
            
            // CellXfs
            Element cellXfs = doc.createElementNS(MAIN_NS, "cellXfs");
            cellXfs.setAttribute("count", String.valueOf(cellFormats.size()));
            List<Element> cellFormatElements = buildCellFormatElements(cellFormats);
            for (Element cellFormatElement : cellFormatElements) {
                cellXfs.appendChild(cellFormatElement);
            }
            stylesheet.appendChild(cellXfs);
            
            // CellStyles
            Element cellStyles = doc.createElementNS(MAIN_NS, "cellStyles");
            cellStyles.setAttribute("count", "1");
            Element normalStyle = doc.createElementNS(MAIN_NS, "cellStyle");
            normalStyle.setAttribute("name", "Normal");
            normalStyle.setAttribute("xfId", "0");
            normalStyle.setAttribute("builtinId", "0");
            cellStyles.appendChild(normalStyle);
            stylesheet.appendChild(cellStyles);
            
            // Differential formats
            if (differentialFormats != null && !differentialFormats.isEmpty()) {
                Element dxfs = doc.createElementNS(MAIN_NS, "dxfs");
                dxfs.setAttribute("count", String.valueOf(differentialFormats.size()));
                List<Element> differentialFormatElements = buildDifferentialFormatElements(differentialFormats);
                for (Element dfElement : differentialFormatElements) {
                    dxfs.appendChild(dfElement);
                }
                stylesheet.appendChild(dxfs);
            }
            
            return doc;
        } catch (Exception e) {
            throw new RuntimeException("Failed to build stylesheet document", e);
        }
    }
    
    /**
     * Reads font values from the stylesheet XML.
     */
    public static List<FontValue> readFontValues(Element root) {
        List<FontValue> fonts = new ArrayList<>();
        NodeList fontNodes = root.getElementsByTagNameNS(MAIN_NS, "font");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < fontNodes.getLength(); i++) {
            Element fontElement = (Element) fontNodes.item(i);
            fonts.add(readFontValue(fontElement));
        }
        return fonts;
    }
    
    /**
     * Reads fill values from the stylesheet XML.
     */
    public static List<FillValue> readFillValues(Element root) {
        List<FillValue> fills = new ArrayList<>();
        NodeList fillNodes = root.getElementsByTagNameNS(MAIN_NS, "fill");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < fillNodes.getLength(); i++) {
            Element fillElement = (Element) fillNodes.item(i);
            fills.add(readFillValue(fillElement));
        }
        return fills;
    }
    
    /**
     * Reads border values from the stylesheet XML.
     */
    public static List<BordersValue> readBordersValues(Element root) {
        List<BordersValue> borders = new ArrayList<>();
        NodeList borderNodes = root.getElementsByTagNameNS(MAIN_NS, "border");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < borderNodes.getLength(); i++) {
            Element borderElement = (Element) borderNodes.item(i);
            borders.add(readBordersValue(borderElement));
        }
        return borders;
    }
    
    /**
     * Reads differential style values from the stylesheet XML.
     */
    public static List<StyleValue> readDifferentialStyleValues(Element root) {
        List<StyleValue> styles = new ArrayList<>();
        NodeList dxfNodes = root.getElementsByTagNameNS(MAIN_NS, "dxf");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < dxfNodes.getLength(); i++) {
            Element dxfElement = (Element) dxfNodes.item(i);
            styles.add(readDifferentialStyleValue(dxfElement));
        }
        return styles;
    }
    
    /**
     * Builds font elements for the stylesheet.
     */
    private static List<Element> buildFontElements(List<FontValue> fonts) {
        List<Element> elements = new ArrayList<>(fonts.size());
        // Walk the current collection so every entry is processed consistently.
        for (FontValue font : fonts) {
            elements.add(buildFontElement(font));
        }
        return elements;
    }
    
    /**
     * Builds fill elements for the stylesheet.
     */
    private static List<Element> buildFillElements(List<FillValue> fills) {
        List<Element> elements = new ArrayList<>(fills.size());
        // Walk the current collection so every entry is processed consistently.
        for (FillValue fill : fills) {
            elements.add(buildFillElement(fill));
        }
        return elements;
    }
    
    /**
     * Builds border elements for the stylesheet.
     */
    private static List<Element> buildBorderElements(List<BordersValue> borders) {
        List<Element> elements = new ArrayList<>(borders.size());
        // Walk the current collection so every entry is processed consistently.
        for (BordersValue border : borders) {
            elements.add(buildBorderElement(border));
        }
        return elements;
    }
    
    /**
     * Builds cell format elements for the stylesheet.
     */
    private static List<Element> buildCellFormatElements(List<CellFormatValue> cellFormats) {
        List<Element> elements = new ArrayList<>(cellFormats.size());
        // Walk the current collection so every entry is processed consistently.
        for (CellFormatValue cellFormat : cellFormats) {
            elements.add(buildCellFormatElement(cellFormat));
        }
        return elements;
    }
    
    /**
     * Builds differential format elements for the stylesheet.
     */
    private static List<Element> buildDifferentialFormatElements(List<StyleValue> differentialFormats) {
        List<Element> elements = new ArrayList<>(differentialFormats.size());
        // Walk the current collection so every entry is processed consistently.
        for (StyleValue style : differentialFormats) {
            elements.add(buildDifferentialFormatElement(style));
        }
        return elements;
    }
    
    /**
     * Parses horizontal alignment from a string value.
     */
    public static HorizontalAlignment parseHorizontalAlignment(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            return HorizontalAlignment.GENERAL;
        }
        String lowerValue = value.toLowerCase(Locale.ROOT);
        switch (lowerValue) {
            case "left":
                return HorizontalAlignment.LEFT;
            case "center":
                return HorizontalAlignment.CENTER;
            case "right":
                return HorizontalAlignment.RIGHT;
            case "fill":
                return HorizontalAlignment.FILL;
            case "justify":
                return HorizontalAlignment.JUSTIFY;
            case "centercontinuous":
                return HorizontalAlignment.CENTER_CONTINUOUS;
            case "distributed":
                return HorizontalAlignment.DISTRIBUTED;
            default:
                return HorizontalAlignment.GENERAL;
        }
    }
    
    /**
     * Parses vertical alignment from a string value.
     */
    public static VerticalAlignment parseVerticalAlignment(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            return VerticalAlignment.BOTTOM;
        }
        String lowerValue = value.toLowerCase(Locale.ROOT);
        switch (lowerValue) {
            case "center":
                return VerticalAlignment.CENTER;
            case "top":
                return VerticalAlignment.TOP;
            case "justify":
                return VerticalAlignment.JUSTIFY;
            case "distributed":
                return VerticalAlignment.DISTRIBUTED;
            default:
                return VerticalAlignment.BOTTOM;
        }
    }
    
    /**
     * Parses an optional boolean attribute.
     */
    private static Boolean parseOptionalBoolAttribute(Element attribute) {
        // Handle the relevant branch before the state changes.
        if (attribute == null) {
            return null;
        }
        String value = attribute.getAttribute("val");
        if (value == null) {
            return null;
        }
        return Boolean.parseBoolean(value);
    }
    
    /**
     * Builds a cell style format element.
     */
    private static Element buildCellStyleFormatElement(CellFormatValue cellFormat) {
        return buildCellFormatElement(cellFormat, false);
    }
    
    /**
     * Reads a differential style value from an XML element.
     */
    private static StyleValue readDifferentialStyleValue(Element dxf) {
        StyleValue style = StyleValue.getDefault().clone();
        
        // Font
        Element fontElement = findElement(dxf, "font");
        if (fontElement != null) {
            style.setFont(readFontValue(fontElement));
        }
        
        // Fill
        Element fillElement = findElement(dxf, "fill");
        if (fillElement != null) {
            FillValue fillValue = readFillValue(fillElement);
            style.setPattern(fillValue.getPattern());
            style.setForegroundColor(fillValue.getForegroundColor());
            style.setBackgroundColor(fillValue.getBackgroundColor());
        }
        
        // Border
        Element borderElement = findElement(dxf, "border");
        if (borderElement != null) {
            style.setBorders(readBordersValue(borderElement));
        }
        
        // Number format
        Element numFmtElement = findElement(dxf, "numFmt");
        if (numFmtElement != null) {
            NumberFormatValue numFmt = new NumberFormatValue();
            String numFmtIdAttr = numFmtElement.getAttribute("numFmtId");
            if (!numFmtIdAttr.isEmpty()) {
                numFmt.setNumber(Integer.parseInt(numFmtIdAttr));
            } else {
                numFmt.setNumber(0);
            }
            String formatCode = numFmtElement.getAttribute("formatCode");
            if (!formatCode.isEmpty()) {
                numFmt.setCustom(formatCode);
            }
            style.setNumberFormat(numFmt);
        }
        
        // Alignment
        Element alignmentElement = findElement(dxf, "alignment");
        if (alignmentElement != null) {
            AlignmentValue alignment = new AlignmentValue();
            alignment.setHorizontal(parseHorizontalAlignment(alignmentElement.getAttribute("horizontal")));
            alignment.setVertical(parseVerticalAlignment(alignmentElement.getAttribute("vertical")));
            alignment.setWrapText(parseBoolAttribute(alignmentElement.getAttribute("wrapText")));
            alignment.setIndentLevel(StyleValueSanitizer.normalizeIndentLevel(parseIntAttribute(alignmentElement.getAttribute("indent"))));
            alignment.setTextRotation(StyleValueSanitizer.normalizeTextRotation(parseIntAttribute(alignmentElement.getAttribute("textRotation"))));
            alignment.setShrinkToFit(parseBoolAttribute(alignmentElement.getAttribute("shrinkToFit")));
            alignment.setReadingOrder(StyleValueSanitizer.normalizeReadingOrder(parseIntAttribute(alignmentElement.getAttribute("readingOrder"))));
            String relativeIndent = alignmentElement.getAttribute("relativeIndent");
            if (!relativeIndent.isEmpty()) {
                alignment.setRelativeIndent(Integer.parseInt(relativeIndent));
            } else {
                alignment.setRelativeIndent(0);
            }
            style.setAlignment(alignment);
        }
        
        // Protection
        Element protectionElement = findElement(dxf, "protection");
        if (protectionElement != null) {
            ProtectionValue protection = new ProtectionValue();
            Boolean locked = parseOptionalBoolAttribute(protectionElement);
            protection.setIsLocked(locked == null || locked);
            protection.setIsHidden(parseBoolAttribute(protectionElement.getAttribute("hidden")));
            style.setProtection(protection);
        }
        
        return style;
    }
    
    /**
     * Reads a font value from an XML element.
     */
    private static FontValue readFontValue(Element font) {
        FontValue fontValue = new FontValue();
        
        Element nameElement = findElement(font, "name");
        // Handle the relevant branch before the state changes.
        if (nameElement != null) {
            String name = nameElement.getAttribute("val");
            if (!name.isEmpty()) {
                fontValue.setName(name);
            } else {
                fontValue.setName("Calibri");
            }
        } else {
            fontValue.setName("Calibri");
        }
        
        Element szElement = findElement(font, "sz");
        if (szElement != null) {
            String sizeStr = szElement.getAttribute("val");
            if (!sizeStr.isEmpty()) {
                try {
                    fontValue.setSize(Double.parseDouble(sizeStr));
                } catch (NumberFormatException e) {
                    fontValue.setSize(11.0);
                }
            } else {
                fontValue.setSize(11.0);
            }
        } else {
            fontValue.setSize(11.0);
        }
        
        fontValue.setBold(findElement(font, "b") != null);
        fontValue.setItalic(findElement(font, "i") != null);
        fontValue.setUnderline(findElement(font, "u") != null);
        fontValue.setStrikeThrough(findElement(font, "strike") != null);
        
        Element colorElement = findElement(font, "color");
        fontValue.setColor(readColorValue(colorElement));
        
        return fontValue;
    }
    
    /**
     * Reads a fill value from an XML element.
     */
    private static FillValue readFillValue(Element fill) {
        FillValue fillValue = new FillValue();
        
        Element patternFillElement = findElement(fill, "patternFill");
        // Handle the relevant branch before the state changes.
        if (patternFillElement == null) {
            return fillValue;
        }
        
        String patternType = patternFillElement.getAttribute("patternType");
        if (patternType.isEmpty()) {
            patternType = "none";
        }
        patternType = patternType.toLowerCase(Locale.ROOT);
        
        FillPatternKind pattern;
        switch (patternType) {
            case "solid":
                pattern = FillPatternKind.SOLID;
                break;
            case "mediumgray":
                pattern = FillPatternKind.MEDIUM_GRAY;
                break;
            case "darkgray":
                pattern = FillPatternKind.DARK_GRAY;
                break;
            case "gray125":
                pattern = FillPatternKind.GRAY_125;
                break;
            case "gray0625":
                pattern = FillPatternKind.GRAY_0625;
                break;
            case "darkhorizontal":
                pattern = FillPatternKind.DARK_HORIZONTAL;
                break;
            case "darkvertical":
                pattern = FillPatternKind.DARK_VERTICAL;
                break;
            case "darkdown":
                pattern = FillPatternKind.DARK_DOWN;
                break;
            case "darkup":
                pattern = FillPatternKind.DARK_UP;
                break;
            case "darkgrid":
                pattern = FillPatternKind.DARK_GRID;
                break;
            case "darktrellis":
                pattern = FillPatternKind.DARK_TRELLIS;
                break;
            case "lighthorizontal":
                pattern = FillPatternKind.LIGHT_HORIZONTAL;
                break;
            case "lightvertical":
                pattern = FillPatternKind.LIGHT_VERTICAL;
                break;
            case "lightdown":
                pattern = FillPatternKind.LIGHT_DOWN;
                break;
            case "lightup":
                pattern = FillPatternKind.LIGHT_UP;
                break;
            case "lightgrid":
                pattern = FillPatternKind.LIGHT_GRID;
                break;
            case "lighttrellis":
                pattern = FillPatternKind.LIGHT_TRELLIS;
                break;
            default:
                pattern = FillPatternKind.NONE;
                break;
        }
        
        fillValue.setPattern(pattern);
        
        Element fgColorElement = findElement(patternFillElement, "fgColor");
        fillValue.setForegroundColor(readColorValue(fgColorElement));
        
        Element bgColorElement = findElement(patternFillElement, "bgColor");
        fillValue.setBackgroundColor(readColorValue(bgColorElement));
        
        return fillValue;
    }
    
    /**
     * Reads a borders value from an XML element.
     */
    private static BordersValue readBordersValue(Element border) {
        BordersValue bordersValue = new BordersValue();
        
        Element leftElement = findElement(border, "left");
        bordersValue.setLeft(readBorderSideValue(leftElement));
        
        Element rightElement = findElement(border, "right");
        bordersValue.setRight(readBorderSideValue(rightElement));
        
        Element topElement = findElement(border, "top");
        bordersValue.setTop(readBorderSideValue(topElement));
        
        Element bottomElement = findElement(border, "bottom");
        bordersValue.setBottom(readBorderSideValue(bottomElement));
        
        Element diagonalElement = findElement(border, "diagonal");
        bordersValue.setDiagonal(readBorderSideValue(diagonalElement));
        
        bordersValue.setDiagonalUp(parseBoolAttribute(border.getAttribute("diagonalUp")));
        bordersValue.setDiagonalDown(parseBoolAttribute(border.getAttribute("diagonalDown")));
        
        return bordersValue;
    }
    
    /**
     * Builds a font element.
     */
    private static Element buildFontElement(FontValue font) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element element = doc.createElementNS(MAIN_NS, "font");
            
            if (font.getBold()) {
                element.appendChild(doc.createElementNS(MAIN_NS, "b"));
            }
            if (font.getItalic()) {
                element.appendChild(doc.createElementNS(MAIN_NS, "i"));
            }
            if (font.getStrikeThrough()) {
                element.appendChild(doc.createElementNS(MAIN_NS, "strike"));
            }
            if (font.getUnderline()) {
                element.appendChild(doc.createElementNS(MAIN_NS, "u"));
            }
            
            Element szElement = doc.createElementNS(MAIN_NS, "sz");
            DecimalFormat df = new DecimalFormat("0.####");
            df.setDecimalSeparatorAlwaysShown(false);
            szElement.setAttribute("val", df.format(font.getSize()));
            element.appendChild(szElement);
            
            Element colorElement = buildColorElement("color", font.getColor());
            if (colorElement != null) {
                element.appendChild(colorElement);
            }
            
            Element nameElement = doc.createElementNS(MAIN_NS, "name");
            nameElement.setAttribute("val", font.getName());
            element.appendChild(nameElement);
            
            return element;
        } catch (Exception e) {
            throw new RuntimeException("Failed to build font element", e);
        }
    }
    
    /**
     * Builds a fill element.
     */
    private static Element buildFillElement(FillValue fill) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element patternFill = doc.createElementNS(MAIN_NS, "patternFill");
            
            switch (fill.getPattern()) {
                case SOLID:
                    patternFill.setAttribute("patternType", "solid");
                    break;
                case MEDIUM_GRAY:
                    patternFill.setAttribute("patternType", "mediumGray");
                    break;
                case DARK_GRAY:
                    patternFill.setAttribute("patternType", "darkGray");
                    break;
                case GRAY_125:
                    patternFill.setAttribute("patternType", "gray125");
                    break;
                case GRAY_0625:
                    patternFill.setAttribute("patternType", "gray0625");
                    break;
                case DARK_HORIZONTAL:
                    patternFill.setAttribute("patternType", "darkHorizontal");
                    break;
                case DARK_VERTICAL:
                    patternFill.setAttribute("patternType", "darkVertical");
                    break;
                case DARK_DOWN:
                    patternFill.setAttribute("patternType", "darkDown");
                    break;
                case DARK_UP:
                    patternFill.setAttribute("patternType", "darkUp");
                    break;
                case DARK_GRID:
                    patternFill.setAttribute("patternType", "darkGrid");
                    break;
                case DARK_TRELLIS:
                    patternFill.setAttribute("patternType", "darkTrellis");
                    break;
                case LIGHT_HORIZONTAL:
                    patternFill.setAttribute("patternType", "lightHorizontal");
                    break;
                case LIGHT_VERTICAL:
                    patternFill.setAttribute("patternType", "lightVertical");
                    break;
                case LIGHT_DOWN:
                    patternFill.setAttribute("patternType", "lightDown");
                    break;
                case LIGHT_UP:
                    patternFill.setAttribute("patternType", "lightUp");
                    break;
                case LIGHT_GRID:
                    patternFill.setAttribute("patternType", "lightGrid");
                    break;
                case LIGHT_TRELLIS:
                    patternFill.setAttribute("patternType", "lightTrellis");
                    break;
                default:
                    patternFill.setAttribute("patternType", "none");
                    break;
            }
            
            Element foregroundColor = buildColorElement("fgColor", fill.getForegroundColor());
            if (foregroundColor != null) {
                patternFill.appendChild(foregroundColor);
            }
            
            Element backgroundColor = buildColorElement("bgColor", fill.getBackgroundColor());
            if (backgroundColor != null) {
                patternFill.appendChild(backgroundColor);
            }
            
            Element fillElement = doc.createElementNS(MAIN_NS, "fill");
            fillElement.appendChild(patternFill);
            return fillElement;
        } catch (Exception e) {
            throw new RuntimeException("Failed to build fill element", e);
        }
    }
    
    /**
     * Builds a border element.
     */
    private static Element buildBorderElement(BordersValue borders) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element element = doc.createElementNS(MAIN_NS, "border");
            
            element.appendChild(buildBorderSideElement("left", borders.getLeft()));
            element.appendChild(buildBorderSideElement("right", borders.getRight()));
            element.appendChild(buildBorderSideElement("top", borders.getTop()));
            element.appendChild(buildBorderSideElement("bottom", borders.getBottom()));
            element.appendChild(buildBorderSideElement("diagonal", borders.getDiagonal()));
            
            if (borders.getDiagonalUp()) {
                element.setAttribute("diagonalUp", "1");
            }
            if (borders.getDiagonalDown()) {
                element.setAttribute("diagonalDown", "1");
            }
            
            return element;
        } catch (Exception e) {
            throw new RuntimeException("Failed to build border element", e);
        }
    }
    
    /**
     * Builds a cell format element.
     */
    private static Element buildCellFormatElement(CellFormatValue cellFormat) {
        return buildCellFormatElement(cellFormat, true);
    }
    
    /**
     * Builds a differential format element.
     */
    private static Element buildDifferentialFormatElement(StyleValue style) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element element = doc.createElementNS(MAIN_NS, "dxf");
            
            if (!fontEquals(style.getFont(), StyleValue.getDefault().getFont())) {
                element.appendChild(buildFontElement(style.getFont()));
            }
            
            if (style.getPattern() != FillPatternKind.NONE || 
                !isEmptyColor(style.getForegroundColor()) || 
                !isEmptyColor(style.getBackgroundColor())) {
                
                FillValue fillValue = new FillValue();
                fillValue.setPattern(style.getPattern());
                fillValue.setForegroundColor(style.getForegroundColor());
                fillValue.setBackgroundColor(style.getBackgroundColor());
                
                element.appendChild(buildFillElement(fillValue));
            }
            
            if (!bordersEqual(style.getBorders(), StyleValue.getDefault().getBorders())) {
                element.appendChild(buildBorderElement(style.getBorders()));
            }
            
            if (style.getNumberFormat().getNumber() != 0 || 
                (style.getNumberFormat().getCustom() != null && !style.getNumberFormat().getCustom().isEmpty())) {
                
                Element numFmtElement = doc.createElementNS(MAIN_NS, "numFmt");
                int numFmtId = style.getNumberFormat().getNumber() >= 0 ? 
                              style.getNumberFormat().getNumber() : 0;
                numFmtElement.setAttribute("numFmtId", String.valueOf(numFmtId));
                
                String custom = style.getNumberFormat().getCustom();
                if (custom != null && !custom.isEmpty()) {
                    numFmtElement.setAttribute("formatCode", custom);
                }
                
                element.appendChild(numFmtElement);
            }
            
            Element alignmentElement = buildAlignmentElement(style.getAlignment());
            if (alignmentElement != null) {
                element.appendChild(alignmentElement);
            }
            
            Element protectionElement = buildProtectionElement(style.getProtection());
            if (protectionElement != null) {
                element.appendChild(protectionElement);
            }
            
            return element;
        } catch (Exception e) {
            throw new RuntimeException("Failed to build differential format element", e);
        }
    }
    
    /**
     * Builds a cell format element with optional xfId attribute.
     */
    private static Element buildCellFormatElement(CellFormatValue cellFormat, boolean includeXfId) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element element = doc.createElementNS(MAIN_NS, "xf");
            element.setAttribute("numFmtId", String.valueOf(cellFormat.getNumFmtId()));
            element.setAttribute("fontId", String.valueOf(cellFormat.getFontId()));
            element.setAttribute("fillId", String.valueOf(cellFormat.getFillId()));
            element.setAttribute("borderId", String.valueOf(cellFormat.getBorderId()));
            
            if (includeXfId) {
                element.setAttribute("xfId", "0");
            }
            
            if (cellFormat.getNumFmtId() != 0) {
                element.setAttribute("applyNumberFormat", "1");
            }
            if (cellFormat.getFontId() != 0) {
                element.setAttribute("applyFont", "1");
            }
            if (cellFormat.getFillId() != 0) {
                element.setAttribute("applyFill", "1");
            }
            if (cellFormat.getBorderId() != 0) {
                element.setAttribute("applyBorder", "1");
            }
            
            Element alignmentElement = buildAlignmentElement(cellFormat.getAlignment());
            if (alignmentElement != null) {
                element.setAttribute("applyAlignment", "1");
                element.appendChild(alignmentElement);
            }
            
            Element protectionElement = buildProtectionElement(cellFormat.getProtection());
            if (protectionElement != null) {
                element.setAttribute("applyProtection", "1");
                element.appendChild(protectionElement);
            }
            
            return element;
        } catch (Exception e) {
            throw new RuntimeException("Failed to build cell format element", e);
        }
    }
    
    /**
     * Reads a border side value from an XML element.
     */
    private static BorderSideValue readBorderSideValue(Element side) {
        // Handle the relevant branch before the state changes.
        if (side == null) {
            return new BorderSideValue();
        }
        
        BorderSideValue sideValue = new BorderSideValue();
        
        String style = side.getAttribute("style");
        sideValue.setStyle(parseBorderStyle(style));
        
        Element colorElement = findElement(side, "color");
        sideValue.setColor(readColorValue(colorElement));
        
        return sideValue;
    }
    
    /**
     * Reads a color value from an XML element.
     */
    private static ColorValue readColorValue(Element colorElement) {
        // Handle the relevant branch before the state changes.
        if (colorElement == null) {
            return new ColorValue((byte)0, (byte)0, (byte)0, (byte)0);
        }
        
        String rgb = colorElement.getAttribute("rgb");
        if (rgb == null || rgb.trim().isEmpty()) {
            return new ColorValue((byte)0, (byte)0, (byte)0, (byte)0);
        }
        
        rgb = rgb.trim().toUpperCase(Locale.ROOT);
        if (rgb.length() == 6) {
            rgb = "FF" + rgb;
        }
        if (rgb.length() != 8) {
            return new ColorValue((byte)0, (byte)0, (byte)0, (byte)0);
        }
        
        try {
            byte a = (byte) Integer.parseInt(rgb.substring(0, 2), 16);
            byte r = (byte) Integer.parseInt(rgb.substring(2, 4), 16);
            byte g = (byte) Integer.parseInt(rgb.substring(4, 6), 16);
            byte b = (byte) Integer.parseInt(rgb.substring(6, 8), 16);
            return new ColorValue(a, r, g, b);
        } catch (NumberFormatException e) {
            return new ColorValue((byte)0, (byte)0, (byte)0, (byte)0);
        }
    }
    
    /**
     * Builds a color element.
     */
    private static Element buildColorElement(String elementName, ColorValue color) {
        // Handle the relevant branch before the state changes.
        if (isEmptyColor(color)) {
            return null;
        }
        
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element colorElement = doc.createElementNS(MAIN_NS, elementName);
            colorElement.setAttribute("rgb", toArgbHex(color));
            return colorElement;
        } catch (Exception e) {
            return null;
        }
    }
    
    /**
     * Builds a border side element.
     */
    private static Element buildBorderSideElement(String sideName, BorderSideValue side) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element element = doc.createElementNS(MAIN_NS, sideName);
            
            String styleName = getBorderStyleName(side.getStyle());
            if (styleName != null && !styleName.isEmpty()) {
                element.setAttribute("style", styleName);
            }
            
            Element colorElement = buildColorElement("color", side.getColor());
            if (colorElement != null) {
                element.appendChild(colorElement);
            }
            
            return element;
        } catch (Exception e) {
            throw new RuntimeException("Failed to build border side element", e);
        }
    }
    
    /**
     * Builds an alignment element.
     */
    private static Element buildAlignmentElement(AlignmentValue alignment) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element element = doc.createElementNS(MAIN_NS, "alignment");
            boolean hasValue = false;
            
            String horizontal = getHorizontalAlignmentName(alignment.getHorizontal());
            if (horizontal != null && !horizontal.isEmpty()) {
                element.setAttribute("horizontal", horizontal);
                hasValue = true;
            }
            
            String vertical = getVerticalAlignmentName(alignment.getVertical());
            if (vertical != null && !vertical.isEmpty()) {
                element.setAttribute("vertical", vertical);
                hasValue = true;
            }
            
            if (alignment.getWrapText()) {
                element.setAttribute("wrapText", "1");
                hasValue = true;
            }
            
            if (alignment.getIndentLevel() > 0) {
                element.setAttribute("indent", String.valueOf(alignment.getIndentLevel()));
                hasValue = true;
            }
            
            if (alignment.getTextRotation() != 0) {
                element.setAttribute("textRotation", String.valueOf(alignment.getTextRotation()));
                hasValue = true;
            }
            
            if (alignment.getShrinkToFit()) {
                element.setAttribute("shrinkToFit", "1");
                hasValue = true;
            }
            
            if (alignment.getReadingOrder() != 0) {
                element.setAttribute("readingOrder", String.valueOf(alignment.getReadingOrder()));
                hasValue = true;
            }
            
            if (alignment.getRelativeIndent() != 0) {
                element.setAttribute("relativeIndent", String.valueOf(alignment.getRelativeIndent()));
                hasValue = true;
            }
            
            return hasValue ? element : null;
        } catch (Exception e) {
            return null;
        }
    }
    
    /**
     * Builds a protection element.
     */
    private static Element buildProtectionElement(ProtectionValue protection) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.newDocument();
            
            Element element = doc.createElementNS(MAIN_NS, "protection");
            boolean hasValue = false;
            
            if (!protection.getIsLocked()) {
                element.setAttribute("locked", "0");
                hasValue = true;
            }
            
            if (protection.getIsHidden()) {
                element.setAttribute("hidden", "1");
                hasValue = true;
            }
            
            return hasValue ? element : null;
        } catch (Exception e) {
            return null;
        }
    }
    
    /**
     * Finds an XML element by local name.
     */
    private static Element findElement(Element parent, String localName) {
        NodeList nodes = parent.getElementsByTagNameNS(MAIN_NS, localName);
        // Handle the relevant branch before the state changes.
        if (nodes.getLength() > 0) {
            return (Element) nodes.item(0);
        }
        return null;
    }
    
    /**
     * Parses a boolean attribute value.
     */
    private static boolean parseBoolAttribute(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isEmpty()) {
            return false;
        }
        return "1".equals(value) || Boolean.parseBoolean(value);
    }
    
    /**
     * Parses an integer attribute value.
     */
    private static int parseIntAttribute(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isEmpty()) {
            return 0;
        }
        try {
            return Integer.parseInt(value);
        } catch (NumberFormatException e) {
            return 0;
        }
    }
    
    
    /**
     * Parses a border style from a string value.
     */
    private static com.aspose.cells_foss.core.BorderStyle parseBorderStyle(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isEmpty()) {
            return com.aspose.cells_foss.core.BorderStyle.NONE;
        }
        
        switch (value.toLowerCase(Locale.ROOT)) {
            case "thin":
                return com.aspose.cells_foss.core.BorderStyle.THIN;
            case "medium":
                return com.aspose.cells_foss.core.BorderStyle.MEDIUM;
            case "thick":
                return com.aspose.cells_foss.core.BorderStyle.THICK;
            case "hair":
                return com.aspose.cells_foss.core.BorderStyle.HAIR;
            case "dotted":
                return com.aspose.cells_foss.core.BorderStyle.DOTTED;
            case "dashed":
                return com.aspose.cells_foss.core.BorderStyle.DASHED;
            case "double":
                return com.aspose.cells_foss.core.BorderStyle.DOUBLE;
            default:
                return com.aspose.cells_foss.core.BorderStyle.NONE;
        }
    }
    
    /**
     * Gets the border style name for an XML attribute.
     */
    private static String getBorderStyleName(com.aspose.cells_foss.core.BorderStyle style) {
        // Translate the internal value into the matching public representation.
        switch (style) {
            case THIN:
                return "thin";
            case MEDIUM:
                return "medium";
            case THICK:
                return "thick";
            case HAIR:
                return "hair";
            case DOTTED:
                return "dotted";
            case DASHED:
                return "dashed";
            case DOUBLE:
                return "double";
            default:
                return null;
        }
    }
    
    /**
     * Gets the horizontal alignment name for an XML attribute.
     */
    private static String getHorizontalAlignmentName(HorizontalAlignment alignment) {
        // Translate the internal value into the matching public representation.
        switch (alignment) {
            case LEFT:
                return "left";
            case CENTER:
                return "center";
            case RIGHT:
                return "right";
            case FILL:
                return "fill";
            case JUSTIFY:
                return "justify";
            case CENTER_CONTINUOUS:
                return "centerContinuous";
            case DISTRIBUTED:
                return "distributed";
            default:
                return null;
        }
    }
    
    /**
     * Gets the vertical alignment name for an XML attribute.
     */
    private static String getVerticalAlignmentName(VerticalAlignment alignment) {
        // Translate the internal value into the matching public representation.
        switch (alignment) {
            case TOP:
                return "top";
            case CENTER:
                return "center";
            case BOTTOM:
                return "bottom";
            case JUSTIFY:
                return "justify";
            case DISTRIBUTED:
                return "distributed";
            default:
                return null;
        }
    }
    
    /**
     * Checks if a color is empty.
     */
    private static boolean isEmptyColor(ColorValue color) {
        return color.getA() == 0 && color.getR() == 0 && color.getG() == 0 && color.getB() == 0;
    }
    
    /**
     * Converts a color to ARGB hex string.
     */
    private static String toArgbHex(ColorValue color) {
        String a = String.format("%02X", color.getA() & 0xFF);
        String r = String.format("%02X", color.getR() & 0xFF);
        String g = String.format("%02X", color.getG() & 0xFF);
        String b = String.format("%02X", color.getB() & 0xFF);
        return a + r + g + b;
    }
    
    /**
     * Compares two fonts for equality.
     */
    private static boolean fontEquals(FontValue font1, FontValue font2) {
        return font1.getName().equals(font2.getName()) &&
               font1.getSize() == font2.getSize() &&
               font1.getBold() == font2.getBold() &&
               font1.getItalic() == font2.getItalic() &&
               font1.getUnderline() == font2.getUnderline() &&
               font1.getStrikeThrough() == font2.getStrikeThrough() &&
               colorEquals(font1.getColor(), font2.getColor());
    }
    
    /**
     * Compares two colors for equality.
     */
    private static boolean colorEquals(ColorValue color1, ColorValue color2) {
        return color1.getA() == color2.getA() &&
               color1.getR() == color2.getR() &&
               color1.getG() == color2.getG() &&
               color1.getB() == color2.getB();
    }
    
    /**
     * Compares two border sets for equality.
     */
    private static boolean bordersEqual(BordersValue borders1, BordersValue borders2) {
        return borderSideEquals(borders1.getLeft(), borders2.getLeft()) &&
               borderSideEquals(borders1.getRight(), borders2.getRight()) &&
               borderSideEquals(borders1.getTop(), borders2.getTop()) &&
               borderSideEquals(borders1.getBottom(), borders2.getBottom()) &&
               borderSideEquals(borders1.getDiagonal(), borders2.getDiagonal()) &&
               borders1.getDiagonalUp() == borders2.getDiagonalUp() &&
               borders1.getDiagonalDown() == borders2.getDiagonalDown();
    }
    
    /**
     * Compares two border side values for equality.
     */
    private static boolean borderSideEquals(BorderSideValue side1, BorderSideValue side2) {
        return side1.getStyle() == side2.getStyle() &&
               colorEquals(side1.getColor(), side2.getColor());
    }
    
    /**
     * Serializes the document to a byte array.
     */
    public static byte[] getDocumentBytes(Document doc) {
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            transformer.setOutputProperty("encoding", "UTF-8");
            transformer.setOutputProperty("indent", "yes");
            
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            transformer.transform(new DOMSource(doc), new StreamResult(out));
            return out.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException("Failed to serialize XML document", e);
        }
    }
}
