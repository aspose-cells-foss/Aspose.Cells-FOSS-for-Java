package org.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.List;

/** Internal model for a drawing object (auto shape) anchored to a worksheet. */
public final class ShapeModel {
    private String name = "";
    private int upperLeftRow;
    private int upperLeftColumn;
    private long upperLeftRowOffset;
    private long upperLeftColumnOffset;
    private int lowerRightRow;
    private int lowerRightColumn;
    private long lowerRightRowOffset;
    private long lowerRightColumnOffset;
    private long extentCx;
    private long extentCy;
    private String geometryType = "rect";
    private String rawStyleXml;
    private String rawTxBodyXml;
    /** Non-null for connectors, group shapes, or other non-standard elements 鈥?emitted verbatim. */
    private String rawElementXml;

    /**
     * Drawing-level rels referenced by rawElementXml that aren't covered by pictures/charts.
     * Each entry is [relId, relType, target]. Emitted verbatim to the drawing rels file.
     */
    private final List<String[]> extraDrawingRels = new ArrayList<>();

    /**
     * Images embedded within this shape (e.g. inside a grpSp) that are not standalone pictures.
     * Keyed by drawing-rels rId (e.g. "rId2") 鈫?raw image bytes.
     */
    private final java.util.Map<String, byte[]> embeddedImageData = new java.util.LinkedHashMap<>();

    public String getName() { return name; }
    public void setName(String name) { this.name = name != null ? name : ""; }

    public int getUpperLeftRow() { return upperLeftRow; }
    public void setUpperLeftRow(int v) { upperLeftRow = v; }

    public int getUpperLeftColumn() { return upperLeftColumn; }
    public void setUpperLeftColumn(int v) { upperLeftColumn = v; }

    public long getUpperLeftRowOffset() { return upperLeftRowOffset; }
    public void setUpperLeftRowOffset(long v) { upperLeftRowOffset = v; }

    public long getUpperLeftColumnOffset() { return upperLeftColumnOffset; }
    public void setUpperLeftColumnOffset(long v) { upperLeftColumnOffset = v; }

    public int getLowerRightRow() { return lowerRightRow; }
    public void setLowerRightRow(int v) { lowerRightRow = v; }

    public int getLowerRightColumn() { return lowerRightColumn; }
    public void setLowerRightColumn(int v) { lowerRightColumn = v; }

    public long getLowerRightRowOffset() { return lowerRightRowOffset; }
    public void setLowerRightRowOffset(long v) { lowerRightRowOffset = v; }

    public long getLowerRightColumnOffset() { return lowerRightColumnOffset; }
    public void setLowerRightColumnOffset(long v) { lowerRightColumnOffset = v; }

    public long getExtentCx() { return extentCx; }
    public void setExtentCx(long v) { extentCx = v; }

    public long getExtentCy() { return extentCy; }
    public void setExtentCy(long v) { extentCy = v; }

    public String getGeometryType() { return geometryType; }
    public void setGeometryType(String v) { geometryType = (v == null || v.isEmpty()) ? "rect" : v; }

    public String getRawStyleXml() { return rawStyleXml; }
    public void setRawStyleXml(String v) { rawStyleXml = v; }

    public String getRawTxBodyXml() { return rawTxBodyXml; }
    public void setRawTxBodyXml(String v) { rawTxBodyXml = v; }

    public String getRawElementXml() { return rawElementXml; }
    public void setRawElementXml(String v) { rawElementXml = v; }

    public List<String[]> getExtraDrawingRels() { return extraDrawingRels; }
    public java.util.Map<String, byte[]> getEmbeddedImageData() { return embeddedImageData; }
}

