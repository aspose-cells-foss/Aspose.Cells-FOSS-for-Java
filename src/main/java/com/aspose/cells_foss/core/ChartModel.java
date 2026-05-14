package com.aspose.cells_foss.core;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Internal model for an embedded chart.
 */
public final class ChartModel {
    private String name = "";
    private String chartType = "unknown";
    private int upperLeftRow;
    private int upperLeftColumn;
    private int lowerRightRow;
    private int lowerRightColumn;
    private long extentCx;
    private long extentCy;
    private boolean isChartEx = false;

    /** Full verbatim XML of the chart definition (xl/charts/chart{N}.xml). */
    private String rawChartXml;

    /** Full verbatim XML of the mc:AlternateContent element (ChartEx only). */
    private String rawGraphicFrameXml;

    /**
     * Chart-level relationships loaded from xl/charts/_rels/chart{N}.xml.rels.
     * Each entry is [relId, relType, originalTarget].
     */
    private final List<String[]> chartRels = new ArrayList<>();

    /**
     * Raw bytes for each chart rel, keyed by relId.
     * Covers XML companion files (style, colors) and binary images.
     */
    private final Map<String, byte[]> chartRelContent = new LinkedHashMap<>();

    public String getName() { return name; }
    public void setName(String name) { this.name = name != null ? name : ""; }

    public String getChartType() { return chartType; }
    public void setChartType(String chartType) { this.chartType = chartType != null ? chartType : "unknown"; }

    public int getUpperLeftRow() { return upperLeftRow; }
    public void setUpperLeftRow(int upperLeftRow) { this.upperLeftRow = upperLeftRow; }

    public int getUpperLeftColumn() { return upperLeftColumn; }
    public void setUpperLeftColumn(int upperLeftColumn) { this.upperLeftColumn = upperLeftColumn; }

    public int getLowerRightRow() { return lowerRightRow; }
    public void setLowerRightRow(int lowerRightRow) { this.lowerRightRow = lowerRightRow; }

    public int getLowerRightColumn() { return lowerRightColumn; }
    public void setLowerRightColumn(int lowerRightColumn) { this.lowerRightColumn = lowerRightColumn; }

    public long getExtentCx() { return extentCx; }
    public void setExtentCx(long extentCx) { this.extentCx = extentCx; }

    public long getExtentCy() { return extentCy; }
    public void setExtentCy(long extentCy) { this.extentCy = extentCy; }

    public boolean isChartEx() { return isChartEx; }
    public void setChartEx(boolean chartEx) { this.isChartEx = chartEx; }

    public String getRawChartXml() { return rawChartXml; }
    public void setRawChartXml(String rawChartXml) { this.rawChartXml = rawChartXml; }

    public String getRawGraphicFrameXml() { return rawGraphicFrameXml; }
    public void setRawGraphicFrameXml(String rawGraphicFrameXml) { this.rawGraphicFrameXml = rawGraphicFrameXml; }

    public List<String[]> getChartRels() { return chartRels; }
    public Map<String, byte[]> getChartRelContent() { return chartRelContent; }

    /** Returns how many image-type rels this chart has (used for global image counter). */
    public int getChartImageCount() {
        return (int) chartRels.stream()
            .filter(r -> r[1].contains("/relationships/image"))
            .count();
    }

    private String originalRelId;
    public String getOriginalRelId() { return originalRelId; }
    public void setOriginalRelId(String v) { this.originalRelId = v; }
}
