package org.aspose.cells_foss;

import org.aspose.cells_foss.core.ChartModel;
import java.util.List;

/** Collection of embedded charts on a worksheet. */
public final class ChartCollection {
    private final List<ChartModel> models;

    ChartCollection(List<ChartModel> models) {
        this.models = models;
    }

    public int getCount() { return models.size(); }

    public Chart get(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("Chart index out of range.");
        return new Chart(models.get(index));
    }

    /**
     * Adds a new chart of the given type with the specified data range and anchor.
     * ChartEx types (Waterfall, Treemap, Sunburst, Histogram, BoxAndWhisker, Funnel, Map)
     * are not supported for programmatic creation and will throw UnsupportedFeatureException.
     *
     * @param type           chart type
     * @param dataRange      cell range formula, e.g. "Sheet1!$B$1:$B$5"
     * @param upperLeftRow   zero-based upper-left row anchor
     * @param upperLeftColumn zero-based upper-left column anchor
     * @param lowerRightRow  zero-based lower-right row anchor
     * @param lowerRightColumn zero-based lower-right column anchor
     * @return zero-based index of the added chart
     */
    public int add(ChartType type, String dataRange,
                   int upperLeftRow, int upperLeftColumn,
                   int lowerRightRow, int lowerRightColumn) {
        switch (type) {
            case WATERFALL, TREEMAP, SUNBURST, HISTOGRAM, BOX_AND_WHISKER, FUNNEL, MAP ->
                throw new UnsupportedFeatureException(
                    "ChartEx type " + type + " is not supported for programmatic chart creation.");
            default -> {}
        }
        if (dataRange == null || dataRange.isBlank())
            throw new CellsException("dataRange must not be blank.");
        if (upperLeftRow < 0 || upperLeftColumn < 0)
            throw new CellsException("Anchor coordinates must be non-negative.");

        ChartModel model = new ChartModel();
        model.setName("Chart " + (models.size() + 1));
        model.setChartType(chartTypeToString(type));
        model.setUpperLeftRow(upperLeftRow);
        model.setUpperLeftColumn(upperLeftColumn);
        model.setLowerRightRow(lowerRightRow);
        model.setLowerRightColumn(lowerRightColumn);
        byte[] xml = XlsxWorkbookChartTemplates.build(type, dataRange);
        model.setRawChartXml(new String(xml, java.nio.charset.StandardCharsets.UTF_8));
        models.add(model);
        return models.size() - 1;
    }

    // -------------------------------------------------------------------------
    // Package-private helpers
    // -------------------------------------------------------------------------

    public static ChartType parseChartType(String raw) {
        if (raw == null) return ChartType.UNKNOWN;
        return switch (raw.toLowerCase()) {
            case "bar"                -> ChartType.BAR;
            case "column"             -> ChartType.COLUMN;
            case "line"               -> ChartType.LINE;
            case "area"               -> ChartType.AREA;
            case "pie"                -> ChartType.PIE;
            case "doughnut"           -> ChartType.DOUGHNUT;
            case "scatter"            -> ChartType.SCATTER;
            case "bubble"             -> ChartType.BUBBLE;
            case "radar"              -> ChartType.RADAR;
            case "stock"              -> ChartType.STOCK;
            case "bar3d"              -> ChartType.BAR_3D;
            case "column3d"           -> ChartType.COLUMN_3D;
            case "line3d"             -> ChartType.LINE_3D;
            case "area3d"             -> ChartType.AREA_3D;
            case "pie3d"              -> ChartType.PIE_3D;
            case "surface3d"          -> ChartType.SURFACE_3D;
            case "surfacewireframe3d" -> ChartType.SURFACE_WIREFRAME_3D;
            case "contour"            -> ChartType.CONTOUR;
            case "waterfall"          -> ChartType.WATERFALL;
            case "treemap"            -> ChartType.TREEMAP;
            case "sunburst"           -> ChartType.SUNBURST;
            case "histogram"          -> ChartType.HISTOGRAM;
            case "boxandwhisker"      -> ChartType.BOX_AND_WHISKER;
            case "funnel"             -> ChartType.FUNNEL;
            case "map"                -> ChartType.MAP;
            default                   -> ChartType.UNKNOWN;
        };
    }

    public static String chartTypeToString(ChartType type) {
        return switch (type) {
            case BAR                -> "bar";
            case COLUMN             -> "column";
            case LINE               -> "line";
            case AREA               -> "area";
            case PIE                -> "pie";
            case DOUGHNUT           -> "doughnut";
            case SCATTER            -> "scatter";
            case BUBBLE             -> "bubble";
            case RADAR              -> "radar";
            case STOCK              -> "stock";
            case BAR_3D             -> "bar3d";
            case COLUMN_3D          -> "column3d";
            case LINE_3D            -> "line3d";
            case AREA_3D            -> "area3d";
            case PIE_3D             -> "pie3d";
            case SURFACE_3D         -> "surface3d";
            case SURFACE_WIREFRAME_3D -> "surfacewireframe3d";
            case CONTOUR            -> "contour";
            case WATERFALL          -> "waterfall";
            case TREEMAP            -> "treemap";
            case SUNBURST           -> "sunburst";
            case HISTOGRAM          -> "histogram";
            case BOX_AND_WHISKER    -> "boxandwhisker";
            case FUNNEL             -> "funnel";
            case MAP                -> "map";
            default                 -> "unknown";
        };
    }
}

