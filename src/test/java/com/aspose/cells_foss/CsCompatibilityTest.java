package com.aspose.cells_foss;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Java port of C# CompatibilityTests/Program.cs.
 * Uses Apache POI (test scope) for XML verification where appropriate.
 * Input files are loaded from src/test/resources/Input/.
 */
class CsCompatibilityTest {

    @TempDir Path tempDir;

    // =========================================================================
    // Load / save overloads
    // =========================================================================

    @Test
    void file_and_stream_load_paths_produce_same_values() throws IOException {
        Workbook workbook = WorkbookScenarioFactory.createMixedCellWorkbook();
        Path path = tempDir.resolve("book.xlsx");
        workbook.save(path.toString());

        try (InputStream stream = Files.newInputStream(path)) {
            Workbook fromFile   = new Workbook(path.toString());
            Workbook fromStream = new Workbook(stream);
            WorkbookScenarioFactory.assertWorkbookDataEquals(fromFile, fromStream);
        }
    }

    @Test
    void save_overloads_produce_equivalent_workbooks() throws IOException {
        Workbook workbook = WorkbookScenarioFactory.createMixedCellWorkbook();
        Path filePath = tempDir.resolve("book-file.xlsx");
        workbook.save(filePath.toString(), SaveFormat.XLSX);

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.save(bos, SaveFormat.XLSX);

        Workbook fromFile   = new Workbook(filePath.toString());
        Workbook fromStream = new Workbook(new ByteArrayInputStream(bos.toByteArray()));
        WorkbookScenarioFactory.assertWorkbookDataEquals(fromFile, fromStream);
    }

    @Test
    void formula_setter_accepts_with_or_without_leading_equal() {
        Workbook wb = new Workbook();
        Cell cell = wb.getWorksheets().get(0).getCells().get("A1");
        cell.putValue(10);
        cell.setFormula("B1+C1");
        assertEquals("=B1+C1", cell.getFormula());
        cell.setFormula("=D1+E1");
        assertEquals("=D1+E1", cell.getFormula());
    }

    @Test
    void exception_mapping_uses_cells_exception_types() {
        assertThrows(CellsException.class,
            () -> new Workbook().getWorksheets().get("missing"));
        assertThrows(CellsException.class,
            () -> new Workbook().getWorksheets().get(0).getCells().get("1A"));
        assertThrows(InvalidFileFormatException.class,
            () -> new Workbook(new ByteArrayInputStream(new byte[]{1, 2, 3, 4})));
    }

    @Test
    void public_type_mapping_matches_after_roundtrip() throws IOException {
        Workbook wb = WorkbookScenarioFactory.createMixedCellWorkbook();
        Path path = tempDir.resolve("types.xlsx");
        wb.save(path.toString());

        Workbook loaded = new Workbook(path.toString());
        assertEquals(CellValueType.STRING,    loaded.getWorksheets().get(0).getCells().get("A1").getType());
        assertEquals(CellValueType.NUMBER,    loaded.getWorksheets().get(0).getCells().get("B1").getType());
        assertEquals(CellValueType.BOOLEAN,   loaded.getWorksheets().get(0).getCells().get("C1").getType());
        assertEquals(CellValueType.NUMBER,    loaded.getWorksheets().get(0).getCells().get("D1").getType());
        assertEquals(CellValueType.NUMBER,    loaded.getWorksheets().get(0).getCells().get("E1").getType());
        assertEquals(CellValueType.DATE_TIME, loaded.getWorksheets().get(0).getCells().get("F1").getType());
        assertEquals(CellValueType.FORMULA,   loaded.getWorksheets().get(0).getCells().get("G1").getType());
    }

    @Test
    void value_property_setter_matches_supported_scalar_behavior() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.getCells().get("A1").setValue("alpha");
        sheet.getCells().get("B1").setValue(12);
        sheet.getCells().get("C1").setValue(true);
        sheet.getCells().get("D1").setValue(LocalDateTime.of(2024, 1, 2, 3, 4, 0));
        sheet.getCells().get("E1").setValue(null);

        assertEquals("alpha", sheet.getCells().get("A1").getValue());
        assertEquals(12, ((Number)sheet.getCells().get("B1").getValue()).intValue());
        assertEquals(true, sheet.getCells().get("C1").getValue());
        assertEquals(CellValueType.DATE_TIME, sheet.getCells().get("D1").getType());
        assertEquals("", sheet.getCells().get("E1").getDisplayStringValue());
    }

    // =========================================================================
    // Worksheet view / protection
    // =========================================================================

    @Test
    void worksheet_view_members_follow_supported_patterns() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        sheet.setTabColor(Color.fromArgb(255, 34, 68, 102));
        sheet.setShowGridlines(false);
        sheet.setShowRowColumnHeaders(false);
        sheet.setShowZeros(false);
        sheet.setRightToLeft(true);
        sheet.setZoom(85);

        assertEquals(Color.fromArgb(255, 34, 68, 102), sheet.getTabColor());
        assertFalse(sheet.getShowGridlines());
        assertFalse(sheet.getShowRowColumnHeaders());
        assertFalse(sheet.getShowZeros());
        assertTrue(sheet.getRightToLeft());
        assertEquals(85, sheet.getZoom());
    }

    @Test
    void worksheet_protection_members_follow_supported_patterns() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);

        sheet.protect();
        sheet.getProtection().setAllowEditingObject(false);
        sheet.getProtection().setAllowEditingScenario(false);
        sheet.getProtection().setAllowFiltering(false);
        sheet.getProtection().setAllowSelectingLockedCell(false);
        sheet.getProtection().setAllowSelectingUnlockedCell(false);

        assertTrue(sheet.getProtection().isProtected());
        assertFalse(sheet.getProtection().getAllowEditingObject());
        assertFalse(sheet.getProtection().getAllowEditingScenario());
        assertFalse(sheet.getProtection().getAllowFiltering());
        assertFalse(sheet.getProtection().getAllowSelectingLockedCell());
        assertFalse(sheet.getProtection().getAllowSelectingUnlockedCell());
    }

    // =========================================================================
    // AutoFilter roundtrip
    // =========================================================================

    @Test
    void autofilter_members_follow_supported_patterns() {
        Workbook wb = AutoFilterScenarioFactory.createAutoFilterWorkbook();
        AutoFilterScenarioFactory.assertAutoFilter(wb);
    }

    @Test
    void file_and_stream_roundtrip_preserve_autofilter() throws IOException {
        Workbook wb = AutoFilterScenarioFactory.createAutoFilterWorkbook();
        Path path = tempDir.resolve("af.xlsx");
        wb.save(path.toString());
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.save(bos, SaveFormat.XLSX);

        AutoFilterScenarioFactory.assertAutoFilter(new Workbook(path.toString()));
        AutoFilterScenarioFactory.assertAutoFilter(new Workbook(new ByteArrayInputStream(bos.toByteArray())));
    }

    @Test
    void excel_input_autofilter_roundtrip_preserves_header_only_range() throws IOException {
        Path inputPath = resolveInput("Autofilter.xlsx");
        Workbook wb = new Workbook(inputPath.toString());
        assertEquals("A2:C2", wb.getWorksheets().get(0).getAutoFilter().getRange());
        assertEquals(0, wb.getWorksheets().get(0).getAutoFilter().getFilterColumns().getCount());

        Path out = tempDir.resolve("af-rt.xlsx");
        wb.save(out.toString());
        Workbook reloaded = new Workbook(out.toString());
        assertEquals("A2:C2", reloaded.getWorksheets().get(0).getAutoFilter().getRange());
        assertEquals(0, reloaded.getWorksheets().get(0).getAutoFilter().getFilterColumns().getCount());

        // Verify worksheet XML
        String wsXml = ZipPackageHelper.readEntryText(out.toString(), "xl/worksheets/sheet1.xml");
        assertTrue(wsXml.contains("A2:C2"), "autoFilter ref must survive roundtrip");
    }

    // =========================================================================
    // Defined names roundtrip
    // =========================================================================

    @Test
    void file_and_stream_roundtrip_preserve_defined_names() throws IOException {
        Workbook wb = DefinedNameScenarioFactory.createDefinedNamesWorkbook();
        Path path = tempDir.resolve("dn.xlsx");
        wb.save(path.toString());
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.save(bos, SaveFormat.XLSX);

        DefinedNameScenarioFactory.assertDefinedNames(new Workbook(path.toString()));
        DefinedNameScenarioFactory.assertDefinedNames(new Workbook(new ByteArrayInputStream(bos.toByteArray())));
    }

    // =========================================================================
    // Styles roundtrip
    // =========================================================================

    @Test
    void file_and_stream_roundtrip_preserve_styles() throws IOException {
        Workbook wb = StyleScenarioFactory.createStyledWorkbook();
        Path path = tempDir.resolve("styled.xlsx");
        wb.save(path.toString());
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.save(bos, SaveFormat.XLSX);

        StyleScenarioFactory.assertPrimaryStyle(new Workbook(path.toString()).getWorksheets().get(0).getCells().get("A1").getStyle());
        StyleScenarioFactory.assertPrimaryStyle(new Workbook(new ByteArrayInputStream(bos.toByteArray())).getWorksheets().get(0).getCells().get("A1").getStyle());
        StyleScenarioFactory.assertCustomNumberStyle(new Workbook(path.toString()).getWorksheets().get(0).getCells().get("B2").getStyle());
        assertEquals(CellValueType.BLANK, new Workbook(path.toString()).getWorksheets().get(0).getCells().get("B2").getType());
    }

    // =========================================================================
    // Worksheet settings roundtrip
    // =========================================================================

    @Test
    void file_and_stream_roundtrip_preserve_worksheet_settings() throws IOException {
        Workbook wb = WorksheetScenarioFactory.createWorksheetSettingsWorkbook();
        Path path = tempDir.resolve("ws.xlsx");
        wb.save(path.toString());
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.save(bos, SaveFormat.XLSX);

        WorksheetScenarioFactory.assertWorksheetSettings(new Workbook(path.toString()));
        WorksheetScenarioFactory.assertWorksheetSettings(new Workbook(new ByteArrayInputStream(bos.toByteArray())));
    }

    // =========================================================================
    // Validation roundtrip
    // =========================================================================

    @Test
    void file_and_stream_roundtrip_preserve_data_validations() throws IOException {
        Workbook wb = ValidationScenarioFactory.createValidationWorkbook();
        Path path = tempDir.resolve("val.xlsx");
        wb.save(path.toString());
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.save(bos, SaveFormat.XLSX);

        ValidationScenarioFactory.assertValidations(new Workbook(path.toString()));
        ValidationScenarioFactory.assertValidations(new Workbook(new ByteArrayInputStream(bos.toByteArray())));
    }

    // =========================================================================
    // Conditional formatting roundtrip
    // =========================================================================

    @Test
    void file_and_stream_roundtrip_preserve_conditional_formattings() throws IOException {
        Workbook wb = ConditionalFormattingScenarioFactory.createConditionalFormattingWorkbook();
        Path path = tempDir.resolve("cf.xlsx");
        wb.save(path.toString());
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.save(bos, SaveFormat.XLSX);

        ConditionalFormattingScenarioFactory.assertConditionalFormattings(new Workbook(path.toString()));
        ConditionalFormattingScenarioFactory.assertConditionalFormattings(new Workbook(new ByteArrayInputStream(bos.toByteArray())));
    }

    @Test
    void file_and_stream_roundtrip_preserve_advanced_conditional_formattings() throws IOException {
        Workbook wb = ConditionalFormattingScenarioFactory.createAdvancedConditionalFormattingWorkbook();
        Path path = tempDir.resolve("acf.xlsx");
        wb.save(path.toString());
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.save(bos, SaveFormat.XLSX);

        ConditionalFormattingScenarioFactory.assertAdvancedConditionalFormattings(new Workbook(path.toString()));
        ConditionalFormattingScenarioFactory.assertAdvancedConditionalFormattings(new Workbook(new ByteArrayInputStream(bos.toByteArray())));
    }

    // =========================================================================
    // Page setup roundtrip
    // =========================================================================

    @Test
    void file_and_stream_roundtrip_preserve_page_setup() throws IOException {
        Workbook wb = PageSetupScenarioFactory.createPageSetupWorkbook();
        Path path = tempDir.resolve("ps.xlsx");
        wb.save(path.toString());
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.save(bos, SaveFormat.XLSX);

        PageSetupScenarioFactory.assertPageSetup(new Workbook(path.toString()));
        PageSetupScenarioFactory.assertPageSetup(new Workbook(new ByteArrayInputStream(bos.toByteArray())));
    }

    // =========================================================================
    // Table roundtrip (from Input file)
    // =========================================================================

    @Test
    void excel_input_table_roundtrip_preserves_table_structure() throws IOException {
        Path inputPath = resolveInput("table.xlsx");
        Workbook wb = new Workbook(inputPath.toString());
        Worksheet sheet = wb.getWorksheets().get(0);

        assertEquals(1, sheet.getListObjects().getCount());
        ListObject table = sheet.getListObjects().get(0);
        assertEquals("Table5", table.getDisplayName());
        assertEquals(0, table.getStartRow());
        assertEquals(0, table.getStartColumn());
        assertEquals(3, table.getEndRow());
        assertEquals(2, table.getEndColumn());
        assertTrue(table.isShowHeaderRow());
        assertFalse(table.isShowTotals());
        assertEquals(TableStyleType.TABLE_STYLE_MEDIUM_2, table.getTableStyleType());
        assertEquals(3, table.getListColumns().getCount());
        assertEquals("Product", table.getListColumns().get(0).getName());
        assertEquals("Region",  table.getListColumns().get(1).getName());
        assertEquals("Revenue", table.getListColumns().get(2).getName());

        Path out = tempDir.resolve("table-rt.xlsx");
        wb.save(out.toString());

        Workbook reloaded = new Workbook(out.toString());
        ListObject rt = reloaded.getWorksheets().get(0).getListObjects().get(0);
        assertEquals("Table5",  rt.getDisplayName());
        assertEquals(3, rt.getEndRow());
        assertEquals(TableStyleType.TABLE_STYLE_MEDIUM_2, rt.getTableStyleType());
        assertEquals("Product", rt.getListColumns().get(0).getName());
        assertEquals("Revenue", rt.getListColumns().get(2).getName());

        String tableXml = ZipPackageHelper.readEntryText(out.toString(), "xl/tables/table1.xml");
        assertTrue(tableXml.contains("A1:C4"), "Table ref must survive roundtrip");
        assertTrue(tableXml.contains("TableStyleMedium2"), "Table style must survive roundtrip");
    }

    // =========================================================================
    // Picture roundtrip (from Input file)
    // =========================================================================

    @Test
    void excel_input_picture_roundtrip_preserves_picture_structure() throws IOException {
        Path inputPath = resolveInput("compare/17.6.3-ByCells.xlsx");
        Workbook wb = new Workbook(inputPath.toString());
        Worksheet sheet = wb.getWorksheets().get(0);

        assertTrue(sheet.getPictures().getCount() > 0, "Expected at least one picture");
        Picture pic = sheet.getPictures().get(0);
        assertTrue(pic.getData() != null && pic.getData().length > 0, "Picture data must be non-empty");
        assertEquals(ImageType.JPEG, pic.getImageType());

        int origLen   = pic.getData().length;
        int origULR   = pic.getUpperLeftRow();
        int origULC   = pic.getUpperLeftColumn();
        int origLRR   = pic.getLowerRightRow();
        int origLRC   = pic.getLowerRightColumn();

        Path out = tempDir.resolve("pic-rt.xlsx");
        wb.save(out.toString());

        Workbook reloaded = new Workbook(out.toString());
        Worksheet rs = reloaded.getWorksheets().get(0);
        assertEquals(sheet.getPictures().getCount(), rs.getPictures().getCount());

        Picture rpic = rs.getPictures().get(0);
        assertEquals(origLen, rpic.getData().length);
        assertEquals(origULR, rpic.getUpperLeftRow());
        assertEquals(origULC, rpic.getUpperLeftColumn());
        assertEquals(origLRR, rpic.getLowerRightRow());
        assertEquals(origLRC, rpic.getLowerRightColumn());
        assertEquals(ImageType.JPEG, rpic.getImageType());
    }

    // =========================================================================
    // Chart roundtrip tests (from Input directories)
    // =========================================================================

    @Test
    void excel_input_area_chart_roundtrip()      { roundtripDir("AreaChart",     true); }
    @Test
    void excel_input_bar_chart_roundtrip()       { roundtripDir("BarChart",      true); }
    @Test
    void excel_input_column_chart_roundtrip()    { roundtripDir("ColumnChart",   true); }
    @Test
    void excel_input_line_chart_roundtrip()      { roundtripDir("linechart",     true); }
    @Test
    void excel_input_pie_chart_roundtrip()       { roundtripDir("PieChart",      true); }
    @Test
    void excel_input_xy_chart_roundtrip()        { roundtripDir("XYChart",       true); }
    @Test
    void excel_input_radar_chart_roundtrip()     { roundtripDir("RadarChart",    true); }
    @Test
    void excel_input_surface_chart_roundtrip()   { roundtripDir("SurfaceChart",  true); }
    @Test
    void excel_input_stock_chart_roundtrip()     { roundtripDir("StockChart",    true); }

    // ChartEx types — may not detect chart type, but files must roundtrip without crash
    @Test
    void excel_input_funnel_chart_roundtrip()    { roundtripDir("FunnelChart",   false); }
    @Test
    void excel_input_histogram_chart_roundtrip() { roundtripDir("HistogramChart",false); }
    @Test
    void excel_input_box_chart_roundtrip()       { roundtripDir("BoxChart",      false); }
    @Test
    void excel_input_sunburst_chart_roundtrip()  { roundtripDir("SunburstChart", false); }
    @Test
    void excel_input_treemap_chart_roundtrip()   { roundtripDir("TreemapChart",  false); }
    @Test
    void excel_input_map_chart_roundtrip()       { roundtripDir("MapChart",      false); }
    @Test
    void excel_input_combo_chart_roundtrip()     { roundtripDir("combochart",    false); }
    @Test
    void excel_input_waterfall_chart_roundtrip() { roundtripDir("waterfallchart",false); }
    @Test
    void excel_input_allcharts_roundtrip()       { roundtripDir("allcharts",     false); }
    @Test
    void excel_input_sparkline_roundtrip_no_crash() { roundtripDir("Sparkline",  false); }

    // =========================================================================
    // Create each programmable chart type as a separate xlsx file
    // =========================================================================

    private static final ChartType[] PROGRAMMABLE_CHART_TYPES = {
        ChartType.COLUMN,               ChartType.BAR,
        ChartType.LINE,                 ChartType.AREA,
        ChartType.PIE,                  ChartType.DOUGHNUT,
        ChartType.SCATTER,              ChartType.BUBBLE,
        ChartType.RADAR,                ChartType.STOCK,
        ChartType.COLUMN_3D,            ChartType.BAR_3D,
        ChartType.LINE_3D,              ChartType.AREA_3D,
        ChartType.PIE_3D,               ChartType.SURFACE_3D,
        ChartType.SURFACE_WIREFRAME_3D, ChartType.CONTOUR,
    };

    @Test
    void create_each_chart_type_as_separate_file() throws IOException {
        java.io.File outDir = new java.io.File("output/outcharts");
        outDir.mkdirs();

        for (ChartType type : PROGRAMMABLE_CHART_TYPES) {
            String sheetName = type.name();
            Workbook wb = ChartScenarioFactory.createChartWorkbook(type, sheetName);

            String outPath = "output/outcharts/" + type.name().toLowerCase() + ".xlsx";
            wb.save(outPath);
            assertTrue(new java.io.File(outPath).exists(),
                "Output file should exist: " + outPath);

            // Verify round-trip
            Workbook reloaded = new Workbook(outPath);
            ChartScenarioFactory.assertChartWorkbook(reloaded, type);
        }
    }

    // =========================================================================
    // Create all programmable chart types in one workbook
    // =========================================================================

    @Test
    void create_all_chart_types_and_save_to_output() throws IOException {
        String outPath = "output/all_charts.xlsx";
        try (Workbook wb = new Workbook()) {
            for (int i = 0; i < PROGRAMMABLE_CHART_TYPES.length; i++) {
                ChartType type = PROGRAMMABLE_CHART_TYPES[i];
                String sheetName = type.name();

                Worksheet sheet;
                if (i == 0) {
                    sheet = wb.getWorksheets().get(0);
                    sheet.setName(sheetName);
                } else {
                    int idx = wb.getWorksheets().add();
                    sheet = wb.getWorksheets().get(idx);
                    sheet.setName(sheetName);
                }

                String[] labels = {"Q1", "Q2", "Q3", "Q4", "Q5"};
                int[]    values = {10,   30,   20,   50,   40};
                for (int r = 0; r < 5; r++) {
                    sheet.getCells().get(r, 0).putValue(labels[r]);
                    sheet.getCells().get(r, 1).putValue(values[r]);
                }

                String dataRange = sheetName + "!$B$1:$B$5";
                int chartIdx = sheet.getCharts().add(type, dataRange, 7, 0, 22, 10);
                assertEquals(0, chartIdx);
                assertEquals(type, sheet.getCharts().get(0).getChartType());
            }

            new java.io.File("output").mkdirs();
            wb.save(outPath);
        }
        assertTrue(new java.io.File(outPath).exists(), "Output file should exist");

        try (Workbook reloaded = new Workbook(outPath)) {
            assertEquals(PROGRAMMABLE_CHART_TYPES.length, reloaded.getWorksheets().getCount());
            for (int i = 0; i < PROGRAMMABLE_CHART_TYPES.length; i++) {
                Worksheet sheet = reloaded.getWorksheets().get(i);
                assertEquals(PROGRAMMABLE_CHART_TYPES[i].name(), sheet.getName());
                assertEquals(1, sheet.getCharts().getCount(),
                    "Sheet " + sheet.getName() + " should have 1 chart");
                assertEquals(PROGRAMMABLE_CHART_TYPES[i], sheet.getCharts().get(0).getChartType(),
                    "Wrong chart type on sheet " + sheet.getName());
            }
        }
    }

    // =========================================================================
    // Save all chart files to output folder
    // =========================================================================

    @Test
    void save_all_chart_files_to_output() throws IOException {
        String[] chartDirs = {
            "AreaChart", "BarChart", "ColumnChart", "linechart", "PieChart",
            "XYChart", "RadarChart", "SurfaceChart", "StockChart",
            "FunnelChart", "HistogramChart", "BoxChart", "SunburstChart",
            "TreemapChart", "MapChart", "combochart", "waterfallchart",
            "allcharts", "Sparkline"
        };

        java.io.File outDir = new java.io.File("output/charts");
        outDir.mkdirs();

        int savedCount = 0;
        for (String subDir : chartDirs) {
            URL dirUrl = getClass().getClassLoader().getResource("Input/" + subDir);
            if (dirUrl == null) continue;

            java.io.File dir;
            try { dir = new java.io.File(dirUrl.toURI()); }
            catch (java.net.URISyntaxException e) { continue; }

            java.io.File[] files = dir.listFiles((d, n) -> n.endsWith(".xlsx"));
            if (files == null) continue;

            for (java.io.File file : files) {
                try (Workbook wb = new Workbook(file.getAbsolutePath())) {
                    boolean hasChart = false;
                    for (int i = 0; i < wb.getWorksheets().getCount(); i++) {
                        if (wb.getWorksheets().get(i).getCharts().getCount() > 0) {
                            hasChart = true;
                            break;
                        }
                    }
                    if (!hasChart) continue;

                    String outPath = "output/charts/" + subDir + "_" + file.getName();
                    wb.save(outPath);
                    assertTrue(new java.io.File(outPath).exists(),
                        "Output file should exist: " + outPath);
                    savedCount++;
                }
            }
        }

        assertTrue(savedCount > 0, "Expected to save at least one chart file");
        System.out.println("Saved " + savedCount + " chart files to output/charts/");
    }

    // =========================================================================
    // Shape roundtrip
    // =========================================================================

    @Test
    void shape_programmatic_roundtrip_preserves_all_shapes() throws IOException {
        Workbook wb = ShapeScenarioFactory.createShapeWorkbook();
        ShapeScenarioFactory.assertShapeWorkbook(wb);

        Path out = tempDir.resolve("shapes.xlsx");
        wb.save(out.toString());

        Workbook reloaded = new Workbook(out.toString());
        ShapeScenarioFactory.assertShapeWorkbook(reloaded);

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.save(bos, SaveFormat.XLSX);
        ShapeScenarioFactory.assertShapeWorkbook(new Workbook(new ByteArrayInputStream(bos.toByteArray())));
    }

    @Test
    void excel_input_shape_xlsx_roundtrip_preserves_shape_count() throws IOException {
        URL url = getClass().getClassLoader().getResource("Input/shape.xlsx");
        assertNotNull(url, "shape.xlsx not found in test resources");

        Path src;
        try { src = Path.of(url.toURI()); }
        catch (Exception e) { throw new IOException(e); }

        Workbook wb = new Workbook(src.toString());
        Worksheet sheet = wb.getWorksheets().get(0);
        int originalCount = sheet.getShapes().getCount();
        assertTrue(originalCount > 0, "shape.xlsx should contain at least one shape");

        Path out = tempDir.resolve("shape-roundtrip.xlsx");
        wb.save(out.toString());

        Workbook reloaded = new Workbook(out.toString());
        assertEquals(originalCount, reloaded.getWorksheets().get(0).getShapes().getCount(),
            "Shape count must be preserved through roundtrip");
    }

    @Test
    void shape_all_types_saved_to_output() throws IOException {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("AllShapes");

        AutoShapeType[] types = {
            AutoShapeType.RECTANGLE, AutoShapeType.ROUNDED_RECTANGLE, AutoShapeType.ELLIPSE,
            AutoShapeType.TRIANGLE, AutoShapeType.RIGHT_TRIANGLE, AutoShapeType.DIAMOND,
            AutoShapeType.PENTAGON, AutoShapeType.HEXAGON, AutoShapeType.OCTAGON,
            AutoShapeType.PLUS, AutoShapeType.CUBE, AutoShapeType.CYLINDER,
            AutoShapeType.HEART, AutoShapeType.LIGHTNING, AutoShapeType.SUN,
            AutoShapeType.MOON, AutoShapeType.CLOUD,
            AutoShapeType.RIGHT_ARROW, AutoShapeType.LEFT_ARROW, AutoShapeType.UP_ARROW,
            AutoShapeType.DOWN_ARROW, AutoShapeType.LEFT_RIGHT_ARROW, AutoShapeType.UP_DOWN_ARROW,
            AutoShapeType.STAR4_POINT, AutoShapeType.STAR5_POINT, AutoShapeType.STAR6_POINT,
            AutoShapeType.STAR8_POINT, AutoShapeType.STAR12_POINT, AutoShapeType.STAR16_POINT,
            AutoShapeType.STAR24_POINT, AutoShapeType.STAR32_POINT,
            AutoShapeType.MATH_PLUS, AutoShapeType.STRAIGHT_CONNECTOR,
        };

        // Lay shapes out in a grid: 3 columns, each shape 3 rows tall × 2 cols wide
        for (int i = 0; i < types.length; i++) {
            int col = (i % 3) * 3;
            int row = (i / 3) * 4;
            int idx = sheet.getShapes().add(row, col, row + 3, col + 2, types[i]);
            sheet.getShapes().get(idx).setName(types[i].name());
        }

        assertEquals(types.length, sheet.getShapes().getCount());

        // Save to output folder
        java.io.File outDir = new java.io.File("output");
        outDir.mkdirs();
        String outPath = "output/all_shapes.xlsx";
        wb.save(outPath);
        assertTrue(new java.io.File(outPath).exists(), "Output file should exist");

        // Verify roundtrip
        Workbook reloaded = new Workbook(outPath);
        assertEquals(types.length, reloaded.getWorksheets().get(0).getShapes().getCount());
        for (int i = 0; i < types.length; i++) {
            assertEquals(types[i].name(), reloaded.getWorksheets().get(0).getShapes().get(i).getName());
        }
    }

    // =========================================================================
    // Comment roundtrip
    // =========================================================================

    @Test
    void comment_programmatic_roundtrip_preserves_count() throws IOException {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.getComments().add(0, 0).setNote("First");
        sheet.getComments().add(1, 1).setNote("Second");
        sheet.getComments().add(2, 0).setNote("Third");

        Path out = tempDir.resolve("comment-count.xlsx");
        wb.save(out.toString());

        Workbook reloaded = new Workbook(out.toString());
        assertEquals(3, reloaded.getWorksheets().get(0).getComments().getCount(),
            "Comment count must be preserved after save/load roundtrip");
    }

    @Test
    void comment_programmatic_roundtrip_preserves_note_and_author() throws IOException {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        Comment c = sheet.getComments().add(3, 5);
        c.setNote("Important: review before merge");
        c.setAuthor("Carol");
        c.setVisible(true);
        c.setWidth(180);
        c.setHeight(90);

        Path out = tempDir.resolve("comment-content.xlsx");
        wb.save(out.toString());

        Workbook reloaded = new Workbook(out.toString());
        Worksheet ls = reloaded.getWorksheets().get(0);
        assertEquals(1, ls.getComments().getCount());
        Comment loaded = ls.getComments().get("F4");
        assertNotNull(loaded, "Comment at F4 must survive roundtrip");
        assertEquals("Important: review before merge", loaded.getNote());
        assertEquals("Carol", loaded.getAuthor());
        assertTrue(loaded.isVisible(), "IsVisible must be preserved");
        assertEquals(180, loaded.getWidth(),  "Width must be preserved");
        assertEquals(90,  loaded.getHeight(), "Height must be preserved");
    }

    // =========================================================================
    // Private helpers
    // =========================================================================

    private void roundtripDir(String subDir, boolean expectAtLeastOneChart) {
        URL dirUrl = getClass().getClassLoader().getResource("Input/" + subDir);
        assertNotNull(dirUrl, "Input directory not found: Input/" + subDir);

        File dir;
        try { dir = new File(dirUrl.toURI()); } catch (java.net.URISyntaxException e) { throw new RuntimeException(e); }
        File[] files = dir.listFiles((d, n) -> n.endsWith(".xlsx"));
        assertNotNull(files);
        assertTrue(files.length > 0, "No xlsx files in Input/" + subDir);

        for (File file : files) {
            Workbook original = new Workbook(file.getAbsolutePath());
            int origChartCount = original.getWorksheets().get(0).getCharts().getCount();

            if (expectAtLeastOneChart) {
                assertTrue(origChartCount > 0,
                    "Expected at least one chart in: " + file.getName());
            }

            Path out = tempDir.resolve(subDir + "_" + file.getName());
            original.save(out.toString());

            Workbook reloaded = new Workbook(out.toString());
            assertEquals(origChartCount, reloaded.getWorksheets().get(0).getCharts().getCount(),
                "Chart count changed after roundtrip for: " + file.getName());
        }
    }

    private static Path resolveInput(String relativePath) {
        URL url = CsCompatibilityTest.class.getClassLoader().getResource("Input/" + relativePath);
        assertNotNull(url, "Input file not found: Input/" + relativePath);
        try { return Path.of(url.toURI()); } catch (java.net.URISyntaxException e) { throw new RuntimeException(e); }
    }
}
