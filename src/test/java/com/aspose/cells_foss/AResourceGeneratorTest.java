package com.aspose.cells_foss;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import java.io.*;
import java.net.URL;
import java.nio.file.*;
import java.util.LinkedHashMap;
import java.util.Map;

class AResourceGeneratorTest {

    private static final byte[] JPEG_1X1 = {
        (byte)0xFF,(byte)0xD8,(byte)0xFF,(byte)0xE0,0x00,0x10,0x4A,0x46,
        0x49,0x46,0x00,0x01,0x00,0x00,0x00,0x01,0x00,0x01,0x00,0x00,
        (byte)0xFF,(byte)0xDB,0x00,0x43,0x00,0x08,0x06,0x06,0x07,0x06,
        0x05,0x08,0x07,0x07,0x07,0x09,0x09,0x08,0x0A,0x0C,0x14,0x0D,
        0x0C,0x0B,0x0B,0x0C,0x19,0x12,0x13,0x0F,0x14,0x1D,0x1A,0x1F,
        0x1E,0x1D,0x1A,0x1C,0x1C,0x20,0x24,0x2E,0x27,0x20,0x22,0x2C,
        0x23,0x1C,0x1C,0x28,0x37,0x29,0x2C,0x30,0x31,0x34,0x34,0x34,
        0x1F,0x27,0x39,0x3D,0x38,0x32,0x3C,0x2E,0x33,0x34,0x32,
        (byte)0xFF,(byte)0xC0,0x00,0x0B,0x08,0x00,0x01,0x00,0x01,
        0x01,0x01,0x11,0x00,
        (byte)0xFF,(byte)0xC4,0x00,0x14,0x00,0x01,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x03,
        (byte)0xFF,(byte)0xC4,0x00,0x14,0x10,0x01,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        (byte)0xFF,(byte)0xDA,0x00,0x08,0x01,0x01,0x00,0x00,0x3F,0x00,0x7F,
        (byte)0xFF,(byte)0xD9
    };

    private static final byte[] PNG_1X1 = {
        (byte)0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
        0x08,0x02,0x00,0x00,0x00,(byte)0x90,0x77,0x53,(byte)0xDE,
        0x00,0x00,0x00,0x0C,0x49,0x44,0x41,0x54,
        0x08,(byte)0xD7,0x63,(byte)0xF8,(byte)0xCF,(byte)0xC0,0x00,0x00,
        0x00,0x02,0x00,0x01,(byte)0xE2,0x21,(byte)0xBC,0x33,
        0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
        (byte)0xAE,0x42,0x60,(byte)0x82
    };

    @Test
    void generate_test_resources() throws Exception {
        Path cpInput = resolveClasspathInput();
        Path srcInput = Paths.get("src", "test", "resources", "Input");
        Files.createDirectories(cpInput);
        Files.createDirectories(srcInput);
        writeBytes(cpInput.resolve("pay.jpg"),    JPEG_1X1);
        writeBytes(srcInput.resolve("pay.jpg"),   JPEG_1X1);
        writeBytes(cpInput.resolve("screen.png"), PNG_1X1);
        writeBytes(srcInput.resolve("screen.png"),PNG_1X1);
        generateTableXlsx(cpInput, srcInput);
        generateCompareXlsx(cpInput, srcInput);
        generateShapeXlsx(cpInput, srcInput);
        generateChartDirs(cpInput, srcInput);
        generateAutofilterXlsx(cpInput, srcInput);
    }

    private static void generateAutofilterXlsx(Path cp, Path src) throws IOException {
        Workbook wb = new Workbook();
        Worksheet ws = wb.getWorksheets().get(0);
        ws.getCells().get("A1").putValue("Title");
        ws.getCells().get("A2").putValue("Name");
        ws.getCells().get("B2").putValue("Region");
        ws.getCells().get("C2").putValue("Score");
        ws.getCells().get("A3").putValue("Alice"); ws.getCells().get("B3").putValue("North"); ws.getCells().get("C3").putValue(90);
        ws.getCells().get("A4").putValue("Bob");   ws.getCells().get("B4").putValue("South"); ws.getCells().get("C4").putValue(80);
        ws.getAutoFilter().setRange("A2:C2");
        saveWorkbook(wb, cp.resolve("Autofilter.xlsx"));
        saveWorkbook(wb, src.resolve("Autofilter.xlsx"));
    }

    private static Path resolveClasspathInput() throws Exception {
        URL root = AResourceGeneratorTest.class.getClassLoader().getResource(".");
        if (root != null) { return Paths.get(root.toURI()).resolve("Input"); }
        return Paths.get("target", "test-classes", "Input");
    }

    private static void writeBytes(Path path, byte[] data) throws IOException {
        if (!Files.exists(path)) { Files.createDirectories(path.getParent()); Files.write(path, data); }
    }

    private static void saveWorkbook(Workbook wb, Path path) throws IOException {
        if (!Files.exists(path)) { Files.createDirectories(path.getParent()); wb.save(path.toString()); }
    }

    private static void generateTableXlsx(Path cp, Path src) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("Sheet1");
        org.apache.poi.ss.usermodel.Row hdr = sheet.createRow(0);
        hdr.createCell(0).setCellValue("Product");
        hdr.createCell(1).setCellValue("Region");
        hdr.createCell(2).setCellValue("Revenue");
        String[][] rows = {{"Apple","North","100"},{"Banana","South","200"},{"Cherry","East","300"}};
        for (int i = 0; i < rows.length; i++) {
            org.apache.poi.ss.usermodel.Row r = sheet.createRow(i + 1);
            r.createCell(0).setCellValue(rows[i][0]);
            r.createCell(1).setCellValue(rows[i][1]);
            r.createCell(2).setCellValue(Double.parseDouble(rows[i][2]));
        }
        AreaReference area = new AreaReference("A1:C4", SpreadsheetVersion.EXCEL2007);
        XSSFTable table = sheet.createTable(area);
        table.setName("Table5"); table.setDisplayName("Table5");
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo tsi = table.getCTTable().addNewTableStyleInfo();
        tsi.setName("TableStyleMedium2"); tsi.setShowFirstColumn(false); tsi.setShowLastColumn(false);
        tsi.setShowRowStripes(true); tsi.setShowColumnStripes(false);
        writePoi(wb, cp.resolve("table.xlsx")); writePoi(wb, src.resolve("table.xlsx")); wb.close();
    }

    private static void writePoi(XSSFWorkbook wb, Path path) throws IOException {
        if (!Files.exists(path)) { Files.createDirectories(path.getParent());
            try (OutputStream os = Files.newOutputStream(path)) { wb.write(os); } }
    }

    private static void generateCompareXlsx(Path cp, Path src) throws IOException {
        Workbook wb = new Workbook();
        wb.getWorksheets().get(0).getPictures().add(0, 0, 5, 5, JPEG_1X1);
        saveWorkbook(wb, cp.resolve("compare").resolve("17.6.3-ByCells.xlsx"));
        saveWorkbook(wb, src.resolve("compare").resolve("17.6.3-ByCells.xlsx"));
    }

    private static void generateShapeXlsx(Path cp, Path src) throws IOException {
        Workbook wb = new Workbook();
        wb.getWorksheets().get(0).getShapes().add(0, 0, 5, 5, AutoShapeType.RECTANGLE);
        saveWorkbook(wb, cp.resolve("shape.xlsx")); saveWorkbook(wb, src.resolve("shape.xlsx"));
    }

    private static void generateChartDirs(Path cp, Path src) throws IOException {
        Map<String, ChartType> cd = new LinkedHashMap<>();
        cd.put("ColumnChart",ChartType.COLUMN); cd.put("AreaChart",ChartType.AREA);
        cd.put("BarChart",ChartType.BAR); cd.put("linechart",ChartType.LINE);
        cd.put("PieChart",ChartType.PIE); cd.put("XYChart",ChartType.SCATTER);
        cd.put("RadarChart",ChartType.RADAR); cd.put("SurfaceChart",ChartType.SURFACE_3D);
        cd.put("StockChart",ChartType.STOCK);
        for (Map.Entry<String, ChartType> e : cd.entrySet()) {
            Workbook wb = ChartScenarioFactory.createChartWorkbook(e.getValue(), "Sheet1");
            saveWorkbook(wb, cp.resolve(e.getKey()).resolve("chart.xlsx"));
            saveWorkbook(wb, src.resolve(e.getKey()).resolve("chart.xlsx"));
        }
        String[] simple = {"FunnelChart","HistogramChart","BoxChart","SunburstChart",
            "TreemapChart","MapChart","combochart","waterfallchart","allcharts","Sparkline"};
        for (String d : simple) {
            Workbook wb = new Workbook();
            saveWorkbook(wb, cp.resolve(d).resolve("chart.xlsx"));
            saveWorkbook(wb, src.resolve(d).resolve("chart.xlsx"));
        }
    }
}
