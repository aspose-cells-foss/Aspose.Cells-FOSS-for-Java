package com.aspose.cells_foss;

import java.io.*;
import java.net.URL;

import static org.junit.jupiter.api.Assertions.*;

/** Factory for picture test workbooks. Mirrors C# PictureScenarioFactory. */
public final class PictureScenarioFactory {

    private PictureScenarioFactory() {}

    public static byte[] loadJpeg() {
        return loadResource("Input/pay.jpg");
    }

    public static byte[] loadPng() {
        return loadResource("Input/screen.png");
    }

    public static Workbook createPictureWorkbook() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Images");

        sheet.getCells().get("A1").putValue("Name");
        sheet.getCells().get("B1").putValue("Value");
        sheet.getCells().get("A2").putValue("Alpha");
        sheet.getCells().get("B2").putValue(42);

        byte[] jpeg = loadJpeg();
        int jpegIdx = sheet.getPictures().add(0, 0, 8, 5, jpeg);
        sheet.getPictures().get(jpegIdx).setName("PayImage");

        byte[] png = loadPng();
        int pngIdx = sheet.getPictures().add(10, 0, 16, 5, png);
        sheet.getPictures().get(pngIdx).setName("ScreenImage");

        return wb;
    }

    public static void assertPictureWorkbook(Workbook wb) {
        Worksheet sheet = wb.getWorksheets().get(0);
        assertEquals(2, sheet.getPictures().getCount());

        Picture jpeg = sheet.getPictures().get(0);
        assertEquals("PayImage",   jpeg.getName());
        assertEquals(0,  jpeg.getUpperLeftRow());
        assertEquals(0,  jpeg.getUpperLeftColumn());
        assertEquals(8,  jpeg.getLowerRightRow());
        assertEquals(5,  jpeg.getLowerRightColumn());
        assertEquals(ImageType.JPEG, jpeg.getImageType());
        assertEquals(loadJpeg().length, jpeg.getData().length);

        Picture png = sheet.getPictures().get(1);
        assertEquals("ScreenImage", png.getName());
        assertEquals(10, png.getUpperLeftRow());
        assertEquals(0,  png.getUpperLeftColumn());
        assertEquals(16, png.getLowerRightRow());
        assertEquals(5,  png.getLowerRightColumn());
        assertEquals(ImageType.PNG, png.getImageType());
        assertEquals(loadPng().length, png.getData().length);
    }

    private static byte[] loadResource(String path) {
        URL url = PictureScenarioFactory.class.getClassLoader().getResource(path);
        if (url == null) throw new RuntimeException("Resource not found: " + path);
        try (InputStream in = url.openStream()) {
            return in.readAllBytes();
        } catch (IOException e) {
            throw new RuntimeException("Failed to load resource: " + path, e);
        }
    }
}
