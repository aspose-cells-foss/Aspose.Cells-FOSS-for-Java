package com.aspose.cells_foss;

import static org.junit.jupiter.api.Assertions.*;

/** Factory for shape test workbooks. Mirrors C# ShapeScenarioFactory. */
public final class ShapeScenarioFactory {

    private ShapeScenarioFactory() {}

    public static Workbook createShapeWorkbook() {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Shapes");

        sheet.getCells().get("A1").putValue("Drawing Objects Demo");

        int rectIdx = sheet.getShapes().add(2, 0, 6, 4, AutoShapeType.RECTANGLE);
        sheet.getShapes().get(rectIdx).setName("MyRectangle");

        int arrowIdx = sheet.getShapes().add(8, 0, 11, 4, AutoShapeType.RIGHT_ARROW);
        sheet.getShapes().get(arrowIdx).setName("MyRightArrow");

        int ellipseIdx = sheet.getShapes().add(13, 0, 17, 4, AutoShapeType.ELLIPSE);
        sheet.getShapes().get(ellipseIdx).setName("MyEllipse");

        int starIdx = sheet.getShapes().add(2, 5, 7, 9, AutoShapeType.STAR5_POINT);
        sheet.getShapes().get(starIdx).setName("MyStar5");

        int downArrowIdx = sheet.getShapes().add(9, 5, 12, 9, AutoShapeType.DOWN_ARROW);
        sheet.getShapes().get(downArrowIdx).setName("MyDownArrow");

        int star12Idx = sheet.getShapes().add(14, 5, 19, 9, AutoShapeType.STAR12_POINT);
        sheet.getShapes().get(star12Idx).setName("MyStar12");

        return wb;
    }

    public static void assertShapeWorkbook(Workbook wb) {
        Worksheet sheet = wb.getWorksheets().get(0);
        assertEquals(6, sheet.getShapes().getCount());

        Shape rect = sheet.getShapes().get(0);
        assertEquals("MyRectangle", rect.getName());
        assertEquals(2, rect.getUpperLeftRow());
        assertEquals(0, rect.getUpperLeftColumn());
        assertEquals(6, rect.getLowerRightRow());
        assertEquals(4, rect.getLowerRightColumn());
        assertEquals(AutoShapeType.RECTANGLE, rect.getAutoShapeType());

        assertEquals("MyRightArrow", sheet.getShapes().get(1).getName());
        assertEquals(AutoShapeType.RIGHT_ARROW, sheet.getShapes().get(1).getAutoShapeType());

        assertEquals("MyEllipse", sheet.getShapes().get(2).getName());
        assertEquals(AutoShapeType.ELLIPSE, sheet.getShapes().get(2).getAutoShapeType());

        assertEquals("MyStar5", sheet.getShapes().get(3).getName());
        assertEquals(AutoShapeType.STAR5_POINT, sheet.getShapes().get(3).getAutoShapeType());

        assertEquals("MyDownArrow", sheet.getShapes().get(4).getName());
        assertEquals(AutoShapeType.DOWN_ARROW, sheet.getShapes().get(4).getAutoShapeType());

        assertEquals("MyStar12", sheet.getShapes().get(5).getName());
        assertEquals(AutoShapeType.STAR12_POINT, sheet.getShapes().get(5).getAutoShapeType());
    }
}
