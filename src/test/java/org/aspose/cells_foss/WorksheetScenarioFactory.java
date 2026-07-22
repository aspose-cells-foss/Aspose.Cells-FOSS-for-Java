package org.aspose.cells_foss;

/**
 * Factory for worksheet-settings test workbooks.
 * Mirrors C# WorksheetScenarioFactory from shared test infrastructure.
 */
public final class WorksheetScenarioFactory {

    /**
     * Verifies that worksheet scenario factory.
     */
    private WorksheetScenarioFactory() {}

    /**
     * Verifies that create worksheet settings workbook.
     */
    public static Workbook createWorksheetSettingsWorkbook() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);
        sheet.setName("Layout");
        sheet.setVisibilityType(VisibilityType.HIDDEN);
        sheet.setTabColor(Color.fromArgb(255, 34, 68, 102));
        sheet.setShowGridlines(false);
        sheet.setShowRowColumnHeaders(false);
        sheet.setShowZeros(false);
        sheet.setRightToLeft(true);
        sheet.setZoom(85);
        sheet.protect();
        sheet.getProtection().setAllowEditingObject(false);
        sheet.getProtection().setAllowEditingScenario(false);
        sheet.getProtection().setAllowFormattingCell(false);
        sheet.getProtection().setAllowInsertingRow(false);
        sheet.getProtection().setAllowFiltering(false);
        sheet.getProtection().setAllowSelectingLockedCell(false);
        sheet.getProtection().setAllowSelectingUnlockedCell(false);

        sheet.getCells().get("A1").putValue("Merged");
        sheet.getCells().get("C4").putValue(99);
        sheet.getCells().getRows().get(1).setHeight(22.5d);
        sheet.getCells().getRows().get(3).setHidden(true);
        sheet.getCells().getColumns().get(0).setWidth(18.25d);
        sheet.getCells().getColumns().get(2).setHidden(true);
        sheet.getCells().merge(0, 0, 2, 2);

        int visibleSheetIndex = workbook.getWorksheets().add();
        Worksheet visibleSheet = workbook.getWorksheets().get(visibleSheetIndex);
        visibleSheet.setName("Visible");
        visibleSheet.getCells().get("A1").putValue("Visible");
        workbook.getWorksheets().setActiveSheetName("Visible");
        return workbook;
    }

    /**
     * Verifies that assert worksheet settings.
     * @param workbook workbook to apply
     */
    public static void assertWorksheetSettings(Workbook workbook) {
        Worksheet sheet = workbook.getWorksheets().get("Layout");
        AssertEx.assertNotNull(sheet, "Expected the worksheet settings scenario to contain the 'Layout' sheet.");
        AssertEx.assertEqual(VisibilityType.HIDDEN, sheet.getVisibilityType());
        AssertEx.assertEqual(Color.fromArgb(255, 34, 68, 102), sheet.getTabColor());
        AssertEx.assertFalse(sheet.getShowGridlines());
        AssertEx.assertFalse(sheet.getShowRowColumnHeaders());
        AssertEx.assertFalse(sheet.getShowZeros());
        AssertEx.assertTrue(sheet.getRightToLeft());
        AssertEx.assertEqual(85, sheet.getZoom());
        AssertEx.assertTrue(sheet.getProtection().isProtected());
        AssertEx.assertFalse(sheet.getProtection().getAllowEditingObject());
        AssertEx.assertFalse(sheet.getProtection().getAllowEditingScenario());
        AssertEx.assertFalse(sheet.getProtection().getAllowFormattingCell());
        AssertEx.assertFalse(sheet.getProtection().getAllowInsertingRow());
        AssertEx.assertFalse(sheet.getProtection().getAllowFiltering());
        AssertEx.assertFalse(sheet.getProtection().getAllowSelectingLockedCell());
        AssertEx.assertFalse(sheet.getProtection().getAllowSelectingUnlockedCell());
        AssertEx.assertEqual("Merged", sheet.getCells().get("A1").getStringValue());
        AssertEx.assertEqual(99, (int) sheet.getCells().get("C4").getValue());
        AssertEx.assertEqual(22.5d, sheet.getCells().getRows().get(1).getHeight());
        AssertEx.assertTrue(sheet.getCells().getRows().get(3).isHidden());
        AssertEx.assertEqual(18.25d, sheet.getCells().getColumns().get(0).getWidth());
        AssertEx.assertTrue(sheet.getCells().getColumns().get(2).isHidden());
        AssertEx.assertEqual(1, sheet.getCells().getMergedCells().size());
        AssertEx.assertEqual(0, sheet.getCells().getMergedCells().get(0).getFirstRow());
        AssertEx.assertEqual(0, sheet.getCells().getMergedCells().get(0).getFirstColumn());
        AssertEx.assertEqual(2, sheet.getCells().getMergedCells().get(0).getTotalRows());
        AssertEx.assertEqual(2, sheet.getCells().getMergedCells().get(0).getTotalColumns());
    }

    /**
     * Verifies that assert worksheet settings scenario has visible sheet.
     * @param workbook workbook to apply
     */
    public static void assertWorksheetSettingsScenarioHasVisibleSheet(Workbook workbook) {
        AssertEx.assertTrue(workbook.getWorksheets().getCount() >= 2);
        Worksheet visibleSheet = workbook.getWorksheets().get("Visible");
        AssertEx.assertNotNull(visibleSheet, "Expected a visible sheet named 'Visible'.");
        AssertEx.assertEqual(VisibilityType.VISIBLE, visibleSheet.getVisibilityType());
        AssertEx.assertEqual("Visible", visibleSheet.getCells().get("A1").getStringValue());
        AssertEx.assertEqual("Visible", workbook.getWorksheets().getActiveSheetName());
    }
}

