package org.aspose.cells_foss;

/**
 * Factory for page-setup test workbooks.
 * Mirrors C# PageSetupScenarioFactory from shared test infrastructure.
 */
public final class PageSetupScenarioFactory {

    /**
     * Verifies that page setup scenario factory.
     */
    private PageSetupScenarioFactory() {}

    /**
     * Verifies that create page setup workbook.
     */
    public static Workbook createPageSetupWorkbook() {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);
        sheet.setName("Print Sheet");
        sheet.getCells().get("A1").putValue("Title");
        sheet.getCells().get("C10").putValue(42);

        PageSetup ps = sheet.getPageSetup();
        ps.setLeftMarginInch(0.25d);
        ps.setRightMarginInch(0.4d);
        ps.setTopMarginInch(0.5d);
        ps.setBottomMarginInch(0.6d);
        ps.setHeaderMarginInch(0.2d);
        ps.setFooterMarginInch(0.22d);
        ps.setOrientation(PageOrientationType.LANDSCAPE);
        ps.setPaperSize(PaperSizeType.PAPER_A4);
        ps.setFirstPageNumber(3);
        ps.setScale(95);
        ps.setFitToPagesWide(1);
        ps.setFitToPagesTall(2);
        ps.setPrintArea("$A$1:$C$10");
        ps.setPrintTitleRows("$1:$2");
        ps.setPrintTitleColumns("$A:$B");
        ps.setLeftHeader("Left Header");
        ps.setCenterHeader("Center Header");
        ps.setRightHeader("Right Header");
        ps.setLeftFooter("Left Footer");
        ps.setCenterFooter("Center Footer");
        ps.setRightFooter("Right Footer");
        ps.setPrintGridlines(true);
        ps.setPrintHeadings(true);
        ps.setCenterHorizontally(true);
        ps.setCenterVertically(true);
        ps.addHorizontalPageBreak(4);
        ps.addHorizontalPageBreak(7);
        ps.addVerticalPageBreak(2);
        return workbook;
    }

    /**
     * Verifies that assert page setup.
     * @param workbook workbook to apply
     */
    public static void assertPageSetup(Workbook workbook) {
        PageSetup ps = workbook.getWorksheets().get(0).getPageSetup();
        AssertEx.assertEqual(0.25d, ps.getLeftMarginInch());
        AssertEx.assertEqual(0.4d, ps.getRightMarginInch());
        AssertEx.assertEqual(0.5d, ps.getTopMarginInch());
        AssertEx.assertEqual(0.6d, ps.getBottomMarginInch());
        AssertEx.assertEqual(0.2d, ps.getHeaderMarginInch());
        AssertEx.assertEqual(0.22d, ps.getFooterMarginInch());
        AssertEx.assertEqual(PageOrientationType.LANDSCAPE, ps.getOrientation());
        AssertEx.assertEqual(PaperSizeType.PAPER_A4, ps.getPaperSize());
        AssertEx.assertEqual(3, (int) ps.getFirstPageNumber());
        AssertEx.assertEqual(95, (int) ps.getScale());
        AssertEx.assertEqual(1, (int) ps.getFitToPagesWide());
        AssertEx.assertEqual(2, (int) ps.getFitToPagesTall());
        AssertEx.assertEqual("$A$1:$C$10", ps.getPrintArea());
        AssertEx.assertEqual("$1:$2", ps.getPrintTitleRows());
        AssertEx.assertEqual("$A:$B", ps.getPrintTitleColumns());
        AssertEx.assertEqual("Left Header", ps.getLeftHeader());
        AssertEx.assertEqual("Center Header", ps.getCenterHeader());
        AssertEx.assertEqual("Right Header", ps.getRightHeader());
        AssertEx.assertEqual("Left Footer", ps.getLeftFooter());
        AssertEx.assertEqual("Center Footer", ps.getCenterFooter());
        AssertEx.assertEqual("Right Footer", ps.getRightFooter());
        AssertEx.assertTrue(ps.getPrintGridlines());
        AssertEx.assertTrue(ps.getPrintHeadings());
        AssertEx.assertTrue(ps.getCenterHorizontally());
        AssertEx.assertTrue(ps.getCenterVertically());
        AssertEx.assertEqual(2, ps.getHorizontalPageBreaks().size());
        AssertEx.assertEqual(4, (int) ps.getHorizontalPageBreaks().get(0));
        AssertEx.assertEqual(7, (int) ps.getHorizontalPageBreaks().get(1));
        AssertEx.assertEqual(1, ps.getVerticalPageBreaks().size());
        AssertEx.assertEqual(2, (int) ps.getVerticalPageBreaks().get(0));
        AssertEx.assertEqual("Title", workbook.getWorksheets().get(0).getCells().get("A1").getStringValue());
        AssertEx.assertEqual("42", workbook.getWorksheets().get(0).getCells().get("C10").getStringValue());
    }
}

