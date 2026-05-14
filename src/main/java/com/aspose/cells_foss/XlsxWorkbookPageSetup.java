package com.aspose.cells_foss;

import com.aspose.cells_foss.core.*;
import org.w3c.dom.*;
import java.util.List;

/**
 * Package-private helpers for building and loading XLSX page setup elements.
 */
final class XlsxWorkbookPageSetup {

    /**
     * Initializes a new XlsxWorkbookPageSetup instance.
     */
    private XlsxWorkbookPageSetup() {}

    // =========================================================================
    // Build
    // =========================================================================

    /**
     * Builds the page setup section.
     * @param ws ws
     * @param sb sb
     */
    static void buildPageSetupSection(WorksheetModel ws, StringBuilder sb) {
        PageSetupModel ps = ws.getPageSetup();

        // printOptions
        PrintOptionsModel po = ps.getPrintOptions();
        if (po.getGridLines() || po.getHeadings() || po.getHorizontalCentered() || po.getVerticalCentered()) {
            sb.append("<printOptions");
            if (po.getGridLines()) sb.append(" gridLines=\"1\"");
            if (po.getHeadings()) sb.append(" headings=\"1\"");
            if (po.getHorizontalCentered()) sb.append(" horizontalCentered=\"1\"");
            if (po.getVerticalCentered()) sb.append(" verticalCentered=\"1\"");
            sb.append("/>");
        }

        // pageMargins
        PageMarginsModel pm = ps.getMargins();
        sb.append("<pageMargins left=\"").append(XlsxWorkbookSerializerCommon.fmt(pm.getLeft()))
          .append("\" right=\"").append(XlsxWorkbookSerializerCommon.fmt(pm.getRight()))
          .append("\" top=\"").append(XlsxWorkbookSerializerCommon.fmt(pm.getTop()))
          .append("\" bottom=\"").append(XlsxWorkbookSerializerCommon.fmt(pm.getBottom()))
          .append("\" header=\"").append(XlsxWorkbookSerializerCommon.fmt(pm.getHeader()))
          .append("\" footer=\"").append(XlsxWorkbookSerializerCommon.fmt(pm.getFooter())).append("\"/>");

        // pageSetup — omit entirely when all settings are default
        int paperSizeCode = PaperSizeType.values()[ps.getPaperSize()].getValue();
        boolean isLandscape = ps.getOrientation() == PageOrientation.LANDSCAPE;
        boolean isPortrait  = ps.getOrientation() == PageOrientation.PORTRAIT;
        boolean hasNonDefaultPageSetup = paperSizeCode != 0 || isLandscape || isPortrait
                || ps.getScale() != null || ps.getFitToWidth() != null
                || ps.getFitToHeight() != null || ps.getFirstPageNumber() != null;
        if (hasNonDefaultPageSetup) {
            String orientStr = isLandscape ? "landscape" : isPortrait ? "portrait" : "default";
            sb.append("<pageSetup");
            if (paperSizeCode != 0) sb.append(" paperSize=\"").append(paperSizeCode).append("\"");
            sb.append(" orientation=\"").append(orientStr).append("\"");
            if (ps.getScale() != null) sb.append(" scale=\"").append(ps.getScale()).append("\"");
            if (ps.getFitToWidth() != null) sb.append(" fitToWidth=\"").append(ps.getFitToWidth()).append("\"");
            if (ps.getFitToHeight() != null) sb.append(" fitToHeight=\"").append(ps.getFitToHeight()).append("\"");
            if (ps.getFirstPageNumber() != null)
                sb.append(" firstPageNumber=\"").append(ps.getFirstPageNumber()).append("\" useFirstPageNumber=\"1\"");
            sb.append("/>");
        }

        // headerFooter
        HeaderFooterModel hf = ps.getHeaderFooter();
        String lh = XlsxWorkbookSerializerCommon.nvl(hf.getLeftHeader());
        String ch = XlsxWorkbookSerializerCommon.nvl(hf.getCenterHeader());
        String rh = XlsxWorkbookSerializerCommon.nvl(hf.getRightHeader());
        String lf = XlsxWorkbookSerializerCommon.nvl(hf.getLeftFooter());
        String cf = XlsxWorkbookSerializerCommon.nvl(hf.getCenterFooter());
        String rf = XlsxWorkbookSerializerCommon.nvl(hf.getRightFooter());
        if (!lh.isEmpty() || !ch.isEmpty() || !rh.isEmpty() || !lf.isEmpty() || !cf.isEmpty() || !rf.isEmpty()) {
            sb.append("<headerFooter>");
            if (!lh.isEmpty() || !ch.isEmpty() || !rh.isEmpty())
                sb.append("<oddHeader>").append(XlsxWorkbookSerializerCommon.xmlText(
                    XlsxWorkbookSerializerCommon.buildHF(lh, ch, rh))).append("</oddHeader>");
            if (!lf.isEmpty() || !cf.isEmpty() || !rf.isEmpty())
                sb.append("<oddFooter>").append(XlsxWorkbookSerializerCommon.xmlText(
                    XlsxWorkbookSerializerCommon.buildHF(lf, cf, rf))).append("</oddFooter>");
            sb.append("</headerFooter>");
        }

        // rowBreaks
        if (!ps.getHorizontalPageBreaks().isEmpty()) {
            List<Integer> breaks = ps.getHorizontalPageBreaks();
            sb.append("<rowBreaks count=\"").append(breaks.size()).append("\" manualBreakCount=\"").append(breaks.size()).append("\">");
            for (int brk : breaks) sb.append("<brk id=\"").append(brk + 1).append("\" max=\"16383\" man=\"1\"/>");
            sb.append("</rowBreaks>");
        }

        // colBreaks
        if (!ps.getVerticalPageBreaks().isEmpty()) {
            List<Integer> breaks = ps.getVerticalPageBreaks();
            sb.append("<colBreaks count=\"").append(breaks.size()).append("\" manualBreakCount=\"").append(breaks.size()).append("\">");
            for (int brk : breaks) sb.append("<brk id=\"").append(brk + 1).append("\" max=\"1048575\" man=\"1\"/>");
            sb.append("</colBreaks>");
        }
    }

    // =========================================================================
    // Load
    // =========================================================================

    /**
     * Loads the page setup.
     * @param ws ws
     * @param doc doc
     */
    static void loadPageSetup(WorksheetModel ws, Document doc) {
        // printOptions
        NodeList poNodes = doc.getElementsByTagNameNS("*", "printOptions");
        if (poNodes.getLength() > 0) {
            Element po = (Element) poNodes.item(0);
            PrintOptionsModel pom = ws.getPageSetup().getPrintOptions();
            if ("1".equals(po.getAttribute("gridLines"))) pom.setGridLines(true);
            if ("1".equals(po.getAttribute("headings"))) pom.setHeadings(true);
            if ("1".equals(po.getAttribute("horizontalCentered"))) pom.setHorizontalCentered(true);
            if ("1".equals(po.getAttribute("verticalCentered"))) pom.setVerticalCentered(true);
        }

        // pageMargins
        NodeList pmNodes = doc.getElementsByTagNameNS("*", "pageMargins");
        if (pmNodes.getLength() > 0) {
            Element pm = (Element) pmNodes.item(0);
            PageMarginsModel pmm = ws.getPageSetup().getMargins();
            double defMargin = 0.7;
            pmm.setLeft(XlsxWorkbookSerializerCommon.parseDouble(pm.getAttribute("left"), defMargin));
            pmm.setRight(XlsxWorkbookSerializerCommon.parseDouble(pm.getAttribute("right"), defMargin));
            pmm.setTop(XlsxWorkbookSerializerCommon.parseDouble(pm.getAttribute("top"), defMargin));
            pmm.setBottom(XlsxWorkbookSerializerCommon.parseDouble(pm.getAttribute("bottom"), defMargin));
            pmm.setHeader(XlsxWorkbookSerializerCommon.parseDouble(pm.getAttribute("header"), 0.3));
            pmm.setFooter(XlsxWorkbookSerializerCommon.parseDouble(pm.getAttribute("footer"), 0.3));
        }

        // pageSetup
        NodeList psNodes = doc.getElementsByTagNameNS("*", "pageSetup");
        if (psNodes.getLength() > 0) {
            Element ps = (Element) psNodes.item(0);
            PageSetupModel psm = ws.getPageSetup();
            String paperSize = ps.getAttribute("paperSize");
            if (!paperSize.isEmpty()) {
                int code = Integer.parseInt(paperSize);
                for (int i = 0; i < PaperSizeType.values().length; i++) {
                    if (PaperSizeType.values()[i].getValue() == code) { psm.setPaperSize(i); break; }
                }
            }
            String orient = ps.getAttribute("orientation");
            if ("landscape".equals(orient)) psm.setOrientation(PageOrientation.LANDSCAPE);
            else if ("portrait".equals(orient)) psm.setOrientation(PageOrientation.PORTRAIT);
            String scale = ps.getAttribute("scale");
            if (!scale.isEmpty()) psm.setScale(Integer.parseInt(scale));
            String fitW = ps.getAttribute("fitToWidth");
            if (!fitW.isEmpty()) psm.setFitToWidth(Integer.parseInt(fitW));
            String fitH = ps.getAttribute("fitToHeight");
            if (!fitH.isEmpty()) psm.setFitToHeight(Integer.parseInt(fitH));
            String fpn = ps.getAttribute("firstPageNumber");
            if (!fpn.isEmpty() && "1".equals(ps.getAttribute("useFirstPageNumber")))
                psm.setFirstPageNumber(Integer.parseInt(fpn));
        }

        // headerFooter
        NodeList hfNodes = doc.getElementsByTagNameNS("*", "headerFooter");
        if (hfNodes.getLength() > 0) {
            Element hf = (Element) hfNodes.item(0);
            HeaderFooterModel hfm = ws.getPageSetup().getHeaderFooter();
            NodeList oh = hf.getElementsByTagNameNS("*", "oddHeader");
            NodeList of = hf.getElementsByTagNameNS("*", "oddFooter");
            if (oh.getLength() > 0) XlsxWorkbookSerializerCommon.parseHF(oh.item(0).getTextContent(), hfm, true);
            if (of.getLength() > 0) XlsxWorkbookSerializerCommon.parseHF(of.item(0).getTextContent(), hfm, false);
        }

        // rowBreaks
        NodeList rbNodes = doc.getElementsByTagNameNS("*", "rowBreaks");
        if (rbNodes.getLength() > 0) {
            NodeList brkNodes = ((Element) rbNodes.item(0)).getElementsByTagNameNS("*", "brk");
            for (int i = 0; i < brkNodes.getLength(); i++) {
                Element brk = (Element) brkNodes.item(i);
                if ("1".equals(brk.getAttribute("man")))
                    ws.getPageSetup().getHorizontalPageBreaks().add(
                        XlsxWorkbookSerializerCommon.parseInt(brk.getAttribute("id"), 0) - 1);
            }
        }

        // colBreaks
        NodeList cbNodes = doc.getElementsByTagNameNS("*", "colBreaks");
        if (cbNodes.getLength() > 0) {
            NodeList brkNodes = ((Element) cbNodes.item(0)).getElementsByTagNameNS("*", "brk");
            for (int i = 0; i < brkNodes.getLength(); i++) {
                Element brk = (Element) brkNodes.item(i);
                if ("1".equals(brk.getAttribute("man")))
                    ws.getPageSetup().getVerticalPageBreaks().add(
                        XlsxWorkbookSerializerCommon.parseInt(brk.getAttribute("id"), 0) - 1);
            }
        }
    }
}
