package com.aspose.cells_foss;

/**
 * Represents the paper size type for a worksheet.
 */
public enum PaperSizeType {
    /** The default paper size. */
    DEFAULT(0),
    /** Paper letter size (8.5 x 11 inches). */
    PAPER_LETTER(1),
    /** Paper tabloid size (11 x 17 inches). */
    PAPER_TABLOID(3),
    /** Paper legal size (8.5 x 14 inches). */
    PAPER_LEGAL(5),
    /** Paper statement size (5.5 x 8.5 inches). */
    PAPER_STATEMENT(6),
    /** Paper executive size (7.25 x 10.5 inches). */
    PAPER_EXECUTIVE(7),
    /** Paper A3 size (297 x 420 mm). */
    PAPER_A3(8),
    /** Paper A4 size (210 x 297 mm). */
    PAPER_A4(9),
    /** Paper A5 size (148 x 210 mm). */
    PAPER_A5(11),
    /** Paper B4 size (250 x 353 mm). */
    PAPER_B4(12),
    /** Paper B5 size (182 x 257 mm). */
    PAPER_B5(13);

    private final int value;

    /**
     * Initializes a new PaperSizeType instance.
     * @param value value to apply
     */
    PaperSizeType(int value) {
        this.value = value;
    }

    /**
     * Returns the value.
     * @return the requested result
     */
    public int getValue() {
        return value;
    }
}