package org.aspose.cells_foss.core;

/**
 * Internal model for an embedded picture/image.
 */
public final class PictureModel {
    private String name = "";
    private int upperLeftRow;
    private int upperLeftColumn;
    private long upperLeftRowOffset;
    private long upperLeftColumnOffset;
    private int lowerRightRow;
    private int lowerRightColumn;
    private long lowerRightRowOffset;
    private long lowerRightColumnOffset;
    private byte[] data;
    private String extension = "png";
    private String originalRelId;

    public String getName() { return name; }
    public void setName(String name) { this.name = name != null ? name : ""; }

    public int getUpperLeftRow() { return upperLeftRow; }
    public void setUpperLeftRow(int upperLeftRow) { this.upperLeftRow = upperLeftRow; }

    public int getUpperLeftColumn() { return upperLeftColumn; }
    public void setUpperLeftColumn(int upperLeftColumn) { this.upperLeftColumn = upperLeftColumn; }

    public long getUpperLeftRowOffset() { return upperLeftRowOffset; }
    public void setUpperLeftRowOffset(long upperLeftRowOffset) { this.upperLeftRowOffset = upperLeftRowOffset; }

    public long getUpperLeftColumnOffset() { return upperLeftColumnOffset; }
    public void setUpperLeftColumnOffset(long upperLeftColumnOffset) { this.upperLeftColumnOffset = upperLeftColumnOffset; }

    public int getLowerRightRow() { return lowerRightRow; }
    public void setLowerRightRow(int lowerRightRow) { this.lowerRightRow = lowerRightRow; }

    public int getLowerRightColumn() { return lowerRightColumn; }
    public void setLowerRightColumn(int lowerRightColumn) { this.lowerRightColumn = lowerRightColumn; }

    public long getLowerRightRowOffset() { return lowerRightRowOffset; }
    public void setLowerRightRowOffset(long lowerRightRowOffset) { this.lowerRightRowOffset = lowerRightRowOffset; }

    public long getLowerRightColumnOffset() { return lowerRightColumnOffset; }
    public void setLowerRightColumnOffset(long lowerRightColumnOffset) { this.lowerRightColumnOffset = lowerRightColumnOffset; }

    public byte[] getData() { return data; }
    public void setData(byte[] data) { this.data = data; }

    public String getExtension() { return extension; }
    public void setExtension(String extension) { this.extension = extension != null ? extension : "png"; }

    /** Original relationship ID from the loaded file, for round-trip renaming. */
    public String getOriginalRelId() { return originalRelId; }
    public void setOriginalRelId(String originalRelId) { this.originalRelId = originalRelId; }
}

