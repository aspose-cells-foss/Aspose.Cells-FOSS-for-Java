package com.aspose.cells_foss;

import com.aspose.cells_foss.core.PictureModel;

/** Represents an embedded image in a worksheet. */
public final class Picture {
    private final PictureModel model;

    Picture(PictureModel model) {
        this.model = model;
    }

    PictureModel getModel() { return model; }

    public String getName() { return model.getName(); }
    public void setName(String name) { model.setName(name != null ? name : ""); }

    public int getUpperLeftRow() { return model.getUpperLeftRow(); }
    public void setUpperLeftRow(int row) {
        if (row < 0) throw new CellsException("UpperLeftRow must be zero or greater.");
        model.setUpperLeftRow(row);
    }

    public int getUpperLeftColumn() { return model.getUpperLeftColumn(); }
    public void setUpperLeftColumn(int col) {
        if (col < 0) throw new CellsException("UpperLeftColumn must be zero or greater.");
        model.setUpperLeftColumn(col);
    }

    public int getLowerRightRow() { return model.getLowerRightRow(); }
    public void setLowerRightRow(int row) {
        if (row < 0) throw new CellsException("LowerRightRow must be zero or greater.");
        model.setLowerRightRow(row);
    }

    public int getLowerRightColumn() { return model.getLowerRightColumn(); }
    public void setLowerRightColumn(int col) {
        if (col < 0) throw new CellsException("LowerRightColumn must be zero or greater.");
        model.setLowerRightColumn(col);
    }

    /** Image type detected from the binary data signature. */
    public ImageType getImageType() {
        return PictureCollection.detectImageType(model.getData());
    }

    public byte[] getData() { return model.getData(); }
    public void setData(byte[] data) {
        if (data == null || data.length == 0)
            throw new CellsException("Image data must be non-empty.");
        model.setData(data);
        model.setExtension(PictureCollection.extensionFromData(data));
    }

    /** MIME content type derived from image data. */
    public String getContentType() {
        return PictureCollection.contentTypeFromExtension(model.getExtension());
    }
}
