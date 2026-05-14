package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ShapeModel;
import java.util.List;

/** Collection of drawing objects (shapes) on a worksheet. */
public final class ShapeCollection {
    private final List<ShapeModel> models;

    ShapeCollection(List<ShapeModel> models) { this.models = models; }

    public int getCount() { return models.size(); }

    public Shape get(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("Shape index " + index + " is out of range.");
        return new Shape(models.get(index));
    }

    /**
     * Adds a new shape and returns its zero-based index.
     *
     * @param upperLeftRow    zero-based row of the upper-left anchor cell
     * @param upperLeftColumn zero-based column of the upper-left anchor cell
     * @param lowerRightRow   zero-based row of the lower-right anchor cell
     * @param lowerRightColumn zero-based column of the lower-right anchor cell
     * @param shapeType       shape preset geometry type
     */
    public int add(int upperLeftRow, int upperLeftColumn,
                   int lowerRightRow, int lowerRightColumn,
                   AutoShapeType shapeType) {
        if (upperLeftRow < 0)    throw new CellsException("upperLeftRow must be non-negative.");
        if (upperLeftColumn < 0) throw new CellsException("upperLeftColumn must be non-negative.");
        if (lowerRightRow < upperLeftRow)       throw new CellsException("lowerRightRow must be >= upperLeftRow.");
        if (lowerRightColumn < upperLeftColumn) throw new CellsException("lowerRightColumn must be >= upperLeftColumn.");

        int index = models.size();
        ShapeModel model = new ShapeModel();
        model.setName("Shape " + (index + 1));
        model.setUpperLeftRow(upperLeftRow);
        model.setUpperLeftColumn(upperLeftColumn);
        model.setLowerRightRow(lowerRightRow);
        model.setLowerRightColumn(lowerRightColumn);
        model.setGeometryType(Shape.autoShapeTypeToGeometry(shapeType));
        models.add(model);
        return index;
    }

    public void removeAt(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("Shape index " + index + " is out of range.");
        models.remove(index);
    }
}
