package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ShapeModel;

/** Represents a drawing object (auto shape) anchored to a worksheet. */
public final class Shape {
    private final ShapeModel model;

    Shape(ShapeModel model) { this.model = model; }

    ShapeModel getModel() { return model; }

    /** Gets or sets the display name of the shape. */
    public String getName() { return model.getName(); }
    public void setName(String value) { model.setName(value != null ? value : ""); }

    /** Zero-based row index of the upper-left anchor cell. */
    public int getUpperLeftRow() { return model.getUpperLeftRow(); }
    public void setUpperLeftRow(int value) {
        if (value < 0) throw new CellsException("UpperLeftRow must be non-negative.");
        model.setUpperLeftRow(value);
    }

    /** Zero-based column index of the upper-left anchor cell. */
    public int getUpperLeftColumn() { return model.getUpperLeftColumn(); }
    public void setUpperLeftColumn(int value) {
        if (value < 0) throw new CellsException("UpperLeftColumn must be non-negative.");
        model.setUpperLeftColumn(value);
    }

    /** Zero-based row index of the lower-right anchor cell. */
    public int getLowerRightRow() { return model.getLowerRightRow(); }
    public void setLowerRightRow(int value) { model.setLowerRightRow(value); }

    /** Zero-based column index of the lower-right anchor cell. */
    public int getLowerRightColumn() { return model.getLowerRightColumn(); }
    public void setLowerRightColumn(int value) { model.setLowerRightColumn(value); }

    /** Raw DrawingML preset geometry string (e.g. "rect", "rightArrow"). */
    public String getGeometryType() { return model.getGeometryType(); }
    public void setGeometryType(String value) {
        model.setGeometryType(value == null || value.isEmpty() ? "rect" : value);
    }

    /** Shape type enum. Setting this also updates {@link #getGeometryType()}. */
    public AutoShapeType getAutoShapeType() {
        return geometryToAutoShapeType(model.getGeometryType());
    }
    public void setAutoShapeType(AutoShapeType value) {
        model.setGeometryType(autoShapeTypeToGeometry(value));
    }

    // -------------------------------------------------------------------------

    static String autoShapeTypeToGeometry(AutoShapeType type) {
        if (type == null) return "rect";
        switch (type) {
            case RECTANGLE:          return "rect";
            case ROUNDED_RECTANGLE:  return "roundRect";
            case ELLIPSE:            return "ellipse";
            case TRIANGLE:           return "triangle";
            case RIGHT_TRIANGLE:     return "rtTriangle";
            case DIAMOND:            return "diamond";
            case PENTAGON:           return "pentagon";
            case HEXAGON:            return "hexagon";
            case OCTAGON:            return "octagon";
            case PLUS:               return "plus";
            case CUBE:               return "cube";
            case CYLINDER:           return "can";
            case HEART:              return "heart";
            case LIGHTNING:          return "lightningBolt";
            case SUN:                return "sun";
            case MOON:               return "moon";
            case CLOUD:              return "cloud";
            case RIGHT_ARROW:        return "rightArrow";
            case LEFT_ARROW:         return "leftArrow";
            case UP_ARROW:           return "upArrow";
            case DOWN_ARROW:         return "downArrow";
            case LEFT_RIGHT_ARROW:   return "leftRightArrow";
            case UP_DOWN_ARROW:      return "upDownArrow";
            case STAR4_POINT:        return "star4";
            case STAR5_POINT:        return "star5";
            case STAR6_POINT:        return "star6";
            case STAR7_POINT:        return "star7";
            case STAR8_POINT:        return "star8";
            case STAR10_POINT:       return "star10";
            case STAR12_POINT:       return "star12";
            case STAR16_POINT:       return "star16";
            case STAR24_POINT:       return "star24";
            case STAR32_POINT:       return "star32";
            case TEXT_BOX:           return "rect";
            case MATH_PLUS:          return "mathPlus";
            case STRAIGHT_CONNECTOR: return "straightConnector1";
            case BENT_CONNECTOR:     return "bentConnector3";
            case CURVED_CONNECTOR:   return "curvedConnector3";
            default:                 return "rect";
        }
    }

    static AutoShapeType geometryToAutoShapeType(String prst) {
        if (prst == null || prst.isEmpty()) return AutoShapeType.UNKNOWN;
        switch (prst) {
            case "rect":                return AutoShapeType.RECTANGLE;
            case "roundRect":           return AutoShapeType.ROUNDED_RECTANGLE;
            case "ellipse":             return AutoShapeType.ELLIPSE;
            case "triangle":            return AutoShapeType.TRIANGLE;
            case "rtTriangle":          return AutoShapeType.RIGHT_TRIANGLE;
            case "diamond":             return AutoShapeType.DIAMOND;
            case "pentagon":            return AutoShapeType.PENTAGON;
            case "hexagon":             return AutoShapeType.HEXAGON;
            case "octagon":             return AutoShapeType.OCTAGON;
            case "plus":                return AutoShapeType.PLUS;
            case "cube":                return AutoShapeType.CUBE;
            case "can":                 return AutoShapeType.CYLINDER;
            case "heart":               return AutoShapeType.HEART;
            case "lightningBolt":       return AutoShapeType.LIGHTNING;
            case "sun":                 return AutoShapeType.SUN;
            case "moon":                return AutoShapeType.MOON;
            case "cloud":               return AutoShapeType.CLOUD;
            case "rightArrow":          return AutoShapeType.RIGHT_ARROW;
            case "leftArrow":           return AutoShapeType.LEFT_ARROW;
            case "upArrow":             return AutoShapeType.UP_ARROW;
            case "downArrow":           return AutoShapeType.DOWN_ARROW;
            case "leftRightArrow":      return AutoShapeType.LEFT_RIGHT_ARROW;
            case "upDownArrow":         return AutoShapeType.UP_DOWN_ARROW;
            case "star4":               return AutoShapeType.STAR4_POINT;
            case "star5":               return AutoShapeType.STAR5_POINT;
            case "star6":               return AutoShapeType.STAR6_POINT;
            case "star7":               return AutoShapeType.STAR7_POINT;
            case "star8":               return AutoShapeType.STAR8_POINT;
            case "star10":              return AutoShapeType.STAR10_POINT;
            case "star12":              return AutoShapeType.STAR12_POINT;
            case "star16":              return AutoShapeType.STAR16_POINT;
            case "star24":              return AutoShapeType.STAR24_POINT;
            case "star32":              return AutoShapeType.STAR32_POINT;
            case "mathPlus":            return AutoShapeType.MATH_PLUS;
            case "straightConnector1":  return AutoShapeType.STRAIGHT_CONNECTOR;
            case "bentConnector2":
            case "bentConnector3":
            case "bentConnector4":
            case "bentConnector5":      return AutoShapeType.BENT_CONNECTOR;
            case "curvedConnector2":
            case "curvedConnector3":
            case "curvedConnector4":
            case "curvedConnector5":    return AutoShapeType.CURVED_CONNECTOR;
            default:                    return AutoShapeType.UNKNOWN;
        }
    }
}
