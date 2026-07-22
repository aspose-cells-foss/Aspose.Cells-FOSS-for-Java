package org.aspose.cells_foss;

import org.aspose.cells_foss.core.PictureModel;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

/** Collection of embedded pictures on a worksheet. */
public final class PictureCollection {
    private final List<PictureModel> models;

    PictureCollection(List<PictureModel> models) {
        this.models = models;
    }

    public int getCount() { return models.size(); }

    public Picture get(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("Picture index out of range.");
        return new Picture(models.get(index));
    }

    /** Adds a picture from raw bytes. Returns the zero-based index of the new picture. */
    public int add(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn, byte[] imageData) {
        validateAnchors(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
        if (imageData == null || imageData.length == 0)
            throw new CellsException("Image data must be non-empty.");
        PictureModel model = buildModel(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, imageData);
        models.add(model);
        return models.size() - 1;
    }

    /** Adds a picture from an InputStream. Returns the zero-based index of the new picture. */
    public int add(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn, InputStream stream) {
        if (stream == null) throw new CellsException("Stream must not be null.");
        byte[] data;
        try {
            data = stream.readAllBytes();
        } catch (IOException e) {
            throw new CellsException("Failed to read image stream: " + e.getMessage());
        }
        return add(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, data);
    }

    /** Adds a picture from a file path. Returns the zero-based index of the new picture. */
    public int add(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn, String fileName) {
        if (fileName == null || fileName.isBlank()) throw new CellsException("File name must be non-empty.");
        byte[] data;
        try {
            data = Files.readAllBytes(Paths.get(fileName));
        } catch (IOException e) {
            throw new CellsException("Failed to read image file: " + e.getMessage());
        }
        return add(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, data);
    }

    public void removeAt(int index) {
        if (index < 0 || index >= models.size())
            throw new CellsException("Picture index out of range.");
        models.remove(index);
    }

    // -------------------------------------------------------------------------
    // Public helpers (also used by Picture.java)
    // -------------------------------------------------------------------------

    public static ImageType detectImageType(byte[] data) {
        if (data == null || data.length < 4) return ImageType.UNKNOWN;
        if (data[0] == (byte) 0xFF && data[1] == (byte) 0xD8) return ImageType.JPEG;
        if (data[0] == (byte) 0x89 && data[1] == 'P' && data[2] == 'N' && data[3] == 'G') return ImageType.PNG;
        if (data[0] == 'G' && data[1] == 'I' && data[2] == 'F') return ImageType.GIF;
        if (data[0] == 'B' && data[1] == 'M') return ImageType.BMP;
        return ImageType.UNKNOWN;
    }

    public static String extensionFromData(byte[] data) {
        return switch (detectImageType(data)) {
            case JPEG -> "jpeg";
            case PNG -> "png";
            case GIF -> "gif";
            case BMP -> "bmp";
            default -> "bin";
        };
    }

    public static String contentTypeFromExtension(String ext) {
        if (ext == null) return "application/octet-stream";
        return switch (ext.toLowerCase()) {
            case "jpeg", "jpg" -> "image/jpeg";
            case "png" -> "image/png";
            case "gif" -> "image/gif";
            case "bmp" -> "image/bmp";
            default -> "application/octet-stream";
        };
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    private static void validateAnchors(int ulr, int ulc, int lrr, int lrc) {
        if (ulr < 0) throw new CellsException("UpperLeftRow must be zero or greater.");
        if (ulc < 0) throw new CellsException("UpperLeftColumn must be zero or greater.");
        if (lrr < ulr) throw new CellsException("LowerRightRow must be >= UpperLeftRow.");
        if (lrc < ulc) throw new CellsException("LowerRightColumn must be >= UpperLeftColumn.");
    }

    private PictureModel buildModel(int ulr, int ulc, int lrr, int lrc, byte[] data) {
        PictureModel model = new PictureModel();
        model.setName("Picture " + (models.size() + 1));
        model.setUpperLeftRow(ulr);
        model.setUpperLeftColumn(ulc);
        model.setLowerRightRow(lrr);
        model.setLowerRightColumn(lrc);
        model.setData(data);
        model.setExtension(extensionFromData(data));
        return model;
    }
}

