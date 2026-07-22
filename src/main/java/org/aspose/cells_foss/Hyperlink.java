package org.aspose.cells_foss;

import org.aspose.cells_foss.core.CellAddress;
import org.aspose.cells_foss.core.HyperlinkModel;
import java.util.List;

/**
 * Represents a hyperlink in a worksheet.
 */
public final class Hyperlink {
    private final List<HyperlinkModel> owner;
    private final HyperlinkModel model;

    /**
     * Initializes a new Hyperlink instance.
     * @param owner owner
     * @param model model
     */
    Hyperlink(List<HyperlinkModel> owner, HyperlinkModel model) {
        this.owner = owner;
        this.model = model;
    }

    /**
     * Gets the cell area covered by this hyperlink.
     */
    public String getArea() {
        String first = new CellAddress(model.getFirstRow(), model.getFirstColumn()).toString();
        // Handle the relevant branch before the state changes.
        if (model.getTotalRows() == 1 && model.getTotalColumns() == 1) {
            return first;
        }

        String last = new CellAddress(model.getFirstRow() + model.getTotalRows() - 1,
                                      model.getFirstColumn() + model.getTotalColumns() - 1).toString();
        return first + ":" + last;
    }

    /**
     * Gets or sets the address of the hyperlink.
     */
    public String getAddress() {
        String address = model.getAddress();
        // Handle the relevant branch before the state changes.
        if (address != null && !address.isEmpty()) {
            return address;
        }

        return model.getSubAddress() != null ? model.getSubAddress() : "";
    }

    /**
     * Sets the address.
     * @param value value to apply
     */
    public void setAddress(String value) {
        assignAddress(value);
    }

    /**
     * Gets the link type based on the address.
     */
    public TargetModeType getLinkType() {
        // Handle the relevant branch before the state changes.
        if (model.getSubAddress() != null && !model.getSubAddress().isEmpty()) {
            return TargetModeType.CELL_REFERENCE;
        }

        String address = model.getAddress();
        if (address == null || address.isEmpty()) {
            return TargetModeType.EXTERNAL;
        }

        if (address.toLowerCase().startsWith("mailto:")) {
            return TargetModeType.EMAIL;
        }

        if (address.startsWith("\\") || address.indexOf(":\\") > 0) {
            return TargetModeType.FILE_PATH;
        }

        return TargetModeType.EXTERNAL;
    }

    /**
     * Gets or sets the screen tip text.
     */
    public String getScreenTip() {
        return model.getScreenTip() != null ? model.getScreenTip() : "";
    }

    /**
     * Sets the screen tip.
     * @param value value to apply
     */
    public void setScreenTip(String value) {
        model.setScreenTip(normalizeText(value));
    }

    /**
     * Gets or sets the text displayed for the hyperlink.
     */
    public String getTextToDisplay() {
        return model.getTextToDisplay() != null ? model.getTextToDisplay() : "";
    }

    /**
     * Sets the text to display.
     * @param value value to apply
     */
    public void setTextToDisplay(String value) {
        model.setTextToDisplay(normalizeText(value));
    }

    /**
     * Deletes this hyperlink.
     */
    public void delete() {
        owner.remove(model);
    }

    /**
     * Assigns address.
     * @param value value to apply
     */
    private void assignAddress(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isBlank()) {
            model.setAddress(null);
            model.setSubAddress(null);
            return;
        }

        String normalized = value.trim();
        if (normalized.startsWith("#")) {
            model.setAddress(null);
            model.setSubAddress(normalized.substring(1));
            return;
        }

        if (normalized.indexOf('!') >= 0) {
            model.setAddress(null);
            model.setSubAddress(normalized);
            return;
        }

        model.setAddress(normalized);
        model.setSubAddress(null);
    }

    /**
     * Normalizes the text.
     * @param value value to apply
     * @return the computed result
     */
    private static String normalizeText(String value) {
        // Handle the relevant branch before the state changes.
        if (value == null || value.isEmpty()) {
            return null;
        }

        return value;
    }
}
