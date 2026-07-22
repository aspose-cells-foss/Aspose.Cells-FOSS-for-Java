package org.aspose.cells_foss.core;

/**
 * Represents a repository for style values.
 */
public final class StyleRepository {

    /**
     * Normalizes a style value by returning a clone of it.
     *
     * @param style the style to normalize
     * @return a clone of the input style
     */
    public StyleValue normalize(StyleValue style) {
        return style.clone();
    }
}
