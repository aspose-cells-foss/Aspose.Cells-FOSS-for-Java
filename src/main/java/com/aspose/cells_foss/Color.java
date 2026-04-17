package com.aspose.cells_foss;

import com.aspose.cells_foss.core.ColorValue;

/**
 * Represents an ARGB color value.
 */
public final class Color {
    private final byte a;
    private final byte r;
    private final byte g;
    private final byte b;

    /**
     * Initializes a new Color instance.
     * @param a a
     * @param r r
     * @param g g
     * @param b b
     */
    public Color(byte a, byte r, byte g, byte b) {
        this.a = a;
        this.r = r;
        this.g = g;
        this.b = b;
    }

    /**
     * Returns the a.
     * @return the requested result
     */
    public byte getA() { return a; }
    /**
     * Returns the r.
     * @return the requested result
     */
    public byte getR() { return r; }
    /**
     * Returns the g.
     * @return the requested result
     */
    public byte getG() { return g; }
    /**
     * Returns the b.
     * @return the requested result
     */
    public byte getB() { return b; }

    /** Returns the empty (transparent black) color. */
    public static Color getEmpty() {
        return new Color((byte) 0, (byte) 0, (byte) 0, (byte) 0);
    }

    /**
     * Processes from argb.
     * @param a a
     * @param r r
     * @param g g
     * @param b b
     * @return the computed result
     */
    public static Color fromArgb(int a, int r, int g, int b) {
        return new Color((byte) a, (byte) r, (byte) g, (byte) b);
    }

    /**
     * Compares this instance with the provided value.
     * @param obj obj
     * @return true when the condition is satisfied
     */
    @Override
    public boolean equals(Object obj) {
        // Handle the relevant branch before the state changes.
        if (this == obj) return true;
        if (obj == null || getClass() != obj.getClass()) return false;
        Color other = (Color) obj;
        return a == other.a && r == other.r && g == other.g && b == other.b;
    }

    /**
     * Indicates whether this instance has h code.
     * @return true when the condition is satisfied
     */
    @Override
    public int hashCode() {
        int hash = Byte.toUnsignedInt(a);
        hash = (hash * 397) ^ Byte.toUnsignedInt(r);
        hash = (hash * 397) ^ Byte.toUnsignedInt(g);
        hash = (hash * 397) ^ Byte.toUnsignedInt(b);
        return hash;
    }

    // Package-internal conversion helpers (C# `internal`)
    /**
     * Converts this instance to the core model representation.
     * @return the computed result
     */
    ColorValue toCore() {
        return new ColorValue(a, r, g, b);
    }

    /**
     * Creates an API object from the core model representation.
     * @param value value to apply
     * @return the computed result
     */
    static Color fromCore(ColorValue value) {
        return new Color(value.getA(), value.getR(), value.getG(), value.getB());
    }
}