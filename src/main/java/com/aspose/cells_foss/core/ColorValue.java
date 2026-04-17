package com.aspose.cells_foss.core;

/**
 * Represents a color value with alpha, red, green, and blue components.
 */
public final class ColorValue {
    private final byte a;
    private final byte r;
    private final byte g;
    private final byte b;

    /**
     * Initializes a new ColorValue instance.
     * @param a a
     * @param r r
     * @param g g
     * @param b b
     */
    public ColorValue(byte a, byte r, byte g, byte b) {
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
        ColorValue other = (ColorValue) obj;
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
}