package org.aspose.cells_foss.core;

/**
 * Represents the border values for a cell style, including left, right, top, bottom,
 * and diagonal borders, as well as diagonal direction flags.
 */
public final class BordersValue {
    private BorderSideValue left = new BorderSideValue();
    private BorderSideValue right = new BorderSideValue();
    private BorderSideValue top = new BorderSideValue();
    private BorderSideValue bottom = new BorderSideValue();
    private BorderSideValue diagonal = new BorderSideValue();
    private boolean diagonalUp;
    private boolean diagonalDown;

    /**
     * Returns the left.
     * @return the requested result
     */
    public BorderSideValue getLeft() { return left; }
    /**
     * Sets the left.
     * @param left left
     */
    public void setLeft(BorderSideValue left) { this.left = left; }

    /**
     * Returns the right.
     * @return the requested result
     */
    public BorderSideValue getRight() { return right; }
    /**
     * Sets the right.
     * @param right right
     */
    public void setRight(BorderSideValue right) { this.right = right; }

    /**
     * Returns the top.
     * @return the requested result
     */
    public BorderSideValue getTop() { return top; }
    /**
     * Sets the top.
     * @param top top
     */
    public void setTop(BorderSideValue top) { this.top = top; }

    /**
     * Returns the bottom.
     * @return the requested result
     */
    public BorderSideValue getBottom() { return bottom; }
    /**
     * Sets the bottom.
     * @param bottom bottom
     */
    public void setBottom(BorderSideValue bottom) { this.bottom = bottom; }

    /**
     * Returns the diagonal.
     * @return the requested result
     */
    public BorderSideValue getDiagonal() { return diagonal; }
    /**
     * Sets the diagonal.
     * @param diagonal diagonal
     */
    public void setDiagonal(BorderSideValue diagonal) { this.diagonal = diagonal; }

    /**
     * Returns the diagonal up.
     * @return the requested result
     */
    public boolean getDiagonalUp() { return diagonalUp; }
    /**
     * Sets the diagonal up.
     * @param diagonalUp diagonal up
     */
    public void setDiagonalUp(boolean diagonalUp) { this.diagonalUp = diagonalUp; }

    /**
     * Returns the diagonal down.
     * @return the requested result
     */
    public boolean getDiagonalDown() { return diagonalDown; }
    /**
     * Sets the diagonal down.
     * @param diagonalDown diagonal down
     */
    public void setDiagonalDown(boolean diagonalDown) { this.diagonalDown = diagonalDown; }

    /**
     * Creates a copy of this instance.
     * @return the computed result
     */
    public BordersValue clone() {
        BordersValue cloned = new BordersValue();
        cloned.left = this.left.clone();
        cloned.right = this.right.clone();
        cloned.top = this.top.clone();
        cloned.bottom = this.bottom.clone();
        cloned.diagonal = this.diagonal.clone();
        cloned.diagonalUp = this.diagonalUp;
        cloned.diagonalDown = this.diagonalDown;
        return cloned;
    }
}
