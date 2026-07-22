package org.aspose.cells_foss;

import org.aspose.cells_foss.core.BordersValue;

/**
 * Represents the border properties for a cell or range in an Excel worksheet.
 */
public class Borders {
    private Border left = new Border();
    private Border right = new Border();
    private Border top = new Border();
    private Border bottom = new Border();
    private Border diagonal = new Border();
    private boolean diagonalUp;
    private boolean diagonalDown;

    /**
     * Returns the left.
     * @return the requested result
     */
    public Border getLeft() { return left; }
    /**
     * Sets the left.
     * @param left left
     */
    public void setLeft(Border left) { this.left = left; }

    /**
     * Returns the right.
     * @return the requested result
     */
    public Border getRight() { return right; }
    /**
     * Sets the right.
     * @param right right
     */
    public void setRight(Border right) { this.right = right; }

    /**
     * Returns the top.
     * @return the requested result
     */
    public Border getTop() { return top; }
    /**
     * Sets the top.
     * @param top top
     */
    public void setTop(Border top) { this.top = top; }

    /**
     * Returns the bottom.
     * @return the requested result
     */
    public Border getBottom() { return bottom; }
    /**
     * Sets the bottom.
     * @param bottom bottom
     */
    public void setBottom(Border bottom) { this.bottom = bottom; }

    /**
     * Returns the diagonal.
     * @return the requested result
     */
    public Border getDiagonal() { return diagonal; }
    /**
     * Sets the diagonal.
     * @param diagonal diagonal
     */
    public void setDiagonal(Border diagonal) { this.diagonal = diagonal; }

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

    // Package-internal model conversion
    /**
     * Creates an API object from the backing model.
     * @param bv bv
     * @return the computed result
     */
    static Borders fromModel(BordersValue bv) {
        Borders b = new Borders();
        // Handle the relevant branch before the state changes.
        if (bv.getLeft() != null)     b.left     = Border.fromModel(bv.getLeft());
        if (bv.getRight() != null)    b.right    = Border.fromModel(bv.getRight());
        if (bv.getTop() != null)      b.top      = Border.fromModel(bv.getTop());
        if (bv.getBottom() != null)   b.bottom   = Border.fromModel(bv.getBottom());
        if (bv.getDiagonal() != null) b.diagonal = Border.fromModel(bv.getDiagonal());
        b.diagonalUp   = bv.getDiagonalUp();
        b.diagonalDown = bv.getDiagonalDown();
        return b;
    }

    /**
     * Converts this instance to the backing model representation.
     * @return the computed result
     */
    BordersValue toModel() {
        BordersValue bv = new BordersValue();
        bv.setLeft(left.toModel());
        bv.setRight(right.toModel());
        bv.setTop(top.toModel());
        bv.setBottom(bottom.toModel());
        bv.setDiagonal(diagonal.toModel());
        bv.setDiagonalUp(diagonalUp);
        bv.setDiagonalDown(diagonalDown);
        return bv;
    }

    /**
     * Creates a deep clone of this Borders instance.
     *
     * @return a new Borders object with cloned border properties
     */
    public Borders clone() {
        Borders cloned = new Borders();
        cloned.setLeft(this.left.clone());
        cloned.setRight(this.right.clone());
        cloned.setTop(this.top.clone());
        cloned.setBottom(this.bottom.clone());
        cloned.setDiagonal(this.diagonal.clone());
        cloned.setDiagonalUp(this.diagonalUp);
        cloned.setDiagonalDown(this.diagonalDown);
        return cloned;
    }
}
