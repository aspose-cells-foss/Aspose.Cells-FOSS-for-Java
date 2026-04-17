package com.aspose.cells_foss.core;

/**
 * Represents protection settings for a cell or range.
 */
public final class ProtectionValue {
    private boolean isLocked = true;
    private boolean isHidden;

    /**
     * Returns the locked.
     * @return the requested result
     */
    public boolean getIsLocked() { return isLocked; }
    /**
     * Sets the locked.
     * @param locked locked
     */
    public void setIsLocked(boolean locked) { isLocked = locked; }

    /**
     * Returns the hidden.
     * @return the requested result
     */
    public boolean getIsHidden() { return isHidden; }
    /**
     * Sets the hidden.
     * @param hidden hidden
     */
    public void setIsHidden(boolean hidden) { isHidden = hidden; }

    /**
     * Creates a copy of this instance.
     * @return the computed result
     */
    public ProtectionValue clone() {
        ProtectionValue cloned = new ProtectionValue();
        cloned.isLocked = this.isLocked;
        cloned.isHidden = this.isHidden;
        return cloned;
    }
}