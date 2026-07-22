package org.aspose.cells_foss;

import org.aspose.cells_foss.core.WorkbookPropertiesModel;

/** Represents workbook-level protection settings (structure, windows, revision). */
public final class WorkbookProtection {
    private final WorkbookPropertiesModel.WorkbookProtectionModel model;

    WorkbookProtection(WorkbookPropertiesModel.WorkbookProtectionModel model) {
        this.model = model;
    }

    public boolean isProtected() { return model.hasStoredState(); }

    public boolean isLockStructure() { return model.getLockStructure(); }
    public void setLockStructure(boolean v) { model.setLockStructure(v); }

    public boolean isLockWindows() { return model.getLockWindows(); }
    public void setLockWindows(boolean v) { model.setLockWindows(v); }

    public boolean isLockRevision() { return model.getLockRevision(); }
    public void setLockRevision(boolean v) { model.setLockRevision(v); }

    public String getWorkbookPassword() { return model.getWorkbookPassword(); }
    public void setWorkbookPassword(String v) { model.setWorkbookPassword(v != null ? v : ""); }

    public String getRevisionsPassword() { return model.getRevisionsPassword(); }
    public void setRevisionsPassword(String v) { model.setRevisionsPassword(v != null ? v : ""); }
}

