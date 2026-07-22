package org.aspose.cells_foss;

import org.aspose.cells_foss.core.WorkbookModel;
import org.aspose.cells_foss.core.WorkbookPropertiesModel;

/** Represents the view / window settings stored in &lt;bookViews&gt;. */
public final class WorkbookView {
    private final WorkbookPropertiesModel.WorkbookViewModel model;
    private final WorkbookModel workbookModel;

    WorkbookView(WorkbookPropertiesModel.WorkbookViewModel model, WorkbookModel workbookModel) {
        this.model = model;
        this.workbookModel = workbookModel;
    }

    /** Zero-based index of the active worksheet (maps to bookView activeTab attribute). */
    public int getActiveTab() { return workbookModel.getActiveSheetIndex(); }
    public void setActiveTab(int v) { workbookModel.setActiveSheetIndex(v); }

    public Integer getXWindow() { return model.getXWindow(); }
    public void setXWindow(Integer v) { model.setXWindow(v); }

    public Integer getYWindow() { return model.getYWindow(); }
    public void setYWindow(Integer v) { model.setYWindow(v); }

    public Integer getWindowWidth() { return model.getWindowWidth(); }
    public void setWindowWidth(Integer v) { model.setWindowWidth(v); }

    public Integer getWindowHeight() { return model.getWindowHeight(); }
    public void setWindowHeight(Integer v) { model.setWindowHeight(v); }

    public Integer getFirstSheet() { return model.getFirstSheet(); }
    public void setFirstSheet(Integer v) { model.setFirstSheet(v); }

    public Boolean isShowHorizontalScroll() { return model.getShowHorizontalScroll(); }
    public void setShowHorizontalScroll(Boolean v) { model.setShowHorizontalScroll(v); }

    public Boolean isShowVerticalScroll() { return model.getShowVerticalScroll(); }
    public void setShowVerticalScroll(Boolean v) { model.setShowVerticalScroll(v); }

    public Boolean isShowSheetTabs() { return model.getShowSheetTabs(); }
    public void setShowSheetTabs(Boolean v) { model.setShowSheetTabs(v); }

    public Integer getTabRatio() { return model.getTabRatio(); }
    public void setTabRatio(Integer v) { model.setTabRatio(v); }

    public String getVisibility() { return model.getVisibility(); }
    public void setVisibility(String v) { model.setVisibility(v != null ? v : ""); }

    public boolean isMinimized() { return model.getMinimized(); }
    public void setMinimized(boolean v) { model.setMinimized(v); }

    public boolean isAutoFilterDateGrouping() { return model.getAutoFilterDateGrouping(); }
    public void setAutoFilterDateGrouping(boolean v) { model.setAutoFilterDateGrouping(v); }
}

