package com.aspose.cells_foss;

import com.aspose.cells_foss.core.WorksheetModel;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

/**
 * Represents a collection of worksheets in a workbook.
 */
public class WorksheetCollection implements Iterable<Worksheet> {
    private final Workbook workbook;

    /**
     * Initializes a new WorksheetCollection instance.
     * @param workbook workbook to apply
     */
    WorksheetCollection(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * Returns the requested item.
     * @param index index
     * @return the requested result
     */
    public Worksheet get(int index) {
        return wrap(workbook.getModel().getWorksheets().get(index));
    }

    /**
     * Returns the requested item.
     * @param name name
     * @return the requested result
     */
    public Worksheet get(String name) {
        int index = indexOf(name);
        // Handle the relevant branch before the state changes.
        if (index < 0) {
            throw new CellsException("Worksheet '" + name + "' was not found.");
        }
        return wrap(workbook.getModel().getWorksheets().get(index));
    }

    /**
     * Returns the count.
     * @return the requested result
     */
    public int getCount() {
        return workbook.getModel().getWorksheets().size();
    }

    /**
     * Returns the active sheet index.
     * @return the requested result
     */
    public int getActiveSheetIndex() {
        return workbook.getModel().getActiveSheetIndex();
    }

    /**
     * Sets the active sheet index.
     * @param value value to apply
     */
    public void setActiveSheetIndex(int value) {
        // Handle the relevant branch before the state changes.
        if (value < 0 || value >= workbook.getModel().getWorksheets().size()) {
            throw new CellsException("ActiveSheetIndex must refer to an existing worksheet.");
        }
        workbook.getModel().setActiveSheetIndex(value);
    }

    /**
     * Returns the active sheet name.
     * @return the requested result
     */
    public String getActiveSheetName() {
        return workbook.getModel().getWorksheets().get(getActiveSheetIndex()).getName();
    }

    /**
     * Sets the active sheet name.
     * @param value value to apply
     */
    public void setActiveSheetName(String value) {
        int index = indexOf(value);
        // Handle the relevant branch before the state changes.
        if (index < 0) {
            throw new CellsException("Worksheet '" + value + "' was not found.");
        }
        setActiveSheetIndex(index);
    }

    /**
     * Adds a new item to the current collection.
     * @return the computed result
     */
    public int add() {
        return add(generateDefaultSheetName());
    }

    /**
     * Adds a new item to the current collection.
     * @param sheetName name to use
     * @return the computed result
     */
    public int add(String sheetName) {
        // Handle the relevant branch before the state changes.
        if (sheetName == null || sheetName.isBlank()) {
            throw new CellsException("Worksheet name must be non-empty.");
        }
        workbook.ensureUniqueSheetName(sheetName, null);
        workbook.getModel().getWorksheets().add(new WorksheetModel(sheetName));
        return workbook.getModel().getWorksheets().size() - 1;
    }

    /**
     * Removes at.
     * @param sheetName name to use
     */
    public void removeAt(String sheetName) {
        int index = indexOf(sheetName);
        // Handle the relevant branch before the state changes.
        if (index < 0) {
            throw new CellsException("Worksheet '" + sheetName + "' was not found.");
        }
        removeAt(index);
    }

    /**
     * Removes at.
     * @param index index
     */
    public void removeAt(int index) {
        // Handle the relevant branch before the state changes.
        if (workbook.getModel().getWorksheets().size() == 1) {
            throw new CellsException("A workbook must contain at least one worksheet.");
        }
        workbook.getModel().getWorksheets().remove(index);
        if (workbook.getModel().getActiveSheetIndex() > index) {
            workbook.getModel().setActiveSheetIndex(workbook.getModel().getActiveSheetIndex() - 1);
        } else if (workbook.getModel().getActiveSheetIndex() >= workbook.getModel().getWorksheets().size()) {
            workbook.getModel().setActiveSheetIndex(workbook.getModel().getWorksheets().size() - 1);
        }

        Integer firstSheet = workbook.getModel().getProperties().getView().getFirstSheet();
        if (firstSheet != null) {
            if (workbook.getModel().getWorksheets().size() == 0) {
                workbook.getModel().getProperties().getView().setFirstSheet(0);
            } else if (firstSheet >= workbook.getModel().getWorksheets().size()) {
                workbook.getModel().getProperties().getView().setFirstSheet(workbook.getModel().getWorksheets().size() - 1);
            }
        }
    }

    /**
     * Returns an iterator over the current collection.
     * @return the computed result
     */
    @Override
    public Iterator<Worksheet> iterator() {
        List<Worksheet> worksheets = new ArrayList<>(workbook.getModel().getWorksheets().size());
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < workbook.getModel().getWorksheets().size(); i++) {
            worksheets.add(wrap(workbook.getModel().getWorksheets().get(i)));
        }
        return worksheets.iterator();
    }

    /**
     * Processes index of.
     * @param name name
     * @return the computed result
     */
    private int indexOf(String name) {
        Objects.requireNonNull(name, "name");
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < workbook.getModel().getWorksheets().size(); i++) {
            WorksheetModel worksheet = workbook.getModel().getWorksheets().get(i);
            if (worksheet.getName().equalsIgnoreCase(name)) {
                return i;
            }
        }
        return -1;
    }

    /**
     * Generates default sheet name.
     * @return the computed result
     */
    private String generateDefaultSheetName() {
        int suffix = 1;
        // Walk the current collection so every entry is processed consistently.
        while (true) {
            String candidate = "Sheet" + suffix;
            if (indexOf(candidate) < 0) {
                return candidate;
            }
            suffix++;
        }
    }

    /**
     * Processes wrap.
     * @param worksheet worksheet to apply
     * @return the computed result
     */
    private Worksheet wrap(WorksheetModel worksheet) {
        return new Worksheet(workbook, worksheet);
    }
}