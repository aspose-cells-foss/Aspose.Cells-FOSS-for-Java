# Aspose.Cells FOSS for Java

Aspose.Cells FOSS for Java is a Java 17 spreadsheet library for creating, loading, modifying, and saving Excel `.xlsx` workbooks.
It exposes a user-facing API in `com.aspose.cells_foss` and keeps OOXML packaging, XML mapping, and XLSX serialization inside this repository.

## Installation

Add the following dependency to your `pom.xml`:

```xml
<dependency>
    <groupId>org.aspose</groupId>
    <artifactId>aspose-cells-foss</artifactId>
    <version>1.0.0</version>
</dependency>
```

## Current Scope

- Load formats: `AUTO`, `XLSX`
- Save format: `XLSX`
- Runtime target: Java 17
- Build tool: Maven
- Test framework: JUnit 5
- Apache POI is used in `test` scope only for compatibility checks

The library stores and round-trips formulas, but it is not a full spreadsheet calculation engine.

## Supported Features

- Workbook and worksheet creation, loading, saving, add/remove/rename, active sheet selection, and visibility
- Cell values for strings, numbers, booleans, date/time values, and formulas
- Cell formatting for fonts, borders, fills, alignment, number formats, and cell protection flags
- Row and column sizing, hiding, and outline grouping/collapsed state
- Worksheet view settings such as tab color, zoom, gridlines, row/column headers, zeros, and right-to-left mode
- Merged cells
- Defined names
- Hyperlinks
- Data validation
- Conditional formatting
- Auto filters with color filters, dynamic filters, top-10 filters, custom filters, and sort conditions
- Page setup
- Worksheet protection
- Freeze panes — freeze rows and/or columns by coordinate or cell name; inspect and clear frozen state
- **Cell comments** — add, edit, and remove comments (notes) with author, text, visibility, and size
- **Embedded pictures** — add images from bytes, streams, or file paths with two-cell anchor positioning; JPEG, PNG, GIF, and BMP detection
- **Embedded charts** — read chart name, type, anchor, and extent; 21 standard chart types can be created programmatically via `ChartCollection.add()`; loaded charts are preserved verbatim across load/save
- **Drawing shapes** — add, retrieve, and remove drawing objects with 35 preset geometries (`AutoShapeType`); complex shapes loaded from XLSX are preserved verbatim
- **Excel tables (ListObjects)** — create, resize, style, and remove structured tables with column definitions and totals rows
- **Workbook protection** — lock structure, windows, and revisions with separate passwords via `WorkbookProtection`
- **Workbook view** — read and write window position, size, active tab, scroll bar and sheet tab visibility via `WorkbookView`
- **Calculation properties** — configure calculation mode, reference mode, iterative calculation, and precision via `CalculationProperties`
- **Error cell values** — `#N/A`, `#VALUE!`, `#REF!` and other error results are loaded, stored, and saved with the correct `t="e"` cell type
- Document properties and workbook property persistence for supported `workbookPr` attributes
- `NumberFormat` utility for looking up and resolving built-in Excel format codes
- Load diagnostics, repair reporting, and preservation of unsupported package parts during load/save flows

## Known Limits

- Saving is currently limited to `.xlsx`
- ChartEx types (Waterfall, Treemap, Sunburst, Histogram, Box and Whisker, Funnel, Map) cannot be created programmatically; they are preserved verbatim across load/save
- Some APIs exist mainly to preserve OOXML metadata and package fidelity rather than to provide full Excel feature parity

## Build And Test

Requirements:

- JDK 17+
- Maven 3.9+

Common commands:

```bash
mvn compile
mvn test
mvn clean package
mvn javadoc:javadoc   # generates docs/apidocs/index.html
```

## Quick Start

Create a workbook, write values, style a cell, and save it:

```java
import com.aspose.cells_foss.Cell;
import com.aspose.cells_foss.Style;
import com.aspose.cells_foss.Workbook;
import com.aspose.cells_foss.Worksheet;

public class Main {
    public static void main(String[] args) {
        try (Workbook workbook = new Workbook()) {
            Worksheet sheet = workbook.getWorksheets().get(0);
            sheet.setName("Report");

            sheet.getCells().get("A1").putValue("Revenue");
            sheet.getCells().get("B1").putValue(12500.75);

            Cell total = sheet.getCells().get("B1");
            Style style = total.getStyle();
            style.getFont().setBold(true);
            style.setCustom("#,##0.00");
            total.setStyle(style);

            sheet.getCells().getRows().get(0).setHeight(22.0);
            sheet.getCells().getColumns().get(1).setWidth(14.5);

            workbook.save("report.xlsx");
        }
    }
}
```

Load an existing workbook with options and inspect diagnostics:

```java
import com.aspose.cells_foss.LoadIssue;
import com.aspose.cells_foss.LoadOptions;
import com.aspose.cells_foss.Workbook;

public class LoadWorkbook {
    public static void main(String[] args) {
        LoadOptions options = new LoadOptions();
        options.setStrictMode(false);
        options.setTryRepairPackage(true);
        options.setTryRepairXml(true);

        try (Workbook workbook = new Workbook("input.xlsx", options)) {
            if (workbook.getLoadDiagnostics().hasRepairs()) {
                for (LoadIssue issue : workbook.getLoadDiagnostics().getIssues()) {
                    System.out.println(issue.getMessage());
                }
            }

            workbook.getDocumentProperties().setAuthor("cells-foss");
            workbook.save("output.xlsx");
        }
    }
}
```

Add data validation and conditional formatting:

```java
import com.aspose.cells_foss.CellArea;
import com.aspose.cells_foss.FormatCondition;
import com.aspose.cells_foss.FormatConditionCollection;
import com.aspose.cells_foss.FormatConditionType;
import com.aspose.cells_foss.OperatorType;
import com.aspose.cells_foss.Style;
import com.aspose.cells_foss.Validation;
import com.aspose.cells_foss.ValidationType;
import com.aspose.cells_foss.Workbook;
import com.aspose.cells_foss.Worksheet;

public class RulesExample {
    public static void main(String[] args) {
        try (Workbook workbook = new Workbook()) {
            Worksheet sheet = workbook.getWorksheets().get(0);

            int validationIndex = sheet.getValidations().add(new CellArea(1, 0, 10, 1));
            Validation validation = sheet.getValidations().get(validationIndex);
            validation.setType(ValidationType.WHOLE_NUMBER);
            validation.setOperator(OperatorType.BETWEEN);
            validation.setFormula1("1");
            validation.setFormula2("100");

            int cfIndex = sheet.getConditionalFormattings().add();
            FormatConditionCollection conditions = sheet.getConditionalFormattings().get(cfIndex);
            conditions.addArea(CellArea.createCellArea("B2", "B11"));
            int conditionIndex = conditions.addCondition(
                    FormatConditionType.CELL_VALUE,
                    OperatorType.BETWEEN,
                    "1",
                    "100");

            FormatCondition condition = conditions.get(conditionIndex);
            Style style = condition.getStyle();
            style.getFont().setBold(true);
            condition.setStyle(style);

            workbook.save("rules.xlsx");
        }
    }
}
```

## Project Layout

```text
src/main/java/com/aspose/cells_foss/
  Public API surface: Workbook, Worksheet, Cell, Style, collections, enums, load/save options

src/main/java/com/aspose/cells_foss/core/
  Internal workbook, worksheet, style, validation, metadata, and packaging models

src/main/java/com/aspose/cells_foss/packaging/
  OOXML package and relationship abstractions

src/main/java/com/aspose/cells_foss/xml/
  XML mappers and parsing helpers

src/main/java/com/aspose/cells_foss/validation/
  Workbook validation messages and validation helpers

src/test/java/com/aspose/cells_foss/
  Scenario-driven integration tests for workbook behavior and XLSX round-tripping

src/test/java/com/aspose/cells_foss/unit/
  Fast, focused unit tests for individual API classes without XLSX I/O

docs/apidocs/
  Generated HTML Javadoc for the public API. Run `mvn javadoc:javadoc` to regenerate.
```

## Testing Coverage

Representative test areas include:

- workbook and worksheet behavior
- cell value handling
- style round-tripping and display text formatting
- page setup
- hyperlinks
- data validation
- conditional formatting
- auto filters
- document and workbook properties
- outline grouping
- cell comments, embedded pictures, embedded charts, drawing shapes, and Excel tables
- freeze panes, workbook protection, workbook view, and calculation properties
- error cell value round-tripping
- compatibility checks against generated XLSX output
- focused unit tests in `src/test/java/com/aspose/cells_foss/unit/`

## Notes For Contributors

- The main public API lives in `com.aspose.cells_foss`
- XLSX serialization and OOXML handling are implemented inside this repository
- If you change supported behavior, update `README.md`, `Agents.md`, and tests together

## License

This repository includes the MIT license at [`License/LICENSE.txt`](License/LICENSE.txt).

## Support

For bug reports, feature requests, and project questions, use the GitHub issue tracker:

- [Project issues](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Java/issues)
