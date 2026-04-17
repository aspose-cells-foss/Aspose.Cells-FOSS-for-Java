# Agents Guide

This file is for coding agents working in this repository.

## Project Summary

- Project: Aspose.Cells FOSS for Java
- Language: Java 17
- Build tool: Maven
- Packaging: `jar`
- Primary scope: read, modify, and write Excel `.xlsx` workbooks
- Public API package: `com.aspose.cells_foss`

This is an XLSX-focused spreadsheet library with a public object model plus internal workbook, style, packaging, XML, and serializer layers.

## Repository Layout

- `pom.xml`
  Maven build definition, Java version, and test dependencies.
- `README.md`
  User-facing overview, supported scope, and API examples.
- `src/main/java/com/aspose/cells_foss/`
  Public API surface and XLSX orchestration entry points.
- `src/main/java/com/aspose/cells_foss/core/`
  Internal workbook, worksheet, style, validation, metadata, and supporting models.
- `src/main/java/com/aspose/cells_foss/packaging/`
  OOXML package and relationship abstractions.
- `src/main/java/com/aspose/cells_foss/xml/`
  XML mappers and parsing helpers.
- `src/main/java/com/aspose/cells_foss/validation/`
  Workbook validation helpers and validation messages.
- `src/test/java/com/aspose/cells_foss/`
  JUnit 5 scenario-driven tests.
- `target/`
  Build output. Do not edit generated contents directly.

## Build And Test

Use Maven from the repository root.

- Compile: `mvn compile`
- Run tests: `mvn test`
- Full build: `mvn clean package`

When validating a change, prefer the smallest useful command first:

- docs-only change: no build required unless you need to verify an example against source
- small API or model change: `mvn test`
- serializer, parser, packaging, or project-wide change: `mvn clean package`

## Implementation Boundaries

Keep responsibilities aligned with the current architecture:

- Public API classes in `com.aspose.cells_foss` should stay user-facing and ergonomic.
- Core model classes in `core` should remain implementation-detail containers.
- OOXML serialization/parsing logic belongs in `XlsxWorkbook*`, `packaging`, and `xml` classes.
- Validation-only helpers belong in `validation`.
- Tests should verify public behavior and file round-tripping, not private implementation details.

Do not move public behavior into `core` just to make a change faster.

## Current Product Scope

Document and preserve the current scope unless the task explicitly expands it:

- Load formats: `AUTO`, `XLSX`
- Save format: `XLSX`
- Formula values are stored and round-tripped, but this repository is not a full spreadsheet calculation engine
- Apache POI is used in tests, not as a runtime dependency
- Unsupported OOXML parts may be preserved across load/save flows when the current pipeline supports that path
- Load operations can emit diagnostics and repair metadata through `LoadDiagnostics`

Important current gap:

- Workbook-level wrapper APIs for protection, view, and calculation are not fully exposed yet; do not document or test them as complete public features unless you implement the missing surface

If you expand supported formats or workbook capabilities, update:

- `LoadFormat` / `SaveFormat`
- serialization/parsing code
- tests
- `README.md`
- `Agents.md`
- `RELEASE_NOTES.md` when the change is release-notable

## Coding Conventions

Match the existing style in surrounding files.

- Use plain Java 17 without Lombok or annotation-driven code generation.
- Prefer straightforward classes and explicit getters/setters.
- Preserve package names and current layering.
- Use `CellsException` or a more specific project exception for library-level failures.
- Validate user inputs early and fail with clear messages.
- Wrap low-level IO or parsing errors in project-specific exceptions where the surrounding code already does so.
- Use try-with-resources for streams and temporary resources.
- Keep comments factual and brief.

## API Change Rules

Be conservative with public API changes.

- Avoid renaming or removing public methods unless explicitly required.
- Keep overloads and enum values backward compatible when possible.
- If behavior changes, add or update tests that show the new public behavior.
- If a method is public in `com.aspose.cells_foss`, assume downstream users may depend on it.
- Keep README examples aligned with actual method names and current public return types.

## Testing Expectations

The test suite is organized by feature area. Representative classes include:

- `WorkbookTest`
- `CellValueTest`
- `StyleTest`
- `PageSetupTest`
- `AutoFilterTest`
- `ConditionalFormattingTest`
- `DataValidationTest`
- `HyperlinkTest`
- `DocumentPropertiesTest`
- `OutlineGroupTest`
- `CompatibilityTest`

Supporting helpers and scenario factories live alongside tests, including:

- `WorkbookScenarioFactory`
- `WorksheetScenarioFactory`
- `AutoFilterScenarioFactory`
- `PageSetupScenarioFactory`
- `TemporaryDirectory`

When changing a feature, extend the nearest existing test class instead of creating a new scattered test unless the feature area is genuinely new.

Keep test style consistent:

- JUnit 5
- descriptive scenario-style test names
- feature-grouped files
- use `TemporaryDirectory` and scenario factories where the repo already does so

## Common Change Patterns

For public API changes:

- update the public wrapper in `com.aspose.cells_foss`
- update the backing model if needed
- update XLSX load/save behavior if persistence is affected
- add or update round-trip tests
- refresh README examples or scope notes when user-facing behavior changes

For serializer or parser changes:

- check both write and read paths
- preserve unsupported parts when the existing load/save flow expects that
- verify generated XML remains consistent with the current OOXML conventions used in the project
- validate diagnostics and repair behavior when the load path changes

For style or formatting changes:

- update both model conversion and serializer logic
- add round-trip coverage in `StyleTest`, `ConditionalFormattingTest`, or related tests

For workbook or worksheet metadata changes:

- inspect `DocumentProperties`, `WorkbookProperties`, `WorkbookSettings`, worksheet view/protection wrappers, and the corresponding `XlsxWorkbook*` serializer code together

## Documentation Expectations

If you add a new notable feature or change supported behavior, update the relevant docs:

- `README.md` for user-facing capabilities, examples, and limitations
- code comments/Javadocs where public behavior changes materially
- `Agents.md` if workflow or repo guidance changes
- `RELEASE_NOTES.md` when the change is release-notable

Do not describe placeholder or partial APIs as complete features.

## Practical Warnings

- Some source comments currently contain imperfect generated phrasing. Do not assume comments are authoritative over code behavior.
- There are a few encoding artifacts in existing markdown/comments. Preserve file readability and avoid introducing new encoding issues.
- Do not edit `target/` outputs by hand.
- The worktree may contain unrelated local files or generated outputs; do not clean them up unless the user asked for it.

## Good Default Workflow

1. Read the relevant public API class.
2. Read the related `core` model and XLSX serializer/parser code.
3. Check the nearest existing test class and scenario factory.
4. Update the smallest coherent set of files.
5. Run `mvn test` or `mvn clean package` when code changes warrant it.
6. Update docs if user-facing behavior changed.
