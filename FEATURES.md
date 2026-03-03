# Excel Template Merger – Feature Summary

This document collects all of the capabilities that have been added to the
service during development.  It can be used as a quick reference or as
material for further user documentation.

## Core functionality

* **YAML‑based configuration** – mappings are defined in `application.yml` or
  a custom YAML file bound to `ExcelTemplateConfig` with `@ConfigurationProperties`.
* **JSON input model** – arbitrary JSON is deserialized into `ExcelData`, a
  flexible map that supports dot/bracket paths.
* **Apache POI integration** – reads a template workbook, writes the
  merged output, supports multiple sheets.
* **REST API** – exposed by `ExcelMergeController` with `/api/excel/merge` and
  health endpoint.

## Cell mapping features

* **Single‑cell mappings** – map a JSON field directly to a specific row/col.
* **Array fills** – support lists from JSON through three modes:
  * `SINGLE` (default) – write the whole list as string/number into one cell.
  * `HORIZONTAL` – spread values across columns.
  * `VERTICAL` – spread values down rows.
  * `RANGE` – fill a rectangular range.
* **Fill direction configuration** – `fillDirection` property on a mapping
  specifies the mode (`HORIZONTAL`, `VERTICAL`, `RANGE`, `SINGLE`).
* **Intelligent re‑indexing** – when preserving formula cells, array values
  are shifted automatically to avoid overwriting formulas.

## Formula handling

* **`preserveFormulas` flag** – when set on a mapping, the service will skip
  cells containing formulas rather than overwrite them.
* **Formula detection helper** – `isCellFormula` checks cell type / cached
  formula.

## Template and merged‑cell support

* **Merged‑cell awareness** – writing into the leading cell of a merged
  region correctly populates the whole span; additional logging added to
  track these writes.
* **Support for pre‑merged headers and value cells** – templates may contain
  merged regions built in Excel; the service simply writes to the first cell
  of each region.
* **Dynamic merging instructions** – the YAML and sample templates now
  demonstrate how to pre‑merge plan headers and value ranges.

## Dynamic layout features (benefit/plan example)

* **Benefit name location** – configuration specifies where the template
  already lists benefit names so values can be aligned by name instead of
  position.
* **Dynamic plan column generation** – when `dynamicPlanColumns` is true, the
  service ignores `fieldMappings` and instead builds a mapping list at
  runtime based on the JSON payload and template structure.  This removes the
  need to know the number of plans in advance.
* **Automatic plan spacing** – plan headers and their benefit columns are
  spaced automatically (one blank column between groups) during mapping.
* **Benefit name matching** – if a benefit value name matches an existing
  template row, that row index is used; otherwise a default sequential row
  is assigned.
* **Runtime plan path override** – new `planListJsonPath` property lets the
  caller specify where the array of plans lives (e.g. `plans` or
  `planlist` or `foo.bar.items`).

## Data path resolution

* Support for nested objects and arrays: 
  `ExcelData.getField("plans[0].benefits[1].value")`.
* Path resolver handles missing keys gracefully and returns `null` if the path
  cannot be resolved.
* Test coverage for various path patterns, including the new alternate root
  syntax.

## Testing and debugging

* **Unit tests** – verify path resolution, mapping generation, formula
  preservation logic, and dynamic configuration.
* **Integration scripts** – shell scripts (`test_merge.sh`,
  `test_benefits.sh`, `test_formula_preservation.sh`) run the application with
  sample data and templates, making it easy to exercise features manually.
* **Logging enhancements** – informative `INFO`/`DEBUG` logs for each mapping
  application and cell write; useful for troubleshooting merged cells or
  empty outputs.

## Examples & samples

* `sample_benefits.json` – JSON payload used in tests and scripts.
* `application_benefits.yml` – comprehensive example configuration showing
  both static mappings and the dynamic/plan features.
* Excel templates (`templates/benefit_plan.xlsx`) with merged cells,
  headers, and benefit names pre‑populated.

## Future ideas (already mentioned in README)

* programmatic merge-specific flags such as `mergeAcross`/`mergeDown`.
* image insertion, style preservation, spreadsheet formula support beyond
  skipping.
* Web UI for managing templates and configurations.


> This feature list is kept in sync with the codebase; refer to the README for
> usage instructions and to the source comments for implementation details.
