package com.example.excel.service;

import com.example.excel.config.ExcelTemplateConfig;
import com.example.excel.model.ExcelData;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.List;

@Slf4j
@Service
@RequiredArgsConstructor
public class ExcelTemplateMergeService {

    private final ExcelTemplateConfig config;

    /**
     * Merge JSON data into Excel template
     */
    public String mergeDataToTemplate(ExcelData data, String outputFileName) {
        long startTime = System.currentTimeMillis();
        
        try {
            // Load the template
            File templateFile = new File(config.getTemplatePath());
            if (!templateFile.exists()) {
                throw new IOException("Template file not found: " + config.getTemplatePath());
            }

            // Open workbook
            Workbook workbook = new XSSFWorkbook(new FileInputStream(templateFile));

            // Determine field mappings; either static list from YAML or
            // dynamically generated based on payload if the flag is set.
            List<ExcelTemplateConfig.FieldMapping> mappings = config.getFieldMappings();
            if (config.isDynamicPlanColumns()) {
                log.debug("Dynamic plan columns enabled, generating mappings from JSON");
                mappings = generateBenefitMappings(workbook, data, config);
            }

            if (mappings != null) {
                for (ExcelTemplateConfig.FieldMapping mapping : mappings) {
                    applyFieldMapping(workbook, data, mapping);
                }
            }

            // Save the modified workbook
            String outputPath = buildOutputPath(outputFileName);
            log.info("Saving workbook to: {}", outputPath);
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                workbook.write(fos);
                log.info("Workbook write completed");
            }
            log.info("Workbook closed and file saved");

            workbook.close();

            long processingTime = System.currentTimeMillis() - startTime;
            log.info("Successfully merged data into template. Output: {}, Time: {}ms", outputPath, processingTime);

            return outputPath;

        } catch (IOException e) {
            log.error("Error merging data to template", e);
            throw new RuntimeException("Failed to merge data to template: " + e.getMessage(), e);
        }
    }

    /**
     * Apply a single field mapping to the workbook
     */
    private void applyFieldMapping(Workbook workbook, ExcelData data,
                                   ExcelTemplateConfig.FieldMapping mapping) {
        Sheet sheet = workbook.getSheet(mapping.getSheetName());
        if (sheet == null) {
            log.warn("Sheet not found: {}", mapping.getSheetName());
            return;
        }

        Object value = data.getField(mapping.getJsonPath());
        
        String fillDir = mapping.getFillDirection() != null 
            ? mapping.getFillDirection().toUpperCase() 
            : "SINGLE";

        log.info("Applying mapping: jsonPath={}, fillDirection={}, value type={}", 
                mapping.getJsonPath(), fillDir, value != null ? value.getClass().getSimpleName() : "null");

        try {
            switch (fillDir) {
                case "HORIZONTAL":
                    fillHorizontal(sheet, value, mapping);
                    break;
                case "VERTICAL":
                    fillVertical(sheet, value, mapping);
                    break;
                case "RANGE":
                    fillRange(sheet, value, mapping);
                    break;
                case "SINGLE":
                default:
                    fillSingle(sheet, value, mapping);
                    break;
            }
        } catch (Exception e) {
            log.error("Error applying field mapping for: {}", mapping.getJsonPath(), e);
        }
    }

    /**
     * Fill a single cell
     */
    private void fillSingle(Sheet sheet, Object value, 
                           ExcelTemplateConfig.FieldMapping mapping) {
        Row row = sheet.getRow(mapping.getRowIndex());
        if (row == null) row = sheet.createRow(mapping.getRowIndex());
        
        Cell cell = row.getCell(mapping.getColumnIndex());
        if (cell == null) cell = row.createCell(mapping.getColumnIndex());
        
        // Check if cell contains formula and if we should preserve it
        if (isCellFormula(cell) && isPreserveFormulas(mapping)) {
            log.info("Skipping cell at ({},{}) - contains formula and preserveFormulas is enabled", 
                    mapping.getRowIndex(), mapping.getColumnIndex());
            return;
        }
        
        cell.setBlank();
        setCellValue(cell, value, mapping.getDataType());
        log.info("after write single cell at ({},{}): {}", mapping.getRowIndex(), mapping.getColumnIndex(), cell.toString());
        log.debug("Filled single cell at ({},{}): {}", mapping.getRowIndex(), mapping.getColumnIndex(), value);
    }

    /**
     * Fill array horizontally across columns with intelligent reindexing
     * If formulas are encountered and preserveFormulas is true, they are skipped
     * but data values are reindexed to fill the next available non-formula cells
     */
    private void fillHorizontal(Sheet sheet, Object value, 
                               ExcelTemplateConfig.FieldMapping mapping) {
        if (!(value instanceof java.util.List)) {
            log.warn("Expected list for HORIZONTAL fill, got: {}", 
                    value != null ? value.getClass().getName() : "null");
            if (value != null) fillSingle(sheet, value, mapping);
            return;
        }

        java.util.List<?> list = (java.util.List<?>) value;
        Row row = sheet.getRow(mapping.getRowIndex());
        if (row == null) row = sheet.createRow(mapping.getRowIndex());

        int dataIndex = 0;  // Tracks which data value to fill
        int cellOffset = 0; // Tracks cell position relative to starting column
        int formulasSkipped = 0;
        
        // Continue until all data values are placed
        while (dataIndex < list.size() && cellOffset < 1000) {  // 1000 is practical column limit
            int colIndex = mapping.getColumnIndex() + cellOffset;
            Cell cell = row.getCell(colIndex);
            if (cell == null) cell = row.createCell(colIndex);
            
            // Check if cell contains formula and if we should preserve it
            if (isCellFormula(cell) && isPreserveFormulas(mapping)) {
                log.debug("Formula cell at ({},{}) preserved; shifting to next position", 
                        mapping.getRowIndex(), colIndex);
                formulasSkipped++;
                cellOffset++;  // Move to next cell but don't consume data
                continue;
            }
            
            // Fill this cell with current data value
            cell.setBlank();
            setCellValue(cell, list.get(dataIndex), mapping.getDataType());
            dataIndex++;      // Move to next data value
            cellOffset++;     // Move to next cell
        }
        
        log.info("Filled {} data values horizontally at row {}, starting from col {} (skipped {} formula cells)", 
            list.size(), mapping.getRowIndex(), mapping.getColumnIndex(), formulasSkipped);
    }

    /**
     * Fill array vertically down rows with intelligent reindexing
     * If formulas are encountered and preserveFormulas is true, they are skipped
     * but data values are reindexed to fill the next available non-formula cells
     */
    private void fillVertical(Sheet sheet, Object value, 
                             ExcelTemplateConfig.FieldMapping mapping) {
        if (!(value instanceof java.util.List)) {
            log.warn("Expected list for VERTICAL fill, got: {}", 
                    value != null ? value.getClass().getName() : "null");
            if (value != null) fillSingle(sheet, value, mapping);
            return;
        }

        java.util.List<?> list = (java.util.List<?>) value;
        int dataIndex = 0;  // Tracks which data value to fill
        int rowOffset = 0;  // Tracks row position relative to starting row
        int formulasSkipped = 0;

        // Continue until all data values are placed
        while (dataIndex < list.size() && rowOffset < 1000) {  // 1000 is practical row limit
            int rowIndex = mapping.getRowIndex() + rowOffset;
            Row row = sheet.getRow(rowIndex);
            if (row == null) row = sheet.createRow(rowIndex);
            
            Cell cell = row.getCell(mapping.getColumnIndex());
            if (cell == null) cell = row.createCell(mapping.getColumnIndex());
            
            // Check if cell contains formula and if we should preserve it
            if (isCellFormula(cell) && isPreserveFormulas(mapping)) {
                log.debug("Formula cell at ({},{}) preserved; shifting to next position", 
                        rowIndex, mapping.getColumnIndex());
                formulasSkipped++;
                rowOffset++;  // Move to next row but don't consume data
                continue;
            }
            
            // Fill this cell with current data value
            cell.setBlank();
            setCellValue(cell, list.get(dataIndex), mapping.getDataType());
            dataIndex++;   // Move to next data value
            rowOffset++;   // Move to next row
        }
        
        log.info("Filled {} data values vertically at col {}, starting from row {} (skipped {} formula cells)", 
            list.size(), mapping.getColumnIndex(), mapping.getRowIndex(), formulasSkipped);
    }

    /**
     * Fill 2D array across a range with intelligent reindexing
     * If formulas are encountered and preserveFormulas is true, they are skipped
     * but data values are reindexed to fill the next available non-formula cells
     */
    private void fillRange(Sheet sheet, Object value, 
                          ExcelTemplateConfig.FieldMapping mapping) {
        if (!(value instanceof java.util.List)) {
            log.warn("Expected list for RANGE fill, got: {}", 
                    value != null ? value.getClass().getName() : "null");
            if (value != null) fillSingle(sheet, value, mapping);
            return;
        }

        java.util.List<?> rows = (java.util.List<?>) value;
        int dataRowIndex = 0;  // Tracks which row of data to fill
        int cellRowOffset = 0; // Tracks row position relative to starting row
        int totalFormulasSkipped = 0;

        // Continue until all data rows are placed
        while (dataRowIndex < rows.size() && cellRowOffset < 1000) {
            int currentRow = mapping.getRowIndex() + cellRowOffset;
            Row row = sheet.getRow(currentRow);
            if (row == null) row = sheet.createRow(currentRow);

            Object rowData = rows.get(dataRowIndex);
            int rowFormulasSkipped = 0;

            if (rowData instanceof java.util.List) {
                java.util.List<?> cols = (java.util.List<?>) rowData;
                int dataColIndex = 0;  // Tracks which data value to fill
                int cellColOffset = 0; // Tracks column position relative to starting column
                
                // Continue until all column data values are placed in this row
                while (dataColIndex < cols.size() && cellColOffset < 1000) {
                    int currentCol = mapping.getColumnIndex() + cellColOffset;
                    Cell cell = row.getCell(currentCol);
                    if (cell == null) cell = row.createCell(currentCol);
                    
                    // Check if cell contains formula and if we should preserve it
                    if (isCellFormula(cell) && isPreserveFormulas(mapping)) {
                        log.debug("Formula cell at ({},{}) preserved; shifting to next column", 
                                currentRow, currentCol);
                        rowFormulasSkipped++;
                        cellColOffset++;  // Move to next column but don't consume data
                        continue;
                    }
                    
                    // Fill this cell with current data value
                    cell.setBlank();
                    setCellValue(cell, cols.get(dataColIndex), mapping.getDataType());
                    dataColIndex++;    // Move to next data value
                    cellColOffset++;   // Move to next cell
                }
            } else {
                // Single value in row
                Cell cell = row.getCell(mapping.getColumnIndex());
                if (cell == null) cell = row.createCell(mapping.getColumnIndex());
                
                // Check if cell contains formula and if we should preserve it
                if (isCellFormula(cell) && isPreserveFormulas(mapping)) {
                    log.debug("Formula cell at ({},{}) preserved; skipping this data row", 
                            currentRow, mapping.getColumnIndex());
                    rowFormulasSkipped++;
                    cellRowOffset++;
                    continue;
                }
                
                cell.setBlank();
                setCellValue(cell, rowData, mapping.getDataType());
            }
            
            totalFormulasSkipped += rowFormulasSkipped;
            dataRowIndex++;    // Move to next data row
            cellRowOffset++;   // Move to next row
        }
        
        log.info("Filled RANGE with {} data rows starting at ({},{}); skipped {} total formula cells", 
            rows.size(), mapping.getRowIndex(), mapping.getColumnIndex(), totalFormulasSkipped);
    }

    /**
     * Check if a cell contains a formula
     */
    private boolean isCellFormula(Cell cell) {
        return cell != null && cell.getCellType() == CellType.FORMULA;
    }

    /**
     * Check if preserveFormulas flag is enabled for the mapping
     */
    private boolean isPreserveFormulas(ExcelTemplateConfig.FieldMapping mapping) {
        return mapping.getPreserveFormulas() != null && mapping.getPreserveFormulas();
    }

    /**
     * Set cell value based on data type - explicitly clear first to avoid caching issues
     */
    private void setCellValue(Cell cell, Object value, String dataType) {
        if (value == null) {
            cell.setBlank();
            return;
        }

        // Explicitly clear the cell to avoid POI's cached value issue
        cell.setBlank();

        String type = dataType != null ? dataType.toLowerCase() : "string";

        switch (type) {
            case "number":
            case "numeric":
                if (value instanceof Number) {
                    cell.setCellValue(((Number) value).doubleValue());
                } else {
                    try {
                        cell.setCellValue(Double.parseDouble(value.toString()));
                    } catch (NumberFormatException e) {
                        cell.setCellValue(value.toString());
                    }
                }
                break;

            case "date":
                if (value instanceof java.util.Date) {
                    cell.setCellValue((java.util.Date) value);
                } else if (value instanceof LocalDate) {
                    cell.setCellValue((LocalDate) value);
                } else {
                    cell.setCellValue(value.toString());
                }
                break;

            case "boolean":
                if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) value);
                } else {
                    cell.setCellValue(Boolean.parseBoolean(value.toString()));
                }
                break;

            case "string":
            default:
                cell.setCellValue(value.toString());
                break;
        }
    }

    /**
     * Build output file path
     */
    private String buildOutputPath(String outputFileName) {
        String outputDir = config.getOutputPath();
        File dir = new File(outputDir);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        return Paths.get(outputDir, outputFileName).toString();
    }

    /**
     * Generate benefit-plan mappings dynamically (see comments earlier in file).
     */
    public List<ExcelTemplateConfig.FieldMapping> generateBenefitMappings(Workbook workbook,
                                                                         ExcelData data,
                                                                         ExcelTemplateConfig config) {
        List<ExcelTemplateConfig.FieldMapping> mappings = new java.util.ArrayList<>();

        // read benefit name positions from template if configured
        java.util.Map<String,Integer> nameToRow = new java.util.HashMap<>();
        if (config.getBenefitNameLocation() != null) {
            ExcelTemplateConfig.BenefitNameLocation loc = config.getBenefitNameLocation();
            Sheet nameSheet = workbook.getSheet(loc.getSheetName());
            if (nameSheet != null) {
                int row = loc.getStartRow();
                while (true) {
                    Row r = nameSheet.getRow(row);
                    if (r == null) break;
                    Cell c = r.getCell(loc.getColumnIndex());
                    if (c == null || c.getCellType() == CellType.BLANK) break;
                    nameToRow.put(c.getStringCellValue(), row);
                    row++;
                    if (loc.getCount() != null && row >= loc.getStartRow() + loc.getCount()) break;
                }
            }
        }

        // allow the caller to override the name/path of the plans array
        String planPath = config.getPlanListJsonPath() != null
                ? config.getPlanListJsonPath()
                : "plans";          // defensive default

        Object planObj = data.getField(planPath);
        if (planObj instanceof java.util.List) {
            java.util.List<?> plans = (java.util.List<?>) planObj;
            int baseHeaderRow = 0;
            int baseHeaderCol = 1;    // column B
            int planSpacing = 3;      // leave one blank column between plan groups
            int defaultBenefitsStartRow = 1; // row 2
            for (int pi = 0; pi < plans.size(); pi++) {
                int headerCol = baseHeaderCol + pi * planSpacing;
                ExcelTemplateConfig.FieldMapping headerMap = new ExcelTemplateConfig.FieldMapping();
                headerMap.setJsonPath(planPath + "[" + pi + "].name");
                headerMap.setSheetName(config.getSheetNames().get(0));
                headerMap.setRowIndex(baseHeaderRow);
                headerMap.setColumnIndex(headerCol);
                headerMap.setDataType("string");
                mappings.add(headerMap);

                Object benefitObj = data.getField(planPath + "[" + pi + "].benefits");
                if (benefitObj instanceof java.util.List) {
                    java.util.List<?> benefits = (java.util.List<?>) benefitObj;
                    for (int bi = 0; bi < benefits.size(); bi++) {
                        Object benefit = benefits.get(bi);
                        String name = null;
                        if (benefit instanceof java.util.Map) {
                            Object o = ((java.util.Map<?,?>) benefit).get("name");
                            if (o != null) name = o.toString();
                        }
                        int rowIndex = defaultBenefitsStartRow + bi;
                        if (name != null && nameToRow.containsKey(name)) {
                            rowIndex = nameToRow.get(name);
                        }
                        ExcelTemplateConfig.FieldMapping bm = new ExcelTemplateConfig.FieldMapping();
                        bm.setJsonPath(planPath + "[" + pi + "].benefits[" + bi + "].value");
                        bm.setSheetName(config.getSheetNames().get(0));
                        bm.setRowIndex(rowIndex);
                        bm.setColumnIndex(headerCol + 1);
                        bm.setDataType("number");
                        mappings.add(bm);
                    }
                }
            }
        }
        return mappings;
    }
}
