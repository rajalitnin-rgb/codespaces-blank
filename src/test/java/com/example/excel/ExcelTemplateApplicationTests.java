package com.example.excel;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class ExcelTemplateApplicationTests {

    @Test
    void contextLoads() {
    }

    @Test
    void benefitPathResolution() throws Exception {
        // load sample JSON and ensure nested/array paths resolve correctly
        java.nio.file.Path jsonPath = java.nio.file.Paths.get("sample_benefits.json");
        String json = java.nio.file.Files.readString(jsonPath);
        com.example.excel.model.ExcelData data = new com.fasterxml.jackson.databind.ObjectMapper()
                .readValue(json, com.example.excel.model.ExcelData.class);

        // top-level plan names
        org.junit.jupiter.api.Assertions.assertEquals("Platinum", data.getField("plans[0].name"));
        org.junit.jupiter.api.Assertions.assertEquals("Gold", data.getField("plans[1].name"));

        // benefit values
        org.junit.jupiter.api.Assertions.assertEquals(100, data.getField("plans[0].benefits[0].value"));
        org.junit.jupiter.api.Assertions.assertEquals(50, data.getField("plans[0].benefits[1].value"));
        org.junit.jupiter.api.Assertions.assertEquals(80, data.getField("plans[1].benefits[0].value"));
        org.junit.jupiter.api.Assertions.assertEquals(5, data.getField("plans[1].benefits[3].value"));
    }

    @Test
    void dynamicBenefitMappingsProduceSameResult() throws Exception {
        // load sample JSON
        java.nio.file.Path jsonPath = java.nio.file.Paths.get("sample_benefits.json");
        String json = java.nio.file.Files.readString(jsonPath);
        com.example.excel.model.ExcelData data = new com.fasterxml.jackson.databind.ObjectMapper()
                .readValue(json, com.example.excel.model.ExcelData.class);

        com.example.excel.config.ExcelTemplateConfig cfg = new com.example.excel.config.ExcelTemplateConfig();
        cfg.setSheetNames(java.util.Collections.singletonList("Sheet1"));
        cfg.setDynamicPlanColumns(true);

        com.example.excel.service.ExcelTemplateMergeService svc =
                new com.example.excel.service.ExcelTemplateMergeService(cfg);

        // open the template so the service can read the pre‑populated names
        java.util.List<com.example.excel.config.ExcelTemplateConfig.FieldMapping> auto;
        try (java.io.FileInputStream fis = new java.io.FileInputStream("templates/benefit_plan.xlsx")) {
            org.apache.poi.ss.usermodel.Workbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis);
            auto = svc.generateBenefitMappings(wb, data, cfg);
        }
        // we expect one header per plan plus 4 values each
        org.junit.jupiter.api.Assertions.assertEquals(2 * 5, auto.size());
        org.junit.jupiter.api.Assertions.assertEquals("plans[0].name", auto.get(0).getJsonPath());
        // column next to header is 2 (C)
        org.junit.jupiter.api.Assertions.assertEquals(2, auto.get(1).getColumnIndex());
        // benefit with name "Vision" should be mapped to the row where the
        // template already has "Vision" (row index 2)
        boolean found = auto.stream().anyMatch(fm -> fm.getJsonPath().equals("plans[0].benefits[1].value")
                && fm.getRowIndex() == 2);
        org.junit.jupiter.api.Assertions.assertTrue(found, "Vision value should go to predefined row");

        // dump mapping details for debugging
        System.out.println("generated mappings:");
        for (com.example.excel.config.ExcelTemplateConfig.FieldMapping fm : auto) {
            System.out.printf("  path=%s row=%d col=%d\n",
                fm.getJsonPath(), fm.getRowIndex(), fm.getColumnIndex());
        }

        // we expect one header per plan plus 4 values each
        org.junit.jupiter.api.Assertions.assertEquals(2 * 5, auto.size());
        // first header mapping should match explicitly constructed path
        org.junit.jupiter.api.Assertions.assertEquals("plans[0].name", auto.get(0).getJsonPath());
        // first benefit value should live one column to the right of the header
        org.junit.jupiter.api.Assertions.assertEquals(2, auto.get(1).getColumnIndex());

        // --- new scenario: ability to override the plans path ---
        String altJson = "{\"planlist\":[{\"name\":\"Solo\",\"benefits\":[{\"name\":\"Foo\",\"value\":5}]}]}";
        com.example.excel.model.ExcelData altData = new com.fasterxml.jackson.databind.ObjectMapper()
                .readValue(altJson, com.example.excel.model.ExcelData.class);
        cfg.setPlanListJsonPath("planlist");
        java.util.List<com.example.excel.config.ExcelTemplateConfig.FieldMapping> altAuto;
        try (java.io.FileInputStream fis = new java.io.FileInputStream("templates/benefit_plan.xlsx")) {
            org.apache.poi.ss.usermodel.Workbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis);
            altAuto = svc.generateBenefitMappings(wb, altData, cfg);
        }
        // path should reflect the overridden root
        org.junit.jupiter.api.Assertions.assertFalse(altAuto.isEmpty());
        org.junit.jupiter.api.Assertions.assertEquals("planlist[0].name", altAuto.get(0).getJsonPath());
    }

}
