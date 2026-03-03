package com.example.excel.controller;

import com.example.excel.model.ExcelData;
import com.example.excel.model.MergeResponse;
import com.example.excel.service.ExcelTemplateMergeService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Slf4j
@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class ExcelMergeController {

    private final ExcelTemplateMergeService mergeService;

    /**
     * Merge JSON data into Excel template
     */
    @PostMapping("/merge")
    public ResponseEntity<MergeResponse> mergeData(@RequestBody ExcelData data) {
        try {
            long startTime = System.currentTimeMillis();
            
            String outputFileName = generateOutputFileName();
            String outputPath = mergeService.mergeDataToTemplate(data, outputFileName);
            
            long processingTime = System.currentTimeMillis() - startTime;
            
            MergeResponse response = new MergeResponse(true, "Data merged successfully");
            response.setOutputPath(outputPath);
            response.setProcessingTimeMs(processingTime);
            
            return ResponseEntity.ok(response);
            
        } catch (Exception e) {
            log.error("Error merging data", e);
            MergeResponse errorResponse = new MergeResponse(false, "Error: " + e.getMessage());
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(errorResponse);
        }
    }

    /**
     * Debug endpoint to echo received data
     */
    @PostMapping("/debug/echo")
    public ResponseEntity<ExcelData> echoData(@RequestBody ExcelData data) {
        log.info("Received data: {}", data.getFields());
        return ResponseEntity.ok(data);
    }

    /**
     * Health check endpoint
     */
    @GetMapping("/health")
    public ResponseEntity<String> health() {
        return ResponseEntity.ok("Excel Template Merger is running");
    }

    /**
     * Generate output file name with timestamp
     */
    private String generateOutputFileName() {
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        return "merged_" + timestamp + ".xlsx";
    }

}
