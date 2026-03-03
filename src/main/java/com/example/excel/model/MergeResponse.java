package com.example.excel.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class MergeResponse {

    private boolean success;
    private String message;
    private String outputPath;
    private long processingTimeMs;

    public MergeResponse(boolean success, String message) {
        this.success = success;
        this.message = message;
    }

}
