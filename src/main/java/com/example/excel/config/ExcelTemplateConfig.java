package com.example.excel.config;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@ConfigurationProperties(prefix = "excel.template")
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class ExcelTemplateConfig {

    private String templatePath;
    private String outputPath;
    private List<String> sheetNames;
    
    /**
     * Optional static list of individual mappings.  When
     * {@link #dynamicPlanColumns} is true this list is ignored and mappings
     * are derived from the incoming JSON.
     */
    private List<FieldMapping> fieldMappings;

    /**
     * If true the service will create plan/benefit column mappings at runtime
     * rather than relying on a fixed YAML list.  Useful when the request JSON
     * contains a variable number of plans.
     */
    private boolean dynamicPlanColumns = false;

    /**
     * JSONPath (simple dot/bracket syntax) pointing at the array containing
     * plans.  The default is "plans" which matches the sample data, but you
     * can override it when your payload uses a different property name such
     * as `planlist` or nests the array under another object.
     *
     * Only the root of the plans array is configured here; individual plan
     * entries are discovered at runtime so you still don’t need to enumerate
     * every plan in YAML.
     */
    private String planListJsonPath = "plans";

    /**
     * Location of pre‑populated benefit names in the template.  When provided
     * the merge logic will read these cells and align values by name rather
     * than by position.
     */
    private BenefitNameLocation benefitNameLocation;

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    public static class BenefitNameLocation {
        private String sheetName;
        private Integer startRow;      // zero-based
        private Integer columnIndex;   // zero-based
        private Integer count;         // optional limit
    }

    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    public static class FieldMapping {
        private String jsonPath;
        private String sheetName;
        private Integer rowIndex;
        private Integer columnIndex;
        private String dataType;
        private String format;
        private String fillDirection;  // SINGLE (default), HORIZONTAL, VERTICAL, RANGE
        private Integer endRowIndex;   // For RANGE fill
        private Integer endColumnIndex; // For RANGE fill
        private Boolean preserveFormulas; // If true, skip cells containing formulas
    }

}
