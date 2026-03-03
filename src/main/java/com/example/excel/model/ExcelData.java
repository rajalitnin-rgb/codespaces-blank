package com.example.excel.model;

import com.fasterxml.jackson.annotation.JsonAnySetter;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.HashMap;
import java.util.Map;

@Data
@NoArgsConstructor
public class ExcelData {

    private Map<String, Object> fields = new HashMap<>();

    @JsonAnySetter
    public void setField(String key, Object value) {
        fields.put(key, value);
    }

    /**
     * Retrieve a value by path. Supports simple dot notation and array indices.
     * Examples:
     *   "firstName" -> top level value
     *   "plans[0].name" -> first element of plans list, property name
     *   "planlist[2].benefits[0].value" -> illustrates alternative root array
     */
    public Object getField(String path) {
        if (path == null) return null;
        String[] parts = path.split("\\.");
        Object current = fields; // start with the root map
        for (String part : parts) {
            if (current == null) {
                break;
            }
            // handle array index syntax on this segment
            if (part.contains("[")) {
                String name = part.substring(0, part.indexOf('['));
                int idx = Integer.parseInt(part.substring(part.indexOf('[') + 1, part.indexOf(']')));
                // first navigate to the named property
                if (current instanceof Map) {
                    current = ((Map<?, ?>) current).get(name);
                } else {
                    current = null;
                }
                // then extract from list if appropriate
                if (current instanceof java.util.List) {
                    java.util.List<?> list = (java.util.List<?>) current;
                    if (idx >= 0 && idx < list.size()) {
                        current = list.get(idx);
                    } else {
                        current = null;
                    }
                }
            } else {
                if (current instanceof Map) {
                    current = ((Map<?, ?>) current).get(part);
                } else {
                    current = null;
                }
            }
        }
        return current;
    }

}
