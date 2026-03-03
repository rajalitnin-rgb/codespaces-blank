# Excel Template Merger

A Spring Boot application that merges JSON data into Excel templates using YAML-based configuration and Apache POI.

## Features

- **YAML Configuration**: Define field mappings and template paths in `application.yml`
- **JSON Input**: Accept JSON data via REST API
- **Apache POI Integration**: Leverage Apache POI for Excel manipulation
- **Multiple Data Types**: Support for string, number, date, and boolean data types
- **Multi-sheet Support**: Map data to different sheets in the Excel template

## Technologies

- **Java 17**: Latest LTS version
- **Spring Boot 3.2.3**: Modern Spring framework
- **Apache POI 5.0.0**: Excel file handling
- **Maven**: Build and dependency management
- **Jackson**: JSON processing with YAML support

## Project Structure

```
excel-template-merger/
├── src/
│   ├── main/
│   │   ├── java/com/example/excel/
│   │   │   ├── ExcelTemplateApplication.java    # Main application class
│   │   │   ├── config/
│   │   │   │   └── ExcelTemplateConfig.java     # YAML configuration binding
│   │   │   ├── model/
│   │   │   │   ├── ExcelData.java               # JSON data model
│   │   │   │   └── MergeResponse.java           # API response model
│   │   │   ├── service/
│   │   │   │   └── ExcelTemplateMergeService.java  # Merge logic
│   │   │   └── controller/
│   │   │       └── ExcelMergeController.java    # REST endpoints
│   │   └── resources/
│   │       ├── application.yml                  # Configuration
│   │       └── templates/                       # Excel templates directory
│   └── test/
│       └── java/com/example/excel/
│           └── ExcelTemplateApplicationTests.java
├── pom.xml                                      # Maven configuration
└── README.md                                    # This file
```

## Configuration

Edit `src/main/resources/application.yml` to configure:

### Template Settings
```yaml
excel:
  template:
    templatePath: ./templates/sample_template.xlsx
    outputPath: ./output
    sheetNames:
      - Sheet1
      - Sheet2
```

### Field Mappings
Define how JSON fields map to Excel cells:

```yaml
fieldMappings:
  - jsonPath: firstName           # JSON field name
    sheetName: Sheet1             # Target sheet
    rowIndex: 0                   # Row number (0-based)
    columnIndex: 0                # Column number (0-based)
    dataType: string              # Data type: string, number, date, boolean
    format: null                  # Optional formatting
```

### Dynamic plan/benefit layouts
You can enable runtime discovery of plans when the JSON payload contains a
variable number of entries.  Only the path to the plans array is configured;
individual objects are iterated so the YAML does **not** need to enumerate
each plan.

```yaml
excel:
  template:
    dynamicPlanColumns: true
    planListJsonPath: plans        # default; override if your array is
                                   # named something else (e.g. "planlist")
    benefitNameLocation:           # optional, for matching by name
      sheetName: Sheet1
      startRow: 1
      columnIndex: 0
```

## Usage

### 1. Prepare Template

Create an Excel template file at `templates/sample_template.xlsx` with the desired structure.

### 2. Configure Mappings

Update `application.yml` with your field mappings, specifying where each JSON field should be placed in the Excel template.

### 3. Build the Project

```bash
mvn clean install
```

### 4. Run the Application

```bash
mvn spring-boot:run
```

The application will start on `http://localhost:8080`

### 5. Merge Data via API

Send a POST request to `/api/excel/merge`:

```bash
curl -X POST http://localhost:8080/api/excel/merge \
  -H "Content-Type: application/json" \
  -d '{
    "firstName": "John",
    "lastName": "Doe",
    "salary": 75000,
    "joinDate": "2023-01-15",
    "isActive": true
  }'
```

### Response

```json
{
  "success": true,
  "message": "Data merged successfully",
  "outputPath": "./output/merged_20240303_143022.xlsx",
  "processingTimeMs": 245
}
```

### 6. Check Health

```bash
curl http://localhost:8080/api/excel/health
```

## Data Type Support

| Type | Description | Example |
|------|-------------|---------|
| `string` | Text data | "John Doe" |
| `number` | Numeric data | 75000 |
| `date` | Date/temporal data | "2023-01-15" |
| `boolean` | Boolean values | true/false |

## Error Handling

The application handles common errors gracefully:

- **Template not found**: Returns HTTP 500 with error message
- **Invalid sheet name**: Logs warning and skips mapping
- **Type conversion errors**: Converts to string as fallback

## Building and Running

### Build
```bash
mvn clean package
```

### Run JAR
```bash
java -jar target/excel-template-merger-1.0.0.jar
```

### Run with Maven
```bash
mvn spring-boot:run
```

## Development

### IDE Setup
This project works with any IDE supporting Maven:
- IntelliJ IDEA
- Visual Studio Code (with Maven extension)
- Eclipse

### Notable Dependencies
- **Spring Boot Starter Web**: REST API support
- **Apache POI**: Excel manipulation
- **Jackson**: JSON/YAML serialization
- **Lombok**: Reduce boilerplate code
- **SLF4J**: Logging

## Future Enhancements

- Support for Excel formulas
- Image insertion support
- Style and formatting preservation
- Batch processing of multiple records
- Database integration for template management
- Web UI for configuration management
- Support for more data types (lists, nested objects)

## License

Apache License 2.0

## Support

For issues or feature requests, please refer to the project documentation or create an issue in the repository.
