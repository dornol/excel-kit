package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.dto.ProductReadDto;
import io.github.dornol.excelkit.example.app.common.DownloadFileType;
import io.github.dornol.excelkit.example.app.common.DownloadUtil;
import io.github.dornol.excelkit.excel.ExcelMapReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelSheetInfo;
import io.github.dornol.excelkit.shared.ExcelKitSchema;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/showcase")
public class ReadShowcaseController {

    private static final Logger log = LoggerFactory.getLogger(ReadShowcaseController.class);

    private static final ExcelKitSchema<ProductReadDto> PRODUCT_SCHEMA = ShowcaseData.PRODUCT_SCHEMA;

    // ========================================================================
    // 3. Schema - unified read/write with column config (name-based read) — write using schema
    // ========================================================================
    @GetMapping("/schema-excel")
    public ResponseEntity<StreamingResponseBody> downloadSchemaExcel() {
        var handler = PRODUCT_SCHEMA.excelWriter()
                .sheetName("Schema Demo")
                .autoFilter(true)
                .freezePane(1)
                .write(ShowcaseData.sampleProducts().stream().map(ShowcaseData::toReadDto));

        return DownloadUtil.builder("schema-excel-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 4. Name-based reading (upload endpoints)
    // ========================================================================
    @PostMapping("/read-by-name-excel")
    @ResponseBody
    public String readByNameExcel(MultipartFile file) throws IOException {
        try (InputStream is = file.getInputStream()) {
            return readAndFormat("Excel",
                    PRODUCT_SCHEMA.excelReader(ProductReadDto::new, null).build(is));
        }
    }

    @PostMapping("/read-by-name-csv")
    @ResponseBody
    public String readByNameCsv(MultipartFile file) throws IOException {
        try (InputStream is = file.getInputStream()) {
            return readAndFormat("CSV",
                    PRODUCT_SCHEMA.csvReader(ProductReadDto::new, null).build(is));
        }
    }

    private String readAndFormat(String type, io.github.dornol.excelkit.shared.AbstractReadHandler<ProductReadDto> handler) {
        List<ProductReadDto> results = new ArrayList<>();
        List<String> errors = new ArrayList<>();

        handler.read(result -> {
            if (result.success()) {
                results.add(result.data());
            } else {
                errors.add(result.messages().toString());
            }
        });

        log.info("Read by name ({}): {} success, {} errors", type, results.size(), errors.size());

        StringBuilder sb = new StringBuilder();
        sb.append("=== Name-Based %s Read Result ===\n".formatted(type));
        sb.append("Success: %d rows, Errors: %d rows\n\n".formatted(results.size(), errors.size()));
        results.forEach(p -> sb.append(p).append("\n"));
        if (!errors.isEmpty()) {
            sb.append("\n--- Errors ---\n");
            errors.forEach(e -> sb.append(e).append("\n"));
        }
        return sb.toString();
    }

    // ========================================================================
    // 10. columnAt - index-based reading
    // ========================================================================
    @PostMapping("/read-by-index")
    @ResponseBody
    public String readByIndex(MultipartFile file) throws IOException {
        List<ProductReadDto> results = new ArrayList<>();

        try (InputStream is = file.getInputStream()) {
            // Read only Name (col 0), Price (col 2), Discount (col 4) by index
            new ExcelReader<>(ProductReadDto::new, null)
                    .columnAt(0, (p, cell) -> p.setName(cell.asString()))
                    .columnAt(2, (p, cell) -> p.setPrice(cell.asInt()))
                    .columnAt(4, (p, cell) -> p.setDiscount(cell.asDouble()))
                    .build(is)
                    .read(result -> {
                        if (result.success()) results.add(result.data());
                    });
        }

        StringBuilder sb = new StringBuilder();
        sb.append("=== Index-Based Read Result ===\n");
        sb.append("Read %d rows (cols 0, 2, 4 only)\n\n".formatted(results.size()));
        results.forEach(p -> sb.append(p).append("\n"));

        log.info("Read by index: {} rows", results.size());
        return sb.toString();
    }

    // ========================================================================
    // 18. Map Reader - read into Map<String, String>
    // ========================================================================
    @PostMapping("/map-reader")
    @ResponseBody
    public String readMap(MultipartFile file) throws IOException {
        try (InputStream is = file.getInputStream()) {
            List<Map<String, String>> results = new ArrayList<>();
            new ExcelMapReader()
                    .build(is)
                    .read(r -> results.add(r.data()));

            StringBuilder sb = new StringBuilder();
            sb.append("=== Map-Based Read Result ===\n");
            sb.append("Read %d rows\n\n".formatted(results.size()));
            if (!results.isEmpty()) {
                sb.append("Headers: %s\n\n".formatted(results.get(0).keySet()));
            }
            results.forEach(row -> sb.append(row).append("\n"));
            return sb.toString();
        }
    }

    // ========================================================================
    // 19. Sheet Info - discover sheet names
    // ========================================================================
    @PostMapping("/sheet-info")
    @ResponseBody
    public String sheetInfo(MultipartFile file) throws IOException {
        byte[] data = file.getBytes();
        List<ExcelSheetInfo> sheets = ExcelReader.getSheetNames(new ByteArrayInputStream(data));

        StringBuilder sb = new StringBuilder();
        sb.append("=== Sheet Info ===\n\n");
        for (ExcelSheetInfo info : sheets) {
            List<String> headers = ExcelReader.getSheetHeaders(new ByteArrayInputStream(data), info.index(), 0);
            sb.append("Sheet %d: \"%s\" — Headers: %s\n".formatted(info.index(), info.name(), headers));
        }
        return sb.toString();
    }

    // ========================================================================
    // 25. Mapping Mode with Custom Conversion - CellData default methods
    // ========================================================================
    @PostMapping("/mapping-read")
    @ResponseBody
    public String readMappingMode(MultipartFile file) throws IOException {
        List<String> results = new ArrayList<>();
        List<String> errors = new ArrayList<>();

        try (InputStream is = file.getInputStream()) {
            ExcelReader.<String[]>mapping(row -> new String[]{
                    row.get("Name").asString("Unknown"),
                    String.valueOf(row.get("Price").asInt(0)),
                    String.valueOf(row.get("Quantity").asInt(0)),
                    String.valueOf(row.get("Discount").asDouble(0.0))
            }).build(is).read(result -> {
                if (result.success()) {
                    String[] data = result.data();
                    results.add("Name=%s, Price=%s, Qty=%s, Discount=%s".formatted(
                            data[0], data[1], data[2], data[3]));
                } else {
                    errors.add(result.messages().toString());
                }
            });
        }

        log.info("Mapping read: {} success, {} errors", results.size(), errors.size());

        StringBuilder sb = new StringBuilder();
        sb.append("=== Mapping Mode Read Result ===\n");
        sb.append("Success: %d rows, Errors: %d rows\n\n".formatted(results.size(), errors.size()));
        results.forEach(r -> sb.append(r).append("\n"));
        if (!errors.isEmpty()) {
            sb.append("\n--- Errors ---\n");
            errors.forEach(e -> sb.append(e).append("\n"));
        }
        return sb.toString();
    }
}
