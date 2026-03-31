package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.example.app.common.DownloadFileType;
import io.github.dornol.excelkit.example.app.common.DownloadUtil;
import io.github.dornol.excelkit.csv.CsvDialect;
import io.github.dornol.excelkit.csv.CsvMapReader;
import io.github.dornol.excelkit.csv.CsvQuoting;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.shared.ExcelKitSchema;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

@Controller
@RequestMapping("/showcase")
public class CsvShowcaseController {

    private static final ExcelKitSchema<io.github.dornol.excelkit.example.app.dto.ProductReadDto> PRODUCT_SCHEMA = ShowcaseData.PRODUCT_SCHEMA;

    private static List<ProductDto> sampleProducts() {
        return ShowcaseData.sampleProducts();
    }

    // ========================================================================
    // Schema CSV - write using schema
    // ========================================================================
    @GetMapping("/schema-csv")
    public ResponseEntity<StreamingResponseBody> downloadSchemaCsv() {
        var handler = PRODUCT_SCHEMA.csvWriter()
                .write(sampleProducts().stream().map(ShowcaseData::toReadDto));

        return DownloadUtil.builder("schema-csv-demo", DownloadFileType.CSV)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 26. CSV Injection Defense Toggle - defense OFF (formulas preserved)
    // ========================================================================
    @GetMapping("/csv-defense-off")
    public ResponseEntity<StreamingResponseBody> downloadCsvDefenseOff() {
        var handler = new CsvWriter<String[]>()
                .csvInjectionDefense(false)
                .column("Label", row -> row[0])
                .column("Value", row -> row[1])
                .write(Stream.of(
                        new String[]{"Sum Formula", "=SUM(A1:A10)"},
                        new String[]{"Phone Number", "+1-234-5678"},
                        new String[]{"Negative Number", "-15.5"},
                        new String[]{"At Symbol", "@importdata(...)"},
                        new String[]{"Normal Text", "Hello World"}
                ));

        return DownloadUtil.builder("csv-defense-off", DownloadFileType.CSV)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 27. CSV Injection Defense Toggle - defense ON (default)
    // ========================================================================
    @GetMapping("/csv-defense-on")
    public ResponseEntity<StreamingResponseBody> downloadCsvDefenseOn() {
        var handler = new CsvWriter<String[]>()
                .csvInjectionDefense(true)
                .column("Label", row -> row[0])
                .column("Value", row -> row[1])
                .write(Stream.of(
                        new String[]{"Sum Formula", "=SUM(A1:A10)"},
                        new String[]{"Phone Number", "+1-234-5678"},
                        new String[]{"Negative Number", "-15.5"},
                        new String[]{"At Symbol", "@importdata(...)"},
                        new String[]{"Normal Text", "Hello World"}
                ));

        return DownloadUtil.builder("csv-defense-on", DownloadFileType.CSV)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 28. CSV Map Reader - read CSV into Map<String, String>
    // ========================================================================
    @PostMapping("/csv-map-reader")
    @ResponseBody
    public String readCsvMap(MultipartFile file) throws IOException {
        try (InputStream is = file.getInputStream()) {
            List<Map<String, String>> results = new ArrayList<>();
            new CsvMapReader()
                    .build(is)
                    .read(r -> results.add(r.data()));

            StringBuilder sb = new StringBuilder();
            sb.append("=== CSV Map-Based Read Result ===\n");
            sb.append("Read %d rows\n\n".formatted(results.size()));
            if (!results.isEmpty()) {
                sb.append("Headers: %s\n\n".formatted(results.get(0).keySet()));
            }
            results.forEach(row -> sb.append(row).append("\n"));
            return sb.toString();
        }
    }

    // ========================================================================
    // CSV Dialect — TSV preset (from NewFeaturesController)
    // ========================================================================
    @GetMapping("/csv-dialect-tsv")
    public ResponseEntity<StreamingResponseBody> downloadCsvDialectTsv() {
        var handler = new CsvWriter<ProductDto>()
                .dialect(CsvDialect.TSV)
                .column("Name", (ProductDto p) -> p.name())
                .column("Category", (ProductDto p) -> p.category())
                .column("Price", (ProductDto p) -> String.valueOf(p.price()))
                .column("Quantity", (ProductDto p) -> String.valueOf(p.quantity()))
                .write(sampleProducts().stream());

        return DownloadUtil.builder("tsv-demo", DownloadFileType.CSV)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // CSV Quoting — ALL strategy (from NewFeaturesController)
    // ========================================================================
    @GetMapping("/csv-quoting-all")
    public ResponseEntity<StreamingResponseBody> downloadCsvQuotingAll() {
        var handler = new CsvWriter<ProductDto>()
                .quoting(CsvQuoting.ALL)
                .column("Name", (ProductDto p) -> p.name())
                .column("Category", (ProductDto p) -> p.category())
                .column("Price", (ProductDto p) -> String.valueOf(p.price()))
                .column("Quantity", (ProductDto p) -> String.valueOf(p.quantity()))
                .write(sampleProducts().stream());

        return DownloadUtil.builder("quoted-csv-demo", DownloadFileType.CSV)
                .body(handler::consumeOutputStream);
    }
}
