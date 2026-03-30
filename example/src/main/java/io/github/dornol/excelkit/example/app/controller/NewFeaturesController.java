package io.github.dornol.excelkit.example.app.controller;

import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.example.app.util.DownloadFileType;
import io.github.dornol.excelkit.example.app.util.DownloadUtil;
import io.github.dornol.excelkit.csv.CsvDialect;
import io.github.dornol.excelkit.csv.CsvQuoting;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelColor;
import io.github.dornol.excelkit.excel.ExcelConditionalRule;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.List;
import java.util.stream.Stream;

/**
 * Showcase controller for 0.9.2 new features:
 * data bar, icon set, CSV dialect, CSV quoting.
 */
@Controller
@RequestMapping("/showcase")
public class NewFeaturesController {

    private static List<ProductDto> sampleProducts() {
        return Stream.generate(ProductDto::random).limit(20).toList();
    }

    // ========================================================================
    // Data Bar / Icon Set
    // ========================================================================
    @GetMapping("/data-bar")
    public ResponseEntity<StreamingResponseBody> downloadDataBar() {
        var writer = new ExcelWriter<ProductDto>(ExcelColor.STEEL_BLUE)
                .sheetName("Data Bar Demo")
                .autoFilter(true)
                .freezePane(1);
        writer.addColumn("Name", ProductDto::name);
        writer.addColumn("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"));
        writer.addColumn("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER));
        writer.addColumn("Discount", ProductDto::discount, c -> c.type(ExcelDataType.DOUBLE_PERCENT));
        var handler = writer
                .conditionalFormatting(cf -> cf.columns(1).dataBar(ExcelColor.BLUE))
                .conditionalFormatting(cf -> cf.columns(2).dataBar(ExcelColor.RED, ExcelColor.GREEN))
                .conditionalFormatting(cf -> cf.columns(3).iconSet(ExcelConditionalRule.IconSetType.ARROWS_3))
                .write(sampleProducts().stream());

        return DownloadUtil.builder("data-bar-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // CSV Dialect — TSV preset
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
    // CSV Quoting — ALL strategy
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
