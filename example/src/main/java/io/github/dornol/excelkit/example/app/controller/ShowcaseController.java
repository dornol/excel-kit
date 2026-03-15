package io.github.dornol.excelkit.example.app.controller;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.example.app.dto.ProductReadDto;
import io.github.dornol.excelkit.example.app.util.DownloadFileType;
import io.github.dornol.excelkit.example.app.util.DownloadUtil;
import io.github.dornol.excelkit.excel.*;
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

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

/**
 * Showcase controller demonstrating new and existing excel-kit features.
 * All endpoints use in-memory data (no DB required).
 */
@Controller
@RequestMapping("/showcase")
public class ShowcaseController {

    private static final Logger log = LoggerFactory.getLogger(ShowcaseController.class);

    private static List<ProductDto> sampleProducts() {
        return Stream.generate(ProductDto::random).limit(20).toList();
    }

    // ========================================================================
    // 1. Formula - SUM/AVERAGE in afterData callback
    // ========================================================================
    @GetMapping("/formula")
    public ResponseEntity<StreamingResponseBody> downloadFormula() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.STEEL_BLUE)
                .sheetName("Formula Demo")
                .autoFilter(true)
                .freezePane(1)
                .column("No.", (ProductDto row, io.github.dornol.excelkit.shared.Cursor cursor) -> cursor.getCurrentTotal()).type(ExcelDataType.LONG)
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price).type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                .column("Quantity", ProductDto::quantity).type(ExcelDataType.INTEGER)
                .column("Subtotal", (ProductDto row, io.github.dornol.excelkit.shared.Cursor cursor) ->
                        "D" + (cursor.getRowOfSheet() + 1) + "*E" + (cursor.getRowOfSheet() + 1))
                    .type(ExcelDataType.FORMULA)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                .afterData(ctx -> {
                    var sheet = ctx.getSheet();
                    int row = ctx.getCurrentRow();
                    String priceCol = SheetContext.columnLetter(3);
                    String qtyCol = SheetContext.columnLetter(4);
                    String subtotalCol = SheetContext.columnLetter(5);

                    var sumRow = sheet.createRow(row);
                    sumRow.createCell(2).setCellValue("합계");
                    sumRow.createCell(3).setCellFormula("SUM(%s2:%s%d)".formatted(priceCol, priceCol, row));
                    sumRow.createCell(4).setCellFormula("SUM(%s2:%s%d)".formatted(qtyCol, qtyCol, row));
                    sumRow.createCell(5).setCellFormula("SUM(%s2:%s%d)".formatted(subtotalCol, subtotalCol, row));

                    var avgRow = sheet.createRow(row + 1);
                    avgRow.createCell(2).setCellValue("평균");
                    avgRow.createCell(3).setCellFormula("AVERAGE(%s2:%s%d)".formatted(priceCol, priceCol, row));
                    avgRow.createCell(4).setCellFormula("AVERAGE(%s2:%s%d)".formatted(qtyCol, qtyCol, row));
                    avgRow.createCell(5).setCellFormula("AVERAGE(%s2:%s%d)".formatted(subtotalCol, subtotalCol, row));

                    return row + 2;
                })
                .write(sampleProducts().stream());

        return DownloadUtil.builder("formula-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 2. Hyperlink - clickable URL columns
    // ========================================================================
    @GetMapping("/hyperlink")
    public ResponseEntity<StreamingResponseBody> downloadHyperlink() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.FOREST_GREEN)
                .sheetName("Hyperlinks")
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price).type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                // Plain URL hyperlink
                .column("URL", ProductDto::url).type(ExcelDataType.HYPERLINK)
                // Hyperlink with custom label
                .column("Link", (ProductDto p) -> new ExcelHyperlink(p.url(), "상세보기"))
                    .type(ExcelDataType.HYPERLINK)
                .write(sampleProducts().stream());

        return DownloadUtil.builder("hyperlink-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 3. Schema - unified read/write with column config
    // ========================================================================
    private static final ExcelKitSchema<ProductDto> PRODUCT_SCHEMA = ExcelKitSchema.<ProductDto>builder()
            .column("Name", ProductDto::name, (p, cell) -> {})
            .column("Category", ProductDto::category, (p, cell) -> {})
            .column("Price", ProductDto::price, (p, cell) -> {},
                    c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
            .column("Quantity", ProductDto::quantity, (p, cell) -> {},
                    c -> c.type(ExcelDataType.INTEGER))
            .column("Discount", ProductDto::discount, (p, cell) -> {},
                    c -> c.type(ExcelDataType.DOUBLE_PERCENT))
            .build();

    private static final ExcelKitSchema<ProductReadDto> PRODUCT_READ_SCHEMA = ExcelKitSchema.<ProductReadDto>builder()
            .column("Name", ProductReadDto::getName, (p, cell) -> p.setName(cell.asString()))
            .column("Category", ProductReadDto::getCategory, (p, cell) -> p.setCategory(cell.asString()))
            .column("Price", ProductReadDto::getPrice, (p, cell) -> p.setPrice(cell.asInt()),
                    c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
            .column("Quantity", ProductReadDto::getQuantity, (p, cell) -> p.setQuantity(cell.asInt()),
                    c -> c.type(ExcelDataType.INTEGER))
            .column("Discount", ProductReadDto::getDiscount, (p, cell) -> p.setDiscount(cell.asDouble()),
                    c -> c.type(ExcelDataType.DOUBLE_PERCENT))
            .build();

    @GetMapping("/schema-excel")
    public ResponseEntity<StreamingResponseBody> downloadSchemaExcel() {
        var handler = PRODUCT_READ_SCHEMA.excelWriter()
                .sheetName("Schema Demo")
                .autoFilter(true)
                .freezePane(1)
                .write(sampleProducts().stream().map(p ->
                        toReadDto(p)));

        return DownloadUtil.builder("schema-excel-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    @GetMapping("/schema-csv")
    public ResponseEntity<StreamingResponseBody> downloadSchemaCsv() {
        var handler = PRODUCT_READ_SCHEMA.csvWriter()
                .write(sampleProducts().stream().map(p -> toReadDto(p)));

        return DownloadUtil.builder("schema-csv-demo", DownloadFileType.CSV)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 4. Name-based reading (upload endpoint)
    // ========================================================================
    @PostMapping("/read-by-name-excel")
    @ResponseBody
    public String readByNameExcel(MultipartFile file) throws IOException {
        List<ProductReadDto> results = new ArrayList<>();
        List<String> errors = new ArrayList<>();

        try (InputStream is = file.getInputStream()) {
            PRODUCT_READ_SCHEMA.excelReader(ProductReadDto::new, null)
                    .build(is)
                    .read(result -> {
                        if (result.success()) {
                            results.add(result.data());
                        } else {
                            errors.add(result.messages().toString());
                        }
                    });
        }

        StringBuilder sb = new StringBuilder();
        sb.append("=== Name-Based Excel Read Result ===\n");
        sb.append("Success: %d rows, Errors: %d rows\n\n".formatted(results.size(), errors.size()));
        results.forEach(p -> sb.append(p).append("\n"));
        if (!errors.isEmpty()) {
            sb.append("\n--- Errors ---\n");
            errors.forEach(e -> sb.append(e).append("\n"));
        }

        log.info("Read by name (Excel): {} success, {} errors", results.size(), errors.size());
        return sb.toString();
    }

    @PostMapping("/read-by-name-csv")
    @ResponseBody
    public String readByNameCsv(MultipartFile file) throws IOException {
        List<ProductReadDto> results = new ArrayList<>();
        List<String> errors = new ArrayList<>();

        try (InputStream is = file.getInputStream()) {
            PRODUCT_READ_SCHEMA.csvReader(ProductReadDto::new, null)
                    .build(is)
                    .read(result -> {
                        if (result.success()) {
                            results.add(result.data());
                        } else {
                            errors.add(result.messages().toString());
                        }
                    });
        }

        StringBuilder sb = new StringBuilder();
        sb.append("=== Name-Based CSV Read Result ===\n");
        sb.append("Success: %d rows, Errors: %d rows\n\n".formatted(results.size(), errors.size()));
        results.forEach(p -> sb.append(p).append("\n"));

        log.info("Read by name (CSV): {} success, {} errors", results.size(), errors.size());
        return sb.toString();
    }

    // ========================================================================
    // 5. Multi-sheet workbook with row coloring, dropdown, callbacks
    // ========================================================================
    @GetMapping("/multi-sheet")
    public ResponseEntity<StreamingResponseBody> downloadMultiSheet() {
        var products = sampleProducts();

        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.CORAL)) {
            // Sheet 1: Electronics
            wb.<ProductDto>sheet("Electronics")
                    .autoFilter()
                    .freezePane(1)
                    .column("Name", ProductDto::name)
                    .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
                    .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                    .column("Status", p -> p.quantity() > 50 ? "In Stock" : "Low Stock",
                            c -> c.dropdown("In Stock", "Low Stock", "Out of Stock"))
                    .rowColor(p -> p.quantity() <= 10 ? ExcelColor.LIGHT_RED : null)
                    .write(products.stream().filter(p -> "Electronics".equals(p.category()) || "Peripherals".equals(p.category())));

            // Sheet 2: Office & Accessories
            wb.<ProductDto>sheet("Office & Accessories")
                    .autoFilter()
                    .column("Name", ProductDto::name)
                    .column("Category", ProductDto::category)
                    .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER))
                    .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                    .column("Discount", ProductDto::discount, c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                    .rowColor(p -> p.discount() >= 0.2 ? ExcelColor.LIGHT_GREEN : null)
                    .write(products.stream().filter(p -> "Office".equals(p.category()) || "Accessories".equals(p.category())));

            // Sheet 3: Summary (different data type - String pairs)
            wb.<String[]>sheet("Summary")
                    .column("Metric", row -> row[0])
                    .column("Value", row -> row[1])
                    .write(Stream.of(
                            new String[]{"Total Products", String.valueOf(products.size())},
                            new String[]{"Categories", "Electronics, Accessories, Office, Peripherals"},
                            new String[]{"Average Price", String.valueOf(products.stream().mapToInt(ProductDto::price).average().orElse(0))},
                            new String[]{"Total Quantity", String.valueOf(products.stream().mapToInt(ProductDto::quantity).sum())}
                    ));

            var handler = wb.finish();
            return DownloadUtil.builder("multi-sheet-demo", DownloadFileType.EXCEL)
                    .body(handler::consumeOutputStream);
        }
    }

    // ========================================================================
    // 6. Full showcase - everything combined
    // ========================================================================
    @GetMapping("/full")
    public ResponseEntity<StreamingResponseBody> downloadFullShowcase() {
        var products = sampleProducts();

        var handler = new ExcelWriter<ProductDto>(ExcelColor.STEEL_BLUE)
                .sheetName("Full Showcase")
                .autoFilter(true)
                .freezePane(1)
                .rowHeight(22)
                // Row coloring (set before column chain)
                .rowColor(p -> {
                    if (p.quantity() <= 10) return ExcelColor.LIGHT_RED;
                    if (p.discount() >= 0.2) return ExcelColor.LIGHT_GREEN;
                    return null;
                })
                // No. column
                .column("No.", (ProductDto row, io.github.dornol.excelkit.shared.Cursor cursor) ->
                        cursor.getCurrentTotal()).type(ExcelDataType.LONG)
                // Basic columns
                .column("Name", ProductDto::name).bold(true)
                .column("Category", ProductDto::category)
                    .dropdown("Electronics", "Accessories", "Office", "Peripherals")
                // Styled number columns
                .column("Price", ProductDto::price)
                    .type(ExcelDataType.INTEGER)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                .column("Quantity", ProductDto::quantity)
                    .type(ExcelDataType.INTEGER)
                    .backgroundColor(ExcelColor.LIGHT_YELLOW)
                .column("Discount", ProductDto::discount)
                    .type(ExcelDataType.DOUBLE_PERCENT)
                // Formula column - subtotal
                .column("Subtotal", (ProductDto row, io.github.dornol.excelkit.shared.Cursor cursor) ->
                        "D" + (cursor.getRowOfSheet() + 1) + "*E" + (cursor.getRowOfSheet() + 1))
                    .type(ExcelDataType.FORMULA)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                // Formula column - discounted price
                .column("After Discount", (ProductDto row, io.github.dornol.excelkit.shared.Cursor cursor) ->
                        "G" + (cursor.getRowOfSheet() + 1) + "*(1-F" + (cursor.getRowOfSheet() + 1) + ")")
                    .type(ExcelDataType.FORMULA)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                // Hyperlink column
                .column("Link", (ProductDto p) -> new ExcelHyperlink(p.url(), "View"))
                    .type(ExcelDataType.HYPERLINK)
                // beforeHeader - title row
                .beforeHeader(ctx -> {
                    var sheet = ctx.getSheet();
                    var titleRow = sheet.createRow(0);
                    var cell = titleRow.createCell(0);
                    cell.setCellValue("Product Report - excel-kit Showcase");
                    var font = ctx.getWorkbook().createFont();
                    font.setBold(true);
                    font.setFontHeightInPoints((short) 14);
                    var style = ctx.getWorkbook().createCellStyle();
                    style.setFont(font);
                    cell.setCellStyle(style);
                    return 2; // skip 1 blank row after title
                })
                // afterData - summary formulas
                .afterData(ctx -> {
                    var sheet = ctx.getSheet();
                    int row = ctx.getCurrentRow();
                    int headerRow = 3; // title(0) + blank(1) + header(2), data starts at 3

                    var blankRow = sheet.createRow(row);
                    row++;

                    var sumRow = sheet.createRow(row);
                    sumRow.createCell(2).setCellValue("합계");
                    sumRow.createCell(3).setCellFormula("SUM(D%d:D%d)".formatted(headerRow + 1, row - 1));
                    sumRow.createCell(4).setCellFormula("SUM(E%d:E%d)".formatted(headerRow + 1, row - 1));
                    sumRow.createCell(6).setCellFormula("SUM(G%d:G%d)".formatted(headerRow + 1, row - 1));
                    sumRow.createCell(7).setCellFormula("SUM(H%d:H%d)".formatted(headerRow + 1, row - 1));
                    row++;

                    var avgRow = sheet.createRow(row);
                    avgRow.createCell(2).setCellValue("평균");
                    avgRow.createCell(3).setCellFormula("AVERAGE(D%d:D%d)".formatted(headerRow + 1, row - 2));
                    avgRow.createCell(4).setCellFormula("AVERAGE(E%d:E%d)".formatted(headerRow + 1, row - 2));
                    avgRow.createCell(5).setCellFormula("AVERAGE(F%d:F%d)".formatted(headerRow + 1, row - 2));
                    row++;

                    return row;
                })
                .write(products.stream());

        return DownloadUtil.builder("full-showcase", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    private static ProductReadDto toReadDto(ProductDto p) {
        var dto = new ProductReadDto();
        dto.setName(p.name());
        dto.setCategory(p.category());
        dto.setPrice(p.price());
        dto.setQuantity(p.quantity());
        dto.setDiscount(p.discount());
        dto.setUrl(p.url());
        return dto;
    }
}
