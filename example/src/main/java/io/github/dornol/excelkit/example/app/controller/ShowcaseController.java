package io.github.dornol.excelkit.example.app.controller;

import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.example.app.dto.ProductReadDto;
import io.github.dornol.excelkit.example.app.util.DownloadFileType;
import io.github.dornol.excelkit.example.app.util.DownloadUtil;
import io.github.dornol.excelkit.csv.CsvMapWriter;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.shared.ExcelKitSchema;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
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
import java.util.*;
import java.util.stream.Stream;

/**
 * Showcase controller demonstrating new and existing excel-kit features.
 * All endpoints use in-memory data (no DB required).
 */
@Controller
@RequestMapping("/showcase")
public class ShowcaseController {

    private static final Logger log = LoggerFactory.getLogger(ShowcaseController.class);

    private static final ExcelKitSchema<ProductReadDto> PRODUCT_SCHEMA = ExcelKitSchema.<ProductReadDto>builder()
            .column("Name", ProductReadDto::getName, (p, cell) -> p.setName(cell.asString()))
            .column("Category", ProductReadDto::getCategory, (p, cell) -> p.setCategory(cell.asString()))
            .column("Price", ProductReadDto::getPrice, (p, cell) -> p.setPrice(cell.asInt()),
                    c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
            .column("Quantity", ProductReadDto::getQuantity, (p, cell) -> p.setQuantity(cell.asInt()),
                    c -> c.type(ExcelDataType.INTEGER))
            .column("Discount", ProductReadDto::getDiscount, (p, cell) -> p.setDiscount(cell.asDouble()),
                    c -> c.type(ExcelDataType.DOUBLE_PERCENT))
            .build();

    private static List<ProductDto> sampleProducts() {
        return Stream.generate(ProductDto::random).limit(20).toList();
    }

    // ========================================================================
    // 1. Formula - FORMULA type column + SUM/AVERAGE in afterData
    // ========================================================================
    @GetMapping("/formula")
    public ResponseEntity<StreamingResponseBody> downloadFormula() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.STEEL_BLUE)
                .sheetName("Formula Demo")
                .autoFilter(true)
                .freezePane(1)
                .column("No.", (row, cursor) -> cursor.getCurrentTotal()).type(ExcelDataType.LONG)
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price).type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                .column("Quantity", ProductDto::quantity).type(ExcelDataType.INTEGER)
                .column("Subtotal", (row, cursor) ->
                        "%s%d*%s%d".formatted(
                                SheetContext.columnLetter(3), cursor.getRowOfSheet() + 1,
                                SheetContext.columnLetter(4), cursor.getRowOfSheet() + 1))
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
    // 2. Hyperlink - plain URL + ExcelHyperlink with custom label
    // ========================================================================
    @GetMapping("/hyperlink")
    public ResponseEntity<StreamingResponseBody> downloadHyperlink() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.FOREST_GREEN)
                .sheetName("Hyperlinks")
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price).type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                .column("URL", ProductDto::url).type(ExcelDataType.HYPERLINK)
                .column("Link", (ProductDto p) -> new ExcelHyperlink(p.url(), "상세보기"))
                    .type(ExcelDataType.HYPERLINK)
                .write(sampleProducts().stream());

        return DownloadUtil.builder("hyperlink-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 3. Schema - unified read/write with column config (name-based read)
    // ========================================================================
    @GetMapping("/schema-excel")
    public ResponseEntity<StreamingResponseBody> downloadSchemaExcel() {
        var handler = PRODUCT_SCHEMA.excelWriter()
                .sheetName("Schema Demo")
                .autoFilter(true)
                .freezePane(1)
                .write(sampleProducts().stream().map(ShowcaseController::toReadDto));

        return DownloadUtil.builder("schema-excel-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    @GetMapping("/schema-csv")
    public ResponseEntity<StreamingResponseBody> downloadSchemaCsv() {
        var handler = PRODUCT_SCHEMA.csvWriter()
                .write(sampleProducts().stream().map(ShowcaseController::toReadDto));

        return DownloadUtil.builder("schema-csv-demo", DownloadFileType.CSV)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 4. Name-based reading (upload endpoint)
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
    // 5. Multi-sheet workbook with row coloring, dropdown, callbacks
    // ========================================================================
    @GetMapping("/multi-sheet")
    public ResponseEntity<StreamingResponseBody> downloadMultiSheet() {
        var products = sampleProducts();

        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.CORAL)) {
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

            wb.<ProductDto>sheet("Office & Accessories")
                    .autoFilter()
                    .column("Name", ProductDto::name)
                    .column("Category", ProductDto::category)
                    .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER))
                    .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                    .column("Discount", ProductDto::discount, c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                    .rowColor(p -> p.discount() >= 0.2 ? ExcelColor.LIGHT_GREEN : null)
                    .write(products.stream().filter(p -> "Office".equals(p.category()) || "Accessories".equals(p.category())));

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
    // 6. Cell color - per-cell conditional background
    // ========================================================================
    @GetMapping("/cell-color")
    public ResponseEntity<StreamingResponseBody> downloadCellColor() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.STEEL_BLUE)
                .sheetName("Cell Color")
                .autoFilter(true)
                .freezePane(1)
                .column("Name", ProductDto::name)
                .column("Price", ProductDto::price)
                    .type(ExcelDataType.INTEGER)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                    .cellColor((value, row) -> {
                        int price = ((Number) value).intValue();
                        if (price >= 30000) return ExcelColor.LIGHT_GREEN;
                        if (price <= 5000) return ExcelColor.LIGHT_RED;
                        return null;
                    })
                .column("Quantity", ProductDto::quantity)
                    .type(ExcelDataType.INTEGER)
                    .cellColor((value, row) -> {
                        int qty = ((Number) value).intValue();
                        return qty <= 10 ? ExcelColor.LIGHT_ORANGE : null;
                    })
                .column("Discount", ProductDto::discount)
                    .type(ExcelDataType.DOUBLE_PERCENT)
                    .cellColor((value, row) ->
                        ((Number) value).doubleValue() >= 0.2 ? ExcelColor.LIGHT_PURPLE : null)
                .write(sampleProducts().stream());

        return DownloadUtil.builder("cell-color-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 7. Group header - merged multi-row headers
    // ========================================================================
    @GetMapping("/group-header")
    public ResponseEntity<StreamingResponseBody> downloadGroupHeader() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.CORAL)
                .sheetName("Group Header")
                .autoFilter(true)
                .freezePane(1)
                .column("No.", (row, cursor) -> cursor.getCurrentTotal()).type(ExcelDataType.LONG)
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price)
                    .type(ExcelDataType.INTEGER)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                    .group("Financial")
                .column("Quantity", ProductDto::quantity)
                    .type(ExcelDataType.INTEGER)
                    .group("Financial")
                .column("Discount", ProductDto::discount)
                    .type(ExcelDataType.DOUBLE_PERCENT)
                    .group("Financial")
                .column("URL", ProductDto::url)
                    .type(ExcelDataType.HYPERLINK)
                    .group("Link")
                .column("Link", (ProductDto p) -> new ExcelHyperlink(p.url(), "View"))
                    .type(ExcelDataType.HYPERLINK)
                    .group("Link")
                .write(sampleProducts().stream());

        return DownloadUtil.builder("group-header-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 8. Rollover - ExcelSheetWriter auto sheet splitting
    // ========================================================================
    @GetMapping("/rollover")
    public ResponseEntity<StreamingResponseBody> downloadRollover() {
        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.FOREST_GREEN)) {
            wb.<ProductDto>sheet("Products")
                    .maxRows(8)
                    .sheetName(idx -> "Products-Page" + (idx + 1))
                    .autoFilter()
                    .freezePane(1)
                    .column("Name", ProductDto::name)
                    .column("Category", ProductDto::category)
                    .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
                    .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                    .onProgress(5, (count, cursor) ->
                            log.info("[Rollover Demo] Processed {} rows", count))
                    .write(sampleProducts().stream());

            var handler = wb.finish();
            return DownloadUtil.builder("rollover-demo", DownloadFileType.EXCEL)
                    .body(handler::consumeOutputStream);
        }
    }

    // ========================================================================
    // 9. Column outline - expand/collapse column groups
    // ========================================================================
    @GetMapping("/outline")
    public ResponseEntity<StreamingResponseBody> downloadOutline() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.GOLD)
                .sheetName("Outline Demo")
                .autoFilter(true)
                .freezePane(1)
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price).type(ExcelDataType.INTEGER)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat()).outline(1)
                .column("Quantity", ProductDto::quantity).type(ExcelDataType.INTEGER).outline(1)
                .column("Discount", ProductDto::discount).type(ExcelDataType.DOUBLE_PERCENT).outline(1)
                .column("URL", ProductDto::url).type(ExcelDataType.HYPERLINK).outline(2)
                .column("Summary", p -> "%s (%s)".formatted(p.name(), p.category()))
                .write(sampleProducts().stream());

        return DownloadUtil.builder("outline-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
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
    // 11. Full showcase - everything combined
    // ========================================================================
    @GetMapping("/full")
    public ResponseEntity<StreamingResponseBody> downloadFullShowcase() {
        String priceCol = SheetContext.columnLetter(3);      // D
        String qtyCol = SheetContext.columnLetter(4);        // E
        String discountCol = SheetContext.columnLetter(5);   // F
        String subtotalCol = SheetContext.columnLetter(6);   // G
        String afterDiscCol = SheetContext.columnLetter(7);  // H

        var handler = new ExcelWriter<ProductDto>(ExcelColor.STEEL_BLUE)
                .sheetName("Full Showcase")
                .autoFilter(true)
                .freezePane(1)
                .rowHeight(22)
                .rowColor(p -> {
                    if (p.quantity() <= 10) return ExcelColor.LIGHT_RED;
                    if (p.discount() >= 0.2) return ExcelColor.LIGHT_GREEN;
                    return null;
                })
                .column("No.", (row, cursor) -> cursor.getCurrentTotal()).type(ExcelDataType.LONG)
                .column("Name", ProductDto::name).bold(true)
                .column("Category", ProductDto::category)
                    .dropdown("Electronics", "Accessories", "Office", "Peripherals")
                .column("Price", ProductDto::price)
                    .type(ExcelDataType.INTEGER)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                    .group("Financial")
                    .cellColor((value, row) -> {
                        int price = ((Number) value).intValue();
                        return price >= 30000 ? ExcelColor.LIGHT_GREEN : null;
                    })
                .column("Quantity", ProductDto::quantity)
                    .type(ExcelDataType.INTEGER)
                    .group("Financial")
                    .cellColor((value, row) ->
                        ((Number) value).intValue() <= 10 ? ExcelColor.LIGHT_ORANGE : null)
                .column("Discount", ProductDto::discount)
                    .type(ExcelDataType.DOUBLE_PERCENT)
                    .group("Financial")
                .column("Subtotal", (row, cursor) -> {
                    int r = cursor.getRowOfSheet() + 1;
                    return "%s%d*%s%d".formatted(priceCol, r, qtyCol, r);
                })
                    .type(ExcelDataType.FORMULA)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                .column("After Discount", (row, cursor) -> {
                    int r = cursor.getRowOfSheet() + 1;
                    return "%s%d*(1-%s%d)".formatted(subtotalCol, r, discountCol, r);
                })
                    .type(ExcelDataType.FORMULA)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                .column("Link", (ProductDto p) -> new ExcelHyperlink(p.url(), "View"))
                    .type(ExcelDataType.HYPERLINK)
                .beforeHeader(ctx -> {
                    var titleRow = ctx.getSheet().createRow(0);
                    var cell = titleRow.createCell(0);
                    cell.setCellValue("Product Report - excel-kit Showcase");
                    var font = ctx.getWorkbook().createFont();
                    font.setBold(true);
                    font.setFontHeightInPoints((short) 14);
                    var style = ctx.getWorkbook().createCellStyle();
                    style.setFont(font);
                    cell.setCellStyle(style);
                    return 2;
                })
                .afterData(ctx -> {
                    var sheet = ctx.getSheet();
                    int row = ctx.getCurrentRow();
                    int dataStart = 4; // title(0) + blank(1) + header(2) + data starts at row 3 → Excel row 4

                    sheet.createRow(row);
                    row++;

                    var sumRow = sheet.createRow(row);
                    sumRow.createCell(2).setCellValue("합계");
                    sumRow.createCell(3).setCellFormula("SUM(%s%d:%s%d)".formatted(priceCol, dataStart, priceCol, row - 1));
                    sumRow.createCell(4).setCellFormula("SUM(%s%d:%s%d)".formatted(qtyCol, dataStart, qtyCol, row - 1));
                    sumRow.createCell(6).setCellFormula("SUM(%s%d:%s%d)".formatted(subtotalCol, dataStart, subtotalCol, row - 1));
                    sumRow.createCell(7).setCellFormula("SUM(%s%d:%s%d)".formatted(afterDiscCol, dataStart, afterDiscCol, row - 1));
                    row++;

                    var avgRow = sheet.createRow(row);
                    avgRow.createCell(2).setCellValue("평균");
                    avgRow.createCell(3).setCellFormula("AVERAGE(%s%d:%s%d)".formatted(priceCol, dataStart, priceCol, row - 2));
                    avgRow.createCell(4).setCellFormula("AVERAGE(%s%d:%s%d)".formatted(qtyCol, dataStart, qtyCol, row - 2));
                    avgRow.createCell(5).setCellFormula("AVERAGE(%s%d:%s%d)".formatted(discountCol, dataStart, discountCol, row - 2));
                    row++;

                    return row;
                })
                .write(sampleProducts().stream());

        return DownloadUtil.builder("full-showcase", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 12. Border Style - configurable cell borders
    // ========================================================================
    @GetMapping("/border-style")
    public ResponseEntity<StreamingResponseBody> downloadBorderStyle() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.STEEL_BLUE)
                .sheetName("Border Styles")
                .column("Name", ProductDto::name).border(ExcelBorderStyle.MEDIUM)
                .column("Category", ProductDto::category).border(ExcelBorderStyle.DASHED)
                .column("Price", ProductDto::price)
                    .type(ExcelDataType.INTEGER)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                    .border(ExcelBorderStyle.THICK)
                .column("Quantity", ProductDto::quantity)
                    .type(ExcelDataType.INTEGER)
                    .border(ExcelBorderStyle.DOTTED)
                .column("No Border", p -> "text").border(ExcelBorderStyle.NONE)
                .write(sampleProducts().stream());

        return DownloadUtil.builder("border-style-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 13. Cell Comments - per-cell notes
    // ========================================================================
    @GetMapping("/cell-comment")
    public ResponseEntity<StreamingResponseBody> downloadCellComment() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.FOREST_GREEN)
                .sheetName("Cell Comments")
                .autoFilter(true)
                .column("Name", ProductDto::name)
                    .comment(p -> "Product: " + p.name() + " (" + p.category() + ")")
                .column("Price", ProductDto::price)
                    .type(ExcelDataType.INTEGER)
                    .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                    .comment(p -> p.price() > 30000 ? "⚠ High price!" : null)
                .column("Quantity", ProductDto::quantity)
                    .type(ExcelDataType.INTEGER)
                    .comment(p -> p.quantity() <= 10 ? "⚠ Low stock alert" : null)
                .column("Discount", ProductDto::discount)
                    .type(ExcelDataType.DOUBLE_PERCENT)
                .write(sampleProducts().stream());

        return DownloadUtil.builder("cell-comment-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 14. Conditional Formatting - Excel native rules
    // ========================================================================
    @GetMapping("/conditional-format")
    public ResponseEntity<StreamingResponseBody> downloadConditionalFormat() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.GOLD)
                .sheetName("Conditional Formatting")
                .autoFilter(true)
                .freezePane(1)
                .addColumn("Name", ProductDto::name)
                .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
                .addColumn("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                .addColumn("Discount", p -> p.discount(), c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .conditionalFormatting(cf -> cf
                        .columns(1)  // Price column
                        .greaterThan("30000", ExcelColor.LIGHT_RED)
                        .lessThan("5000", ExcelColor.LIGHT_GREEN))
                .conditionalFormatting(cf -> cf
                        .columns(2)  // Quantity column
                        .lessThanOrEqual("10", ExcelColor.LIGHT_ORANGE))
                .conditionalFormatting(cf -> cf
                        .columns(3)  // Discount column
                        .greaterThanOrEqual("0.2", ExcelColor.LIGHT_PURPLE))
                .write(sampleProducts().stream());

        return DownloadUtil.builder("conditional-format-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 15. Sheet Protection - lock/unlock columns
    // ========================================================================
    @GetMapping("/sheet-protection")
    public ResponseEntity<StreamingResponseBody> downloadSheetProtection() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.CORAL)
                .sheetName("Protected Sheet")
                .autoFilter(true)
                .addColumn("Name (locked)", ProductDto::name, c -> c.locked(true))
                .addColumn("Category (locked)", ProductDto::category, c -> c.locked(true))
                .addColumn("Price (editable)", p -> p.price(), c -> c
                        .type(ExcelDataType.INTEGER)
                        .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                        .locked(false)
                        .backgroundColor(ExcelColor.LIGHT_GREEN))
                .addColumn("Quantity (editable)", p -> p.quantity(), c -> c
                        .type(ExcelDataType.INTEGER)
                        .locked(false)
                        .backgroundColor(ExcelColor.LIGHT_GREEN))
                .addColumn("Discount (locked)", p -> p.discount(), c -> c
                        .type(ExcelDataType.DOUBLE_PERCENT)
                        .locked(true))
                .protectSheet("1234")
                .write(sampleProducts().stream());

        return DownloadUtil.builder("sheet-protection-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 16. Chart - bar/line/pie chart generation
    // ========================================================================
    @GetMapping("/chart")
    public ResponseEntity<StreamingResponseBody> downloadChart() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.STEEL_BLUE)
                .sheetName("Chart Demo")
                .addColumn("Name", ProductDto::name)
                .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
                .addColumn("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Product Price vs Quantity")
                        .categoryColumn(0)
                        .valueColumn(1, "Price")
                        .valueColumn(2, "Quantity"))
                .write(sampleProducts().stream().limit(10));

        return DownloadUtil.builder("chart-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 17. Map Writer - write Map<String, Object> data
    // ========================================================================
    @GetMapping("/map-writer")
    public ResponseEntity<StreamingResponseBody> downloadMapWriter() {
        var products = sampleProducts();
        var maps = products.stream()
                .<Map<String, Object>>map(p -> {
                    Map<String, Object> m = new LinkedHashMap<>();
                    m.put("Name", p.name());
                    m.put("Category", p.category());
                    m.put("Price", p.price());
                    m.put("Quantity", p.quantity());
                    m.put("Discount", String.format("%.0f%%", p.discount() * 100));
                    return m;
                });

        var handler = new ExcelMapWriter("Name", "Category", "Price", "Quantity", "Discount")
                .write(maps);

        return DownloadUtil.builder("map-writer-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
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
    // 20. Workbook Protection - protectWorkbook + protectSheet combined
    // ========================================================================
    @GetMapping("/workbook-protection")
    public ResponseEntity<StreamingResponseBody> downloadWorkbookProtection() {
        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
            wb.protectWorkbook("secret");

            wb.<ProductDto>sheet("Protected Data")
                    .protectSheet("1234")
                    .autoFilter()
                    .freezePane(1)
                    .column("Name (locked)", ProductDto::name, c -> c.locked(true))
                    .column("Category (locked)", ProductDto::category, c -> c.locked(true))
                    .column("Price (editable)", p -> p.price(), c -> c
                            .type(ExcelDataType.INTEGER)
                            .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                            .locked(false)
                            .backgroundColor(ExcelColor.LIGHT_GREEN))
                    .column("Quantity (editable)", p -> p.quantity(), c -> c
                            .type(ExcelDataType.INTEGER)
                            .locked(false)
                            .backgroundColor(ExcelColor.LIGHT_GREEN))
                    .column("Discount (locked)", p -> p.discount(), c -> c
                            .type(ExcelDataType.DOUBLE_PERCENT)
                            .locked(true))
                    .write(sampleProducts().stream());

            var handler = wb.finish();
            return DownloadUtil.builder("workbook-protection-demo", DownloadFileType.EXCEL)
                    .body(handler::consumeOutputStream);
        }
    }

    // ========================================================================
    // 21. Header Font - headerFontName + headerFontSize
    // ========================================================================
    @GetMapping("/header-font")
    public ResponseEntity<StreamingResponseBody> downloadHeaderFont() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.CORAL)
                .sheetName("Header Font Demo")
                .headerFontName("Arial")
                .headerFontSize(14)
                .autoFilter(true)
                .freezePane(1)
                .addColumn("Name", ProductDto::name)
                .addColumn("Category", ProductDto::category)
                .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
                .addColumn("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                .addColumn("Discount", p -> p.discount(), c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .write(sampleProducts().stream());

        return DownloadUtil.builder("header-font-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 22. Default Style - defaultStyle with fontName, fontSize, alignment
    // ========================================================================
    @GetMapping("/default-style")
    public ResponseEntity<StreamingResponseBody> downloadDefaultStyle() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.FOREST_GREEN)
                .sheetName("Default Style Demo")
                .autoFilter(true)
                .freezePane(1)
                .defaultStyle(d -> d
                        .fontName("Arial")
                        .fontSize(10)
                        .alignment(HorizontalAlignment.LEFT))
                .addColumn("Name", ProductDto::name)
                .addColumn("Category", ProductDto::category)
                .addColumn("Price", p -> p.price(), c -> c
                        .type(ExcelDataType.INTEGER)
                        .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                        .alignment(HorizontalAlignment.RIGHT))
                .addColumn("Quantity", p -> p.quantity(), c -> c
                        .type(ExcelDataType.INTEGER)
                        .alignment(HorizontalAlignment.RIGHT))
                .addColumn("Discount", p -> p.discount(), c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .write(sampleProducts().stream());

        return DownloadUtil.builder("default-style-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 23. Summary Rows - summary with sum + average
    // ========================================================================
    @GetMapping("/summary")
    public ResponseEntity<StreamingResponseBody> downloadSummary() {
        var handler = new ExcelWriter<ProductDto>(ExcelColor.GOLD)
                .sheetName("Summary Demo")
                .autoFilter(true)
                .freezePane(1)
                .addColumn("Name", ProductDto::name)
                .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
                .addColumn("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s
                        .label("Total")
                        .sum("Price").sum("Quantity")
                        .average("Price").average("Quantity"))
                .write(sampleProducts().stream());

        return DownloadUtil.builder("summary-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    // ========================================================================
    // 24. Named Range + List Validation - combined demo
    // ========================================================================
    @GetMapping("/named-range")
    public ResponseEntity<StreamingResponseBody> downloadNamedRange() {
        var categories = List.of("Electronics", "Accessories", "Office", "Peripherals");

        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
            wb.<String>sheet("Options")
                    .column("Category", s -> s)
                    .afterData(ctx -> {
                        ctx.namedRange("CategoryList", 0, 1, categories.size());
                        return ctx.getCurrentRow();
                    })
                    .write(categories.stream());

            wb.<ProductDto>sheet("Data")
                    .autoFilter()
                    .freezePane(1)
                    .column("Name", ProductDto::name)
                    .column("Category", ProductDto::category, c -> c
                            .validation(ExcelValidation.listFromRange("Options!$A$2:$A$5")))
                    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
                    .column("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                    .write(sampleProducts().stream());

            var handler = wb.finish();
            return DownloadUtil.builder("named-range-demo", DownloadFileType.EXCEL)
                    .body(handler::consumeOutputStream);
        }
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

    // ============================================================
    // Template-based writing
    // ============================================================

    /**
     * Creates a template in memory (simulates loading a .xlsx template file),
     * fills in cell values and list data, and streams the result.
     */
    @GetMapping("/template-write")
    public ResponseEntity<StreamingResponseBody> templateWrite() {
        return DownloadUtil.builder("invoice", DownloadFileType.EXCEL).body(out -> {
            // 1. Create a template in memory (in real apps, load from classpath/filesystem)
            byte[] templateBytes;
            try (var twb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
                var sheet = twb.createSheet("Invoice");
                var titleRow = sheet.createRow(0);
                titleRow.createCell(0).setCellValue("INVOICE");
                sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0, 3));

                sheet.createRow(2).createCell(0).setCellValue("Client:");
                sheet.createRow(3).createCell(0).setCellValue("Date:");

                var headerRow = sheet.createRow(5);
                headerRow.createCell(0).setCellValue("Product");
                headerRow.createCell(1).setCellValue("Qty");
                headerRow.createCell(2).setCellValue("Price");
                headerRow.createCell(3).setCellValue("Amount");

                var bos = new java.io.ByteArrayOutputStream();
                twb.write(bos);
                templateBytes = bos.toByteArray();
            }

            // 2. Fill template with data
            var products = sampleProducts();
            try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(templateBytes))) {
                // Cell-level writes (fill placeholders)
                writer.cell("B3", "Acme Corporation")
                      .cell("B4", java.time.LocalDate.now());

                // List streaming (starting at row 6, headers are already in row 5)
                writer.<ProductDto>list(6)
                      .column("Product", ProductDto::name)
                      .column("Qty", p -> 10, c -> c.type(ExcelDataType.INTEGER))
                      .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER))
                      .column("Amount", p -> p.price() * 10, c -> c.type(ExcelDataType.INTEGER))
                      .afterData(ctx -> {
                          var row = ctx.getSheet().createRow(ctx.getCurrentRow());
                          row.createCell(0).setCellValue("Total");
                          row.createCell(3).setCellFormula(
                                  "SUM(D7:D" + ctx.getCurrentRow() + ")");
                          return ctx.getCurrentRow() + 1;
                      })
                      .write(products.stream());

                writer.finish().consumeOutputStream(out);
            }
        });
    }

    /**
     * Demonstrates cell-only template writing (e.g., certificate/document style).
     */
    @GetMapping("/template-cell-only")
    public ResponseEntity<StreamingResponseBody> templateCellOnly() {
        return DownloadUtil.builder("certificate", DownloadFileType.EXCEL).body(out -> {
            byte[] templateBytes;
            try (var twb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
                var sheet = twb.createSheet("Certificate");
                sheet.createRow(1).createCell(1).setCellValue("Certificate of Employment");
                sheet.createRow(3).createCell(0).setCellValue("Name:");
                sheet.createRow(4).createCell(0).setCellValue("Department:");
                sheet.createRow(5).createCell(0).setCellValue("Position:");
                sheet.createRow(6).createCell(0).setCellValue("Join Date:");
                sheet.createRow(8).createCell(0).setCellValue("Issue Date:");

                var bos = new java.io.ByteArrayOutputStream();
                twb.write(bos);
                templateBytes = bos.toByteArray();
            }

            try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(templateBytes))) {
                writer.cell("B4", "John Doe")
                      .cell("B5", "Engineering")
                      .cell("B6", "Senior Developer")
                      .cell("B7", java.time.LocalDate.of(2020, 3, 15))
                      .cell("B9", java.time.LocalDate.now())
                      .finish()
                      .consumeOutputStream(out);
            }
        });
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
