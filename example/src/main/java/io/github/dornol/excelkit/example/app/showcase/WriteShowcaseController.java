package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.example.app.common.DownloadFileType;
import io.github.dornol.excelkit.example.app.common.DownloadUtil;
import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.shared.ExcelKitSchema;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

@Controller
@RequestMapping("/showcase")
public class WriteShowcaseController {

    private static final Logger log = LoggerFactory.getLogger(WriteShowcaseController.class);

    private static final ExcelKitSchema<io.github.dornol.excelkit.example.app.dto.ProductReadDto> PRODUCT_SCHEMA = ShowcaseData.PRODUCT_SCHEMA;

    private static List<ProductDto> sampleProducts() {
        return ShowcaseData.sampleProducts();
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
                .column("Price", ProductDto::price).type(ExcelDataType.INTEGER).format("#,##0")
                .column("Quantity", ProductDto::quantity).type(ExcelDataType.INTEGER)
                .column("Subtotal", (row, cursor) ->
                        "%s%d*%s%d".formatted(
                                SheetContext.columnLetter(3), cursor.getRowOfSheet() + 1,
                                SheetContext.columnLetter(4), cursor.getRowOfSheet() + 1))
                    .type(ExcelDataType.FORMULA)
                    .format("#,##0")
                .afterData(ctx -> {
                    var sheet = ctx.getSheet();
                    int row = ctx.getCurrentRow();
                    String priceCol = SheetContext.columnLetter(3);
                    String qtyCol = SheetContext.columnLetter(4);
                    String subtotalCol = SheetContext.columnLetter(5);

                    var sumRow = sheet.createRow(row);
                    sumRow.createCell(2).setCellValue("Total");
                    sumRow.createCell(3).setCellFormula("SUM(%s2:%s%d)".formatted(priceCol, priceCol, row));
                    sumRow.createCell(4).setCellFormula("SUM(%s2:%s%d)".formatted(qtyCol, qtyCol, row));
                    sumRow.createCell(5).setCellFormula("SUM(%s2:%s%d)".formatted(subtotalCol, subtotalCol, row));

                    var avgRow = sheet.createRow(row + 1);
                    avgRow.createCell(2).setCellValue("Average");
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
                .column("Price", ProductDto::price).type(ExcelDataType.INTEGER).format("#,##0")
                .column("URL", ProductDto::url).type(ExcelDataType.HYPERLINK)
                .column("Link", (ProductDto p) -> new ExcelHyperlink(p.url(), "Details"))
                    .type(ExcelDataType.HYPERLINK)
                .write(sampleProducts().stream());

        return DownloadUtil.builder("hyperlink-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
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
                    .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
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
                    .format("#,##0")
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
                    .format("#,##0")
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
                    .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
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
                    .format("#,##0").outline(1)
                .column("Quantity", ProductDto::quantity).type(ExcelDataType.INTEGER).outline(1)
                .column("Discount", ProductDto::discount).type(ExcelDataType.DOUBLE_PERCENT).outline(1)
                .column("URL", ProductDto::url).type(ExcelDataType.HYPERLINK).outline(2)
                .column("Summary", p -> "%s (%s)".formatted(p.name(), p.category()))
                .write(sampleProducts().stream());

        return DownloadUtil.builder("outline-demo", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
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
                    .format("#,##0")
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
                    .format("#,##0")
                .column("After Discount", (row, cursor) -> {
                    int r = cursor.getRowOfSheet() + 1;
                    return "%s%d*(1-%s%d)".formatted(subtotalCol, r, discountCol, r);
                })
                    .type(ExcelDataType.FORMULA)
                    .format("#,##0")
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
                    sumRow.createCell(2).setCellValue("Total");
                    sumRow.createCell(3).setCellFormula("SUM(%s%d:%s%d)".formatted(priceCol, dataStart, priceCol, row - 1));
                    sumRow.createCell(4).setCellFormula("SUM(%s%d:%s%d)".formatted(qtyCol, dataStart, qtyCol, row - 1));
                    sumRow.createCell(6).setCellFormula("SUM(%s%d:%s%d)".formatted(subtotalCol, dataStart, subtotalCol, row - 1));
                    sumRow.createCell(7).setCellFormula("SUM(%s%d:%s%d)".formatted(afterDiscCol, dataStart, afterDiscCol, row - 1));
                    row++;

                    var avgRow = sheet.createRow(row);
                    avgRow.createCell(2).setCellValue("Average");
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
                    .format("#,##0")
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
                    .format("#,##0")
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
                .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
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
                        .format("#,##0")
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
                .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
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
                            .format("#,##0")
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
                .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
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
                        .format("#,##0")
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
                .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
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
                    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                    .column("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                    .write(sampleProducts().stream());

            var handler = wb.finish();
            return DownloadUtil.builder("named-range-demo", DownloadFileType.EXCEL)
                    .body(handler::consumeOutputStream);
        }
    }

    // ========================================================================
    // Data Bar / Icon Set (from NewFeaturesController)
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

}
