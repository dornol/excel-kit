package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.example.app.common.DownloadResponse;
import io.github.dornol.excelkit.excel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.List;
import java.util.stream.Stream;

@Controller
@RequestMapping("/showcase")
public class WriteShowcaseController {

    private static final Logger log = LoggerFactory.getLogger(WriteShowcaseController.class);

    private static List<ProductDto> sampleProducts() {
        return ShowcaseData.sampleProducts();
    }

    // ========================================================================
    // 1. Formula - FORMULA type column + SUM/AVERAGE in afterData
    // ========================================================================
    @GetMapping("/formula")
    public ResponseEntity<StreamingResponseBody> downloadFormula() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Formula Demo")
                .autoFilter(true)
                .freezeRows(1)
                .column("No.", (row, cursor) -> cursor.getCurrentTotal(), cfg -> cfg.type(ExcelDataType.LONG))
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price, cfg -> cfg.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Quantity", ProductDto::quantity, cfg -> cfg.type(ExcelDataType.INTEGER))
                .column("Subtotal", (row, cursor) ->
                        "%s%d*%s%d".formatted(
                                SheetContext.columnLetter(3), cursor.getRowOfSheet() + 1,
                                SheetContext.columnLetter(4), cursor.getRowOfSheet() + 1),
                        cfg -> cfg.type(ExcelDataType.FORMULA).format("#,##0"))
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

        return DownloadResponse.excel("formula-demo")
                .body(handler::writeTo);
    }

    // ========================================================================
    // 2. Hyperlink - plain URL + ExcelHyperlink with custom label
    // ========================================================================
    @GetMapping("/hyperlink")
    public ResponseEntity<StreamingResponseBody> downloadHyperlink() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.FOREST_GREEN)
                .sheetName("Hyperlinks")
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price, cfg -> cfg.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("URL", ProductDto::url, cfg -> cfg.type(ExcelDataType.HYPERLINK))
                .column("Link", (ProductDto p) -> new ExcelHyperlink(p.url(), "Details"),
                        cfg -> cfg.type(ExcelDataType.HYPERLINK))
                .write(sampleProducts().stream());

        return DownloadResponse.excel("hyperlink-demo")
                .body(handler::writeTo);
    }

    // ========================================================================
    // 5. Multi-sheet workbook with row coloring, dropdown, callbacks
    // ========================================================================
    @GetMapping("/multi-sheet")
    public ResponseEntity<StreamingResponseBody> downloadMultiSheet() {
        var products = sampleProducts();

        try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.CORAL)) {
            wb.<ProductDto>sheet("Electronics")
                    .autoFilter()
                    .freezeRows(1)
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
            return DownloadResponse.excel("multi-sheet-demo")
                    .body(handler::writeTo);
        }
    }

    // ========================================================================
    // 6. Cell color - per-cell conditional background
    // ========================================================================
    @GetMapping("/cell-color")
    public ResponseEntity<StreamingResponseBody> downloadCellColor() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Cell Color")
                .autoFilter(true)
                .freezeRows(1)
                .column("Name", ProductDto::name)
                .column("Price", ProductDto::price, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .format("#,##0")
                    .cellColor((value, row) -> {
                        int price = ((Number) value).intValue();
                        if (price >= 30000) return ExcelColor.LIGHT_GREEN;
                        if (price <= 5000) return ExcelColor.LIGHT_RED;
                        return null;
                    }))
                .column("Quantity", ProductDto::quantity, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .cellColor((value, row) -> {
                        int qty = ((Number) value).intValue();
                        return qty <= 10 ? ExcelColor.LIGHT_ORANGE : null;
                    }))
                .column("Discount", ProductDto::discount, cfg -> cfg
                    .type(ExcelDataType.DOUBLE_PERCENT)
                    .cellColor((value, row) ->
                        ((Number) value).doubleValue() >= 0.2 ? ExcelColor.LIGHT_PURPLE : null))
                .write(sampleProducts().stream());

        return DownloadResponse.excel("cell-color-demo")
                .body(handler::writeTo);
    }

    // ========================================================================
    // 7. Group header - merged multi-row headers
    // ========================================================================
    @GetMapping("/group-header")
    public ResponseEntity<StreamingResponseBody> downloadGroupHeader() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.CORAL)
                .sheetName("Group Header")
                .autoFilter(true)
                .freezeRows(1)
                .column("No.", (row, cursor) -> cursor.getCurrentTotal(), cfg -> cfg.type(ExcelDataType.LONG))
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .format("#,##0")
                    .group("Financial"))
                .column("Quantity", ProductDto::quantity, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .group("Financial"))
                .column("Discount", ProductDto::discount, cfg -> cfg
                    .type(ExcelDataType.DOUBLE_PERCENT)
                    .group("Financial"))
                .column("URL", ProductDto::url, cfg -> cfg
                    .type(ExcelDataType.HYPERLINK)
                    .group("Link"))
                .column("Link", (ProductDto p) -> new ExcelHyperlink(p.url(), "View"),
                        cfg -> cfg.type(ExcelDataType.HYPERLINK).group("Link"))
                .write(sampleProducts().stream());

        return DownloadResponse.excel("group-header-demo")
                .body(handler::writeTo);
    }

    // ========================================================================
    // 8. Rollover - ExcelSheetWriter auto sheet splitting
    // ========================================================================
    @GetMapping("/rollover")
    public ResponseEntity<StreamingResponseBody> downloadRollover() {
        try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.FOREST_GREEN)) {
            wb.<ProductDto>sheet("Products")
                    .maxRows(8)
                    .sheetName(idx -> "Products-Page" + (idx + 1))
                    .autoFilter()
                    .freezeRows(1)
                    .column("Name", ProductDto::name)
                    .column("Category", ProductDto::category)
                    .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                    .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                    .onProgress(5, (count, cursor) ->
                            log.info("[Rollover Demo] Processed {} rows", count))
                    .write(sampleProducts().stream());

            var handler = wb.finish();
            return DownloadResponse.excel("rollover-demo")
                    .body(handler::writeTo);
        }
    }

    // ========================================================================
    // 9. Column outline - expand/collapse column groups
    // ========================================================================
    @GetMapping("/outline")
    public ResponseEntity<StreamingResponseBody> downloadOutline() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.GOLD)
                .sheetName("Outline Demo")
                .autoFilter(true)
                .freezeRows(1)
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price, cfg -> cfg.type(ExcelDataType.INTEGER).format("#,##0").outline(1))
                .column("Quantity", ProductDto::quantity, cfg -> cfg.type(ExcelDataType.INTEGER).outline(1))
                .column("Discount", ProductDto::discount, cfg -> cfg.type(ExcelDataType.DOUBLE_PERCENT).outline(1))
                .column("URL", ProductDto::url, cfg -> cfg.type(ExcelDataType.HYPERLINK).outline(2))
                .column("Summary", p -> "%s (%s)".formatted(p.name(), p.category()))
                .write(sampleProducts().stream());

        return DownloadResponse.excel("outline-demo")
                .body(handler::writeTo);
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

        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Full Showcase")
                .autoFilter(true)
                .freezeRows(1)
                .rowHeight(22)
                .rowColor(p -> {
                    if (p.quantity() <= 10) return ExcelColor.LIGHT_RED;
                    if (p.discount() >= 0.2) return ExcelColor.LIGHT_GREEN;
                    return null;
                })
                .column("No.", (row, cursor) -> cursor.getCurrentTotal(), cfg -> cfg.type(ExcelDataType.LONG))
                .column("Name", ProductDto::name, cfg -> cfg.bold(true))
                .column("Category", ProductDto::category, cfg -> cfg
                    .dropdown("Electronics", "Accessories", "Office", "Peripherals"))
                .column("Price", ProductDto::price, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .format("#,##0")
                    .group("Financial")
                    .cellColor((value, row) -> {
                        int price = ((Number) value).intValue();
                        return price >= 30000 ? ExcelColor.LIGHT_GREEN : null;
                    }))
                .column("Quantity", ProductDto::quantity, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .group("Financial")
                    .cellColor((value, row) ->
                        ((Number) value).intValue() <= 10 ? ExcelColor.LIGHT_ORANGE : null))
                .column("Discount", ProductDto::discount, cfg -> cfg
                    .type(ExcelDataType.DOUBLE_PERCENT)
                    .group("Financial"))
                .column("Subtotal", (row, cursor) -> {
                    int r = cursor.getRowOfSheet() + 1;
                    return "%s%d*%s%d".formatted(priceCol, r, qtyCol, r);
                }, cfg -> cfg.type(ExcelDataType.FORMULA).format("#,##0"))
                .column("After Discount", (row, cursor) -> {
                    int r = cursor.getRowOfSheet() + 1;
                    return "%s%d*(1-%s%d)".formatted(subtotalCol, r, discountCol, r);
                }, cfg -> cfg.type(ExcelDataType.FORMULA).format("#,##0"))
                .column("Link", (ProductDto p) -> new ExcelHyperlink(p.url(), "View"),
                        cfg -> cfg.type(ExcelDataType.HYPERLINK))
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

        return DownloadResponse.excel("full-showcase")
                .body(handler::writeTo);
    }

}
