package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.spring.ExcelKitResponse;
import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.excel.feature.ExcelFormula;
import io.github.dornol.excelkit.excel.style.ExcelBorderStyle;
import io.github.dornol.excelkit.excel.style.ExcelDataFormat;
import io.github.dornol.excelkit.excel.feature.ExcelCellComment;
import io.github.dornol.excelkit.excel.ExcelChartConfig;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.excel.ExcelConditionalRule;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelHandler;
import io.github.dornol.excelkit.excel.ExcelValidation;
import io.github.dornol.excelkit.excel.ExcelWorkbook;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

@Controller
@RequestMapping("/showcase")
public class WriteShowcaseAdvancedController {

    private static List<ProductDto> sampleProducts() {
        return ShowcaseData.sampleProducts();
    }

    @GetMapping("/border-style")
    public ResponseEntity<StreamingResponseBody> downloadBorderStyle() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Border Styles")
                .column("Name", ProductDto::name, cfg -> cfg.border(ExcelBorderStyle.MEDIUM))
                .column("Category", ProductDto::category, cfg -> cfg.border(ExcelBorderStyle.DASHED))
                .column("Price", ProductDto::price, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .format("#,##0")
                    .border(ExcelBorderStyle.THICK))
                .column("Quantity", ProductDto::quantity, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .border(ExcelBorderStyle.DOTTED))
                .column("No Border", p -> "text", cfg -> cfg.border(ExcelBorderStyle.NONE))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("border-style-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/cell-comment")
    public ResponseEntity<StreamingResponseBody> downloadCellComment() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.FOREST_GREEN)
                .sheetName("Cell Comments")
                .autoFilter(true)
                .column("Name", ProductDto::name, cfg -> cfg
                    .comment(p -> "Product: " + p.name() + " (" + p.category() + ")"))
                .column("Price", ProductDto::price, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .format("#,##0")
                    .comment(p -> p.price() > 30000 ? "⚠ High price!" : null))
                .column("Quantity", ProductDto::quantity, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .comment(p -> p.quantity() <= 10 ? "⚠ Low stock alert" : null))
                .column("Discount", ProductDto::discount, cfg -> cfg
                    .type(ExcelDataType.DOUBLE_PERCENT))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("cell-comment-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/header-comment")
    public ResponseEntity<StreamingResponseBody> downloadHeaderComment() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Header Comments")
                .autoFilter(true)
                .column("Name", ProductDto::name, cfg -> cfg
                    .headerComment("Product name (max 50 chars)"))
                .column("Price", ProductDto::price, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .format("#,##0")
                    .headerComment("Unit price in KRW (no decimals)"))
                .column("Quantity", ProductDto::quantity, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .headerComment("Stock quantity (integer only)"))
                .column("Discount", ProductDto::discount, cfg -> cfg
                    .type(ExcelDataType.DOUBLE_PERCENT)
                    .headerComment("Discount rate (0.0 - 1.0)"))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("header-comment-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/group-header-multi")
    public ResponseEntity<StreamingResponseBody> downloadMultiLevelGroupHeader() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Multi-Level Group")
                .autoFilter(true)
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category, c -> c.group("Meta"))
                .column("Price", ProductDto::price, c -> c
                        .type(ExcelDataType.INTEGER)
                        .format("#,##0")
                        .group("Financial", "Revenue"))
                .column("Quantity", ProductDto::quantity, c -> c
                        .type(ExcelDataType.INTEGER)
                        .group("Financial", "Revenue"))
                .column("Discount", ProductDto::discount, c -> c
                        .type(ExcelDataType.DOUBLE_PERCENT)
                        .group("Financial"))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("group-header-multi-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/comment-size")
    public ResponseEntity<StreamingResponseBody> downloadCommentSize() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Comment Size")
                .autoFilter(true)
                .column("Name", ProductDto::name, cfg -> cfg
                    .headerComment(ExcelCellComment.of("Product name (max 50 chars)")
                            .author("System").size(4, 3))
                    .comment(p -> "Category: " + p.category())
                    .commentSize(5, 4))
                .column("Price", ProductDto::price, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .format("#,##0")
                    .headerComment(ExcelCellComment.of("KRW, no decimals").size(3, 2)))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("comment-size-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/conditional-format")
    public ResponseEntity<StreamingResponseBody> downloadConditionalFormat() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.GOLD)
                .sheetName("Conditional Formatting")
                .autoFilter(true)
                .freezeRows(1)
                .column("Name", ProductDto::name)
                .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                .column("Discount", p -> p.discount(), c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .greaterThan("30000", ExcelColor.LIGHT_RED)
                        .lessThan("5000", ExcelColor.LIGHT_GREEN))
                .conditionalFormatting(cf -> cf
                        .columns(2)
                        .lessThanOrEqual("10", ExcelColor.LIGHT_ORANGE))
                .conditionalFormatting(cf -> cf
                        .columns(3)
                        .greaterThanOrEqual("0.2", ExcelColor.LIGHT_PURPLE))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("conditional-format-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/sheet-protection")
    public ResponseEntity<StreamingResponseBody> downloadSheetProtection() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.CORAL)
                .sheetName("Protected Sheet")
                .autoFilter(true)
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
                .protectSheet("1234")
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("sheet-protection-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/chart")
    public ResponseEntity<StreamingResponseBody> downloadChart() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Chart Demo")
                .column("Name", ProductDto::name)
                .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Product Price vs Quantity")
                        .categoryColumn(0)
                        .valueColumn(1, "Price")
                        .valueColumn(2, "Quantity"))
                .write(sampleProducts().stream().limit(10));

        return ExcelKitResponse.excel("chart-demo")
                .body(handler::writeTo);
    }

    /** A combination chart: columns, a line on the same axis and a line on a second axis (fixed data, checkable by eye). */
    record MonthSales(String month, int revenue, int target, double growth) {}

    @GetMapping("/chart-combo")
    public ResponseEntity<StreamingResponseBody> downloadChartCombo() {
        var months = List.of(new MonthSales("Jan", 100, 110, 0.05), new MonthSales("Feb", 120, 110, 0.20),
                new MonthSales("Mar", 90, 120, -0.25), new MonthSales("Apr", 150, 130, 0.67),
                new MonthSales("May", 170, 150, 0.13), new MonthSales("Jun", 160, 170, -0.06));
        var handler = ExcelWriter.<MonthSales>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Combo Chart")
                .column("Month", MonthSales::month)
                .column("Revenue", MonthSales::revenue, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Target", MonthSales::target, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Growth", MonthSales::growth, c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Revenue, target and growth")
                        .categoryColumn(0)
                        .valueColumn(1, "Revenue")
                        .lineColumn(2, "Target")
                        .secondaryLineColumn(3, "Growth")
                        .valueAxisTitle("Revenue")
                        .secondaryValueAxisTitle("Growth")
                        .legendPosition(ExcelChartConfig.LegendPosition.BOTTOM)
                        .position(5, 1, 14, 18))
                .write(months.stream());

        return ExcelKitResponse.excel("chart-combo-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/empty-message")
    public ResponseEntity<StreamingResponseBody> downloadEmptyMessage() {
        try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
            wb.<ProductDto>sheet("No rows")
                    .autoFilter().freezeRows(1)
                    .emptyMessage("No data")
                    .column("Name", ProductDto::name)
                    .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                    .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Name", "Total").sum("Price"))
                    .write(Stream.<ProductDto>empty());
            wb.<ProductDto>sheet("With rows")
                    .autoFilter().freezeRows(1)
                    .emptyMessage("No data")
                    .column("Name", ProductDto::name)
                    .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                    .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Name", "Total").sum("Price"))
                    .write(sampleProducts().stream().limit(5));
            var handler = wb.finish();
            return ExcelKitResponse.excel("empty-message-demo").body(handler::writeTo);
        }
    }

    @GetMapping("/map-writer")
    public ResponseEntity<StreamingResponseBody> downloadMapWriter() {
        var maps = sampleProducts().stream()
                .<Map<String, Object>>map(p -> {
                    Map<String, Object> m = new LinkedHashMap<>();
                    m.put("Name", p.name());
                    m.put("Category", p.category());
                    m.put("Price", p.price());
                    m.put("Quantity", p.quantity());
                    m.put("Discount", String.format("%.0f%%", p.discount() * 100));
                    return m;
                });

        var handler = ExcelWriter.forMap("Name", "Category", "Price", "Quantity", "Discount")
                .write(maps);

        return ExcelKitResponse.excel("map-writer-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/workbook-protection")
    public ResponseEntity<StreamingResponseBody> downloadWorkbookProtection() {
        try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
            wb.protectWorkbook("secret");

            wb.<ProductDto>sheet("Protected Data")
                    .protectSheet("1234")
                    .autoFilter()
                    .freezeRows(1)
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
            return ExcelKitResponse.excel("workbook-protection-demo")
                    .body(handler::writeTo);
        }
    }

    @GetMapping("/header-font")
    public ResponseEntity<StreamingResponseBody> downloadHeaderFont() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.CORAL)
                .sheetName("Header Font Demo")
                .headerFontName("Arial")
                .headerFontSize(14)
                .autoFilter(true)
                .freezeRows(1)
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                .column("Discount", p -> p.discount(), c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("header-font-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/header-font-color")
    public ResponseEntity<StreamingResponseBody> downloadHeaderFontColor() {
        boolean hasStockAlert = true;

        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Header Font Color Demo")
                .headerFontName("Arial")
                .headerFontSize(13)
                .autoFilter(true)
                .freezeRows(1)
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", p -> p.price(), c -> c
                        .type(ExcelDataType.INTEGER)
                        .format("#,##0")
                        .headerFontColor(ExcelColor.RED))
                .column("Quantity", p -> p.quantity(), c -> c
                        .type(ExcelDataType.INTEGER)
                        .headerFontColor(hasStockAlert ? ExcelColor.RED : null))
                .column("Discount", p -> p.discount(), c -> c
                        .type(ExcelDataType.DOUBLE_PERCENT))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("header-font-color-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/default-style")
    public ResponseEntity<StreamingResponseBody> downloadDefaultStyle() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.FOREST_GREEN)
                .sheetName("Default Style Demo")
                .autoFilter(true)
                .freezeRows(1)
                .defaultStyle(d -> d
                        .fontName("Arial")
                        .fontSize(10)
                        .alignment(HorizontalAlignment.LEFT))
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", p -> p.price(), c -> c
                        .type(ExcelDataType.INTEGER)
                        .format("#,##0")
                        .alignment(HorizontalAlignment.RIGHT))
                .column("Quantity", p -> p.quantity(), c -> c
                        .type(ExcelDataType.INTEGER)
                        .alignment(HorizontalAlignment.RIGHT))
                .column("Discount", p -> p.discount(), c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("default-style-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/summary")
    public ResponseEntity<StreamingResponseBody> downloadSummary() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.GOLD)
                .sheetName("Summary Demo")
                .autoFilter(true)
                .freezeRows(1)
                .column("Name", ProductDto::name)
                .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s
                        .label("Total")
                        .sum("Price").sum("Quantity")
                        .average("Price").average("Quantity"))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("summary-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/named-range")
    public ResponseEntity<StreamingResponseBody> downloadNamedRange() {
        var categories = List.of("Electronics", "Accessories", "Office", "Peripherals");

        try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
            wb.<String>sheet("Options")
                    .column("Category", s -> s)
                    .afterData(ctx -> {
                        ctx.namedRange("CategoryList", 0, 1, categories.size());
                        return ctx.getCurrentRow();
                    })
                    .write(categories.stream());

            wb.<ProductDto>sheet("Data")
                    .autoFilter()
                    .freezeRows(1)
                    .column("Name", ProductDto::name)
                    .column("Category", ProductDto::category, c -> c
                            .validation(ExcelValidation.listFromRange("Options!$A$2:$A$5")))
                    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                    .column("Quantity", p -> p.quantity(), c -> c.type(ExcelDataType.INTEGER))
                    .write(sampleProducts().stream());

            var handler = wb.finish();
            return ExcelKitResponse.excel("named-range-demo")
                    .body(handler::writeTo);
        }
    }

    @GetMapping("/hidden-sheet")
    public ResponseEntity<StreamingResponseBody> downloadHiddenSheet() {
        var categories = List.of("Electronics", "Accessories", "Office", "Peripherals");

        try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
            // The master codes behind the dropdown: first in the file, but hidden (right-click a tab > Unhide)
            wb.<String>sheet("Codes")
                    .column("Category", s -> s)
                    .hideSheet()
                    .write(categories.stream());

            wb.<ProductDto>sheet("Data")
                    .zoom(90)
                    .autoFilter()
                    .freezeRows(1)
                    .column("Name", ProductDto::name)
                    .column("Category", ProductDto::category, c -> c
                            .validation(ExcelValidation.listFromRange("Codes!$A$2:$A$5")))
                    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                    .write(sampleProducts().stream());

            var handler = wb.finish();
            return ExcelKitResponse.excel("hidden-sheet-demo")
                    .body(handler::writeTo);
        }
    }

    @GetMapping("/page-break")
    public ResponseEntity<StreamingResponseBody> downloadPageBreak() {
        // One printed page per category: the list is sorted by category and a page break goes where it changes
        var products = sampleProducts().stream()
                .sorted(java.util.Comparator.comparing(ProductDto::category)).toList();
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Page Break Demo")
                .freezeRows(1)
                .pageBreakOnChange(ProductDto::category)
                .printSetup(ps -> ps.repeatHeaderRows().fitToPageWidth().footerCenter("Page &P of &N"))
                .column("Category", ProductDto::category)
                .column("Name", ProductDto::name)
                .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .write(products);
        return ExcelKitResponse.excel("page-break-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/merge-same-values")
    public ResponseEntity<StreamingResponseBody> downloadMergeSameValues() {
        // Sorted so that equal values are neighbours: the category is shown once over its rows
        var products = sampleProducts().stream()
                .sorted(java.util.Comparator.comparing(ProductDto::category).thenComparing(ProductDto::name)).toList();
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Merge Demo")
                .freezeRows(1)
                .column("Category", ProductDto::category, c -> c.mergeSameValues())
                .column("Name", ProductDto::name)
                .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .write(products);
        return ExcelKitResponse.excel("merge-same-values-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/striped-rows")
    public ResponseEntity<StreamingResponseBody> downloadStripedRows() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Striped Rows")
                .freezeRows(1)
                .stripedRows()
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                .write(sampleProducts().stream());
        return ExcelKitResponse.excel("striped-rows-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/data-bar")
    public ResponseEntity<StreamingResponseBody> downloadDataBar() {
        var writer = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Data Bar Demo")
                .autoFilter(true)
                .freezeRows(1);
        writer.column("Name", ProductDto::name);
        writer.column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"));
        writer.column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER));
        writer.column("Discount", ProductDto::discount, c -> c.type(ExcelDataType.DOUBLE_PERCENT));
        var handler = writer
                .conditionalFormatting(cf -> cf.columns(1).dataBar(ExcelColor.BLUE))
                .conditionalFormatting(cf -> cf.columns(2).dataBar(ExcelColor.GREEN))
                .conditionalFormatting(cf -> cf.columns(3).iconSet(ExcelConditionalRule.IconSetType.ARROWS_3))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("data-bar-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/freeze-cols")
    public ResponseEntity<StreamingResponseBody> downloadFreezeCols() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Freeze Cols Demo")
                .autoFilter(true)
                .freezeCols(2)
                .column("Name", ProductDto::name)
                .column("Category", ProductDto::category)
                .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                .column("Discount", ProductDto::discount, c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("freeze-cols-demo")
                .body(handler::writeTo);
    }

    @GetMapping("/late-password")
    public ResponseEntity<StreamingResponseBody> downloadLateBoundPassword() {
        ExcelHandler handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.CORAL)
                .sheetName("Late Password")
                .autoFilter(true)
                .freezeRows(1)
                .column("Name", ProductDto::name)
                .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .write(sampleProducts().stream());

        String password = "demo123";
        return ExcelKitResponse.excel("late-password-demo")
                .body(os -> handler.writeTo(os, password));
    }

    /**
     * Formats and formulas: ids with leading zeros that stay text when a cell is edited, presets for numbers, percents
     * and dates, a formula that carries its result (it shows in Protected View and previews), and a styled summary row.
     */
    @GetMapping("/formats-and-formulas")
    public ResponseEntity<StreamingResponseBody> downloadFormatsAndFormulas() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Formats and Formulas")
                .freezeRows(1)
                // "0012" is text; in a General cell Excel turns it into the number 12 when the cell is edited
                .column("Code", p -> String.format("%04d", p.price() / 1000), c -> c.format(ExcelDataFormat.TEXT))
                .column("Name", ProductDto::name)
                .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format(ExcelDataFormat.NUMBER))
                .column("Discount", ProductDto::discount, c -> c.type(ExcelDataType.DOUBLE).format(ExcelDataFormat.PERCENT_1))
                .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                // the formula with its result: Protected View and previews show the number; Excel recalculates when it opens the file
                // in a (row, cursor) function cursor.getRowOfSheet() is the row as Excel numbers it, so a formula can point at its own row
                .column("Value (formula + result)", (p, cursor) -> ExcelFormula.of(
                        "C" + cursor.getRowOfSheet() + "*E" + cursor.getRowOfSheet(), (long) p.price() * p.quantity()),
                        c -> c.format(ExcelDataFormat.NUMBER))
                .column("Value (formula only)", (p, cursor) -> ExcelFormula.of(
                        "C" + cursor.getRowOfSheet() + "*E" + cursor.getRowOfSheet()), c -> c.format(ExcelDataFormat.NUMBER))
                .summary(s -> s.label("Name", "Total").sum("Price").sum("Quantity").average("Price")
                        .style(r -> r.backgroundColor(ExcelColor.LIGHT_YELLOW)))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("formats-and-formulas-demo")
                .body(handler::writeTo);
    }

    /** Axis labels of a chart over small whole numbers: the value axis is told to step by 1, not 0.5 (0 1 1 2 2 3 3 4). */
    @GetMapping("/chart-axis")
    public ResponseEntity<StreamingResponseBody> downloadChartAxis() {
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Chart Axis")
                .column("Name", ProductDto::name)
                .column("Quantity", p -> p.quantity() % 4, c -> c.type(ExcelDataType.INTEGER))
                .chart(ch -> ch
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Quantity (axis step 1)")
                        .categoryColumn(0)
                        .valueColumn(1, "Quantity")
                        .valueAxisMajorUnit(1))
                .write(sampleProducts().stream());

        return ExcelKitResponse.excel("chart-axis-demo")
                .body(handler::writeTo);
    }
}
