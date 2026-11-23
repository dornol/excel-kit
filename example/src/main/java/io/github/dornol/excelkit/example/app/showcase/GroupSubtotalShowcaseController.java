package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelValidation;
import io.github.dornol.excelkit.excel.ExcelWorkbook;
import io.github.dornol.excelkit.excel.ExcelWriteException;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.spring.ExcelKitResponse;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.IntStream;

/**
 * Group subtotals: {@code groupSubtotal(key, ...)} writes a {@code SUBTOTAL} formula row below every run of rows that
 * share a key. The sales data is fixed (not random) so that the subtotals and totals can be checked by hand.
 */
@Controller
@RequestMapping("/showcase")
public class GroupSubtotalShowcaseController {

    /** One deal. Regions come in blocks of six rows: East, North, South, West (24 rows). */
    record Sale(String region, String rep, int units, int amount) {
        int unitPrice() { return amount / units; }
    }

    private static final String[] REGIONS = {"East", "North", "South", "West"};
    private static final String[] REPS = {"Kim", "Lee", "Park", "Choi", "Jung", "Han"};

    /** Fixed data, sorted by region: East 1000.. , North 2000.., and so on; amount = (region + 1) * 1000 + deal * 250. */
    static List<Sale> sales() {
        List<Sale> rows = new ArrayList<>();
        for (int r = 0; r < REGIONS.length; r++) {
            for (int i = 0; i < REPS.length; i++) {
                rows.add(new Sale(REGIONS[r], REPS[i], 5 + i, (r + 1) * 1000 + i * 250));
            }
        }
        return rows;
    }

    // ========================================================================
    // 1. The basics: a subtotal per region and a grand total that does not count the subtotals twice
    // ========================================================================
    @GetMapping("/subtotal")
    public ResponseEntity<StreamingResponseBody> downloadSubtotal() {
        var handler = ExcelWriter.<Sale>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Sales")
                .freezeRows(1)
                .column("Region", Sale::region)
                .column("Rep", Sale::rep)
                .column("Units", Sale::units, c -> c.type(ExcelDataType.INTEGER))
                .column("Amount", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .groupSubtotal(Sale::region, s -> s
                        .label("Rep", region -> region + " subtotal")
                        .sum("Units").sum("Amount"))
                .summary(s -> s.label("Rep", "Grand total").sum("Units").sum("Amount"))
                .write(sales().stream());
        return ExcelKitResponse.excel("subtotal-demo").body(handler::writeTo);
    }

    // ========================================================================
    // 2. Every function: sum, average, count, min, max (one function per column)
    // ========================================================================
    @GetMapping("/subtotal-functions")
    public ResponseEntity<StreamingResponseBody> downloadSubtotalFunctions() {
        var handler = ExcelWriter.<Sale>create().headerColor(ExcelColor.FOREST_GREEN)
                .sheetName("Functions")
                .freezeRows(1)
                .column("Region", Sale::region)
                .column("Rep", Sale::rep)
                .column("Deals", s -> 1, c -> c.type(ExcelDataType.INTEGER))
                .column("Units", Sale::units, c -> c.type(ExcelDataType.INTEGER))
                .column("Amount", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Unit price", Sale::unitPrice, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Largest", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Smallest", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .groupSubtotal(Sale::region, s -> s
                        .label("Rep", region -> region + " (subtotal)")
                        .count("Deals").sum("Units").sum("Amount")
                        .average("Unit price").max("Largest").min("Smallest"))
                .summary(s -> s.label("Rep", "All regions")
                        .count("Deals").sum("Units").sum("Amount")
                        .average("Unit price").max("Largest").min("Smallest"))
                .write(sales().stream());
        return ExcelKitResponse.excel("subtotal-functions-demo").body(handler::writeTo);
    }

    // ========================================================================
    // 3. Layout: merged region cell, zebra stripes, a page per region, filter, custom subtotal style
    // ========================================================================
    @GetMapping("/subtotal-layout")
    public ResponseEntity<StreamingResponseBody> downloadSubtotalLayout() {
        var handler = ExcelWriter.<Sale>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Layout")
                .autoFilter(true)
                .freezeRows(1)
                .stripedRows()
                .pageBreakOnChange(Sale::region)
                .column("Region", Sale::region, c -> c.mergeSameValues())
                .column("Rep", Sale::rep)
                .column("Units", Sale::units, c -> c.type(ExcelDataType.INTEGER))
                .column("Amount", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .groupSubtotal(Sale::region, s -> s
                        .label("Rep", region -> region + " total")
                        .sum("Units").sum("Amount")
                        .style(st -> st.bold(true).italic(true).fontColor(ExcelColor.STEEL_BLUE)
                                .backgroundColor(ExcelColor.LIGHT_BLUE)))
                .summary(s -> s.label("Rep", "Grand total").sum("Amount"))
                .write(sales().stream());
        return ExcelKitResponse.excel("subtotal-layout-demo").body(handler::writeTo);
    }

    // ========================================================================
    // 4. Rollover: 24 rows over sheets of 10; each sheet closes its group and has its own total
    // ========================================================================
    @GetMapping("/subtotal-rollover")
    public ResponseEntity<StreamingResponseBody> downloadSubtotalRollover() {
        var handler = ExcelWriter.<Sale>create().headerColor(ExcelColor.CORAL)
                .sheetName("Sales")
                .maxRows(10)
                .autoFilter(true)
                .rowNumberColumn("No")
                .column("Region", Sale::region)
                .column("Rep", Sale::rep)
                .column("Amount", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .groupSubtotal(Sale::region, s -> s
                        .label("Rep", region -> region + " (this sheet)")
                        .sum("Amount"))
                .summary(s -> s.label("Rep", "Sheet total").sum("Amount"))
                .write(sales().stream());
        return ExcelKitResponse.excel("subtotal-rollover-demo").body(handler::writeTo);
    }

    // ========================================================================
    // 5. Sheet protection: the data columns can be edited, the subtotal rows stay locked (password: 1234)
    // ========================================================================
    @GetMapping("/subtotal-protection")
    public ResponseEntity<StreamingResponseBody> downloadSubtotalProtection() {
        var handler = ExcelWriter.<Sale>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Protected")
                .protectSheet("1234")
                .column("Region", Sale::region)
                .column("Rep", Sale::rep, c -> c.locked(false))
                .column("Units", Sale::units, c -> c.type(ExcelDataType.INTEGER).locked(false))
                .column("Amount", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0").locked(false))
                .groupSubtotal(Sale::region, s -> s
                        .label("Rep", region -> region + " subtotal")
                        .sum("Units").sum("Amount"))
                .summary(s -> s.label("Rep", "Grand total").sum("Units").sum("Amount"))
                .write(sales().stream());
        return ExcelKitResponse.excel("subtotal-protection-demo").body(handler::writeTo);
    }

    // ========================================================================
    // 6. Rules: a conditional format, a dropdown and a validation run over the subtotal rows too
    // ========================================================================
    @GetMapping("/subtotal-rules")
    public ResponseEntity<StreamingResponseBody> downloadSubtotalRules() {
        var handler = ExcelWriter.<Sale>create().headerColor(ExcelColor.STEEL_BLUE)
                .sheetName("Rules")
                .autoFilter(true)
                .column("Region", Sale::region, c -> c.dropdown(REGIONS))
                .column("Rep", Sale::rep)
                .column("Units", Sale::units, c -> c.type(ExcelDataType.INTEGER)
                        .validation(ExcelValidation.integerBetween(0, 500)))
                .column("Amount", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .conditionalFormatting(c -> c.columns(3).greaterThan("2500", ExcelColor.LIGHT_YELLOW))
                .groupSubtotal(Sale::region, s -> s
                        .label("Rep", region -> region + " subtotal")
                        .sum("Units").sum("Amount"))
                .summary(s -> s.label("Rep", "Grand total").sum("Units").sum("Amount"))
                .write(sales().stream());
        return ExcelKitResponse.excel("subtotal-rules-demo").body(handler::writeTo);
    }

    // ========================================================================
    // 7. A workbook: one sheet with subtotals next to a plain one
    // ========================================================================
    @GetMapping("/subtotal-workbook")
    public ResponseEntity<StreamingResponseBody> downloadSubtotalWorkbook() {
        try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
            wb.<Sale>sheet("By region")
                    .freezeRows(1)
                    .column("Region", Sale::region)
                    .column("Rep", Sale::rep)
                    .column("Amount", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                    .groupSubtotal(Sale::region, s -> s.label("Rep", region -> region + " subtotal").sum("Amount"))
                    .summary(s -> s.label("Rep", "Total").sum("Amount"))
                    .write(sales().stream());

            wb.<Sale>sheet("Plain list")
                    .freezeRows(1)
                    .column("Region", Sale::region)
                    .column("Rep", Sale::rep)
                    .column("Amount", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                    .summary(s -> s.label("Rep", "Total").sum("Amount"))
                    .write(sales().stream());

            var handler = wb.finish();
            return ExcelKitResponse.excel("subtotal-workbook-demo").body(handler::writeTo);
        }
    }

    // ========================================================================
    // 8. Random product data, subtotal per category (sorted by category first)
    // ========================================================================
    @GetMapping("/subtotal-products")
    public ResponseEntity<StreamingResponseBody> downloadSubtotalProducts() {
        var products = ShowcaseData.sampleProducts().stream()
                .sorted(Comparator.comparing(ProductDto::category).thenComparing(ProductDto::name))
                .toList();
        var handler = ExcelWriter.<ProductDto>create().headerColor(ExcelColor.FOREST_GREEN)
                .sheetName("Products")
                .freezeRows(1)
                .column("Category", ProductDto::category)
                .column("Name", ProductDto::name)
                .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Quantity", ProductDto::quantity, c -> c.type(ExcelDataType.INTEGER))
                .groupSubtotal(ProductDto::category, s -> s
                        .label("Name", category -> category + " subtotal")
                        .sum("Price").sum("Quantity"))
                .summary(s -> s.label("Name", "Total").sum("Price").sum("Quantity"))
                .write(products.stream());
        return ExcelKitResponse.excel("subtotal-products-demo").body(handler::writeTo);
    }

    // ========================================================================
    // 9. The contract: rows must be sorted by the key (here they are not: a key that returns gets a subtotal of its own)
    // ========================================================================
    @GetMapping("/subtotal-unsorted")
    public ResponseEntity<StreamingResponseBody> downloadSubtotalUnsorted() {
        // East, North, East, North...: every row is its own group
        var rows = IntStream.range(0, 12).mapToObj(i -> sales().get((i % 2) * 6 + i / 2)).toList();
        var handler = ExcelWriter.<Sale>create().headerColor(ExcelColor.ORANGE)
                .sheetName("Unsorted")
                .column("Region", Sale::region)
                .column("Rep", Sale::rep)
                .column("Amount", Sale::amount, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .groupSubtotal(Sale::region, s -> s.label("Rep", region -> region + " subtotal").sum("Amount"))
                .write(rows.stream());
        return ExcelKitResponse.excel("subtotal-unsorted-demo").body(handler::writeTo);
    }

    // ========================================================================
    // 10. What is refused: a table, a chart or a named range cannot be combined with subtotals
    // ========================================================================
    @GetMapping("/subtotal-rejected")
    @ResponseBody
    public ResponseEntity<String> subtotalRejected() {
        try {
            ExcelWriter.<Sale>create()
                    .column("Region", Sale::region)
                    .column("Amount", Sale::amount)
                    .table("Sales")
                    .groupSubtotal(Sale::region, s -> s.sum("Amount"))
                    .write(sales().stream());
            return ResponseEntity.ok("Unexpected: the write succeeded");
        } catch (ExcelWriteException e) {
            return ResponseEntity.badRequest().body("table(...) + groupSubtotal(...) is refused: " + e.getMessage());
        }
    }
}
