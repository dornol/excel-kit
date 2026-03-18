package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ExcelKitSchema;
import io.github.dornol.excelkit.shared.ReadResult;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Round-trip integration tests: write Excel/CSV with excel-kit, read back, and verify.
 * Covers features from v0.7.1 through v0.8.x.
 */
class RoundTripIntegrationTest {

    // ========================================================================
    // 1. Vertical Alignment / WrapText / FontName / Indentation (v0.7.1)
    // ========================================================================
    @Nested
    class VerticalAlignmentWrapTextFontNameIndentation {

        record StyledRow(String label, String description) {}

        @Test
        void shouldPreserveVerticalAlignmentWrapTextFontNameAndIndentation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<StyledRow>()
                    .column("Label", StyledRow::label)
                        .verticalAlignment(VerticalAlignment.TOP)
                        .wrapText(false)
                        .fontName("Arial")
                        .indentation(2)
                    .column("Description", StyledRow::description)
                        .verticalAlignment(VerticalAlignment.BOTTOM)
                        .wrapText(true)
                        .fontName("Courier New")
                        .indentation(0)
                    .write(Stream.of(
                            new StyledRow("Item A", "First description"),
                            new StyledRow("Item B", "Second description")))
                    .consumeOutputStream(out);

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                Row dataRow = sheet.getRow(1);

                // Column 0: Label
                CellStyle style0 = dataRow.getCell(0).getCellStyle();
                assertEquals(VerticalAlignment.TOP, style0.getVerticalAlignment(),
                        "Label column should have TOP vertical alignment");
                assertFalse(style0.getWrapText(), "Label column should have wrapText=false");
                XSSFFont font0 = wb.getFontAt(style0.getFontIndex());
                assertEquals("Arial", font0.getFontName(), "Label column should use Arial font");
                assertEquals(2, style0.getIndention(), "Label column should have indentation=2");

                // Column 1: Description
                CellStyle style1 = dataRow.getCell(1).getCellStyle();
                assertEquals(VerticalAlignment.BOTTOM, style1.getVerticalAlignment(),
                        "Description column should have BOTTOM vertical alignment");
                assertTrue(style1.getWrapText(), "Description column should have wrapText=true");
                XSSFFont font1 = wb.getFontAt(style1.getFontIndex());
                assertEquals("Courier New", font1.getFontName(), "Description column should use Courier New font");
                assertEquals(0, style1.getIndention(), "Description column should have indentation=0");
            }
        }
    }

    // ========================================================================
    // 2. Workbook Protection (v0.7.2)
    // ========================================================================
    @Nested
    class WorkbookProtection {

        record SimpleRow(String value) {}

        @Test
        void shouldPreserveWorkbookAndSheetProtection() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<SimpleRow>()
                    .protectWorkbook("wbPass")
                    .protectSheet("sheetPass")
                    .column("Value", SimpleRow::value)
                    .write(Stream.of(new SimpleRow("data")))
                    .consumeOutputStream(out);

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertTrue(wb.isStructureLocked(), "Workbook structure should be locked");
                assertTrue(wb.getSheetAt(0).isSheetLocked(), "Sheet should be locked");
            }
        }
    }

    // ========================================================================
    // 3. Header Font (v0.7.2)
    // ========================================================================
    @Nested
    class HeaderFont {

        record DataRow(String name, int qty) {}

        @Test
        void shouldPreserveHeaderFontNameAndSize() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<DataRow>()
                    .headerFontName("Courier New")
                    .headerFontSize(16)
                    .column("Name", DataRow::name)
                    .column("Qty", d -> d.qty()).type(ExcelDataType.INTEGER)
                    .write(Stream.of(new DataRow("Widget", 10)))
                    .consumeOutputStream(out);

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Row headerRow = wb.getSheetAt(0).getRow(0);
                CellStyle headerStyle = headerRow.getCell(0).getCellStyle();
                XSSFFont headerFont = wb.getFontAt(headerStyle.getFontIndex());

                assertEquals("Courier New", headerFont.getFontName(),
                        "Header font name should be Courier New");
                assertEquals(16, headerFont.getFontHeightInPoints(),
                        "Header font size should be 16pt");
            }
        }
    }

    // ========================================================================
    // 4. Default Style (v0.7.2)
    // ========================================================================
    @Nested
    class DefaultStyle {

        record TwoCol(String col1, String col2) {}

        @Test
        void shouldApplyDefaultStyleWithPerColumnOverride() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<TwoCol>()
                    .defaultStyle(d -> d
                            .fontName("Arial")
                            .bold(true)
                            .alignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT))
                    .column("Col1", TwoCol::col1)
                    .column("Col2", TwoCol::col2)
                        .bold(false)
                    .write(Stream.of(new TwoCol("A", "B")))
                    .consumeOutputStream(out);

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Row dataRow = wb.getSheetAt(0).getRow(1);

                // Col1: inherits defaults (Arial, bold, LEFT)
                CellStyle style1 = dataRow.getCell(0).getCellStyle();
                XSSFFont font1 = wb.getFontAt(style1.getFontIndex());
                assertEquals("Arial", font1.getFontName(), "Col1 should use Arial from default");
                assertTrue(font1.getBold(), "Col1 should be bold from default");
                assertEquals(HorizontalAlignment.LEFT, style1.getAlignment(),
                        "Col1 should be LEFT aligned from default");

                // Col2: overrides bold=false, inherits Arial + LEFT
                CellStyle style2 = dataRow.getCell(1).getCellStyle();
                XSSFFont font2 = wb.getFontAt(style2.getFontIndex());
                assertEquals("Arial", font2.getFontName(), "Col2 should use Arial from default");
                assertFalse(font2.getBold(), "Col2 should override bold to false");
                assertEquals(HorizontalAlignment.LEFT, style2.getAlignment(),
                        "Col2 should be LEFT aligned from default");
            }
        }
    }

    // ========================================================================
    // 5. Summary Rows (v0.7.2)
    // ========================================================================
    @Nested
    class SummaryRows {

        record SaleRow(String item, int price, int qty) {}

        @Test
        void shouldWriteSummaryRowsWithFormulas() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<SaleRow>()
                    .addColumn("Item", SaleRow::item)
                    .addColumn("Price", s -> s.price(), c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Qty", s -> s.qty(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Total").sum("Price").sum("Qty"))
                    .write(Stream.of(
                            new SaleRow("Apple", 100, 5),
                            new SaleRow("Banana", 200, 3),
                            new SaleRow("Cherry", 300, 7)))
                    .consumeOutputStream(out);

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                // Header at row 0, data rows at 1-3, summary at row 4
                Row summaryRow = sheet.getRow(4);
                assertNotNull(summaryRow, "Summary row should exist at row 4");

                // Label in column 0
                assertEquals("Total", summaryRow.getCell(0).getStringCellValue(),
                        "Summary label should be 'Total'");

                // Price formula in column 1 (SUM of B2:B4)
                Cell priceCell = summaryRow.getCell(1);
                assertEquals(CellType.FORMULA, priceCell.getCellType(),
                        "Price summary should be a formula");
                assertEquals("SUM(B2:B4)", priceCell.getCellFormula(),
                        "Price formula should sum B2:B4");

                // Qty formula in column 2 (SUM of C2:C4)
                Cell qtyCell = summaryRow.getCell(2);
                assertEquals(CellType.FORMULA, qtyCell.getCellType(),
                        "Qty summary should be a formula");
                assertEquals("SUM(C2:C4)", qtyCell.getCellFormula(),
                        "Qty formula should sum C2:C4");

                // Evaluate formulas and verify values
                FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                assertEquals(600.0, evaluator.evaluate(priceCell).getNumberValue(), 0.01,
                        "SUM of Price should be 600");
                assertEquals(15.0, evaluator.evaluate(qtyCell).getNumberValue(), 0.01,
                        "SUM of Qty should be 15");
            }
        }
    }

    // ========================================================================
    // 6. Named Range (v0.7.2)
    // ========================================================================
    @Nested
    class NamedRange {

        record Category(String name) {}

        @Test
        void shouldCreateNamedRangeViaAfterData() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Category>()
                    .sheetName("Categories")
                    .column("Name", Category::name)
                    .afterData(ctx -> {
                        // Create a named range covering the data in column A
                        // headerRowIndex=0, data starts at row 1
                        ctx.namedRange("CategoryList", 0, 1, ctx.getCurrentRow() - 1);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of(
                            new Category("Electronics"),
                            new Category("Books"),
                            new Category("Clothing")))
                    .consumeOutputStream(out);

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Name namedRange = wb.getName("CategoryList");
                assertNotNull(namedRange, "Named range 'CategoryList' should exist");
                String ref = namedRange.getRefersToFormula();
                assertNotNull(ref, "Named range reference should not be null");
                assertTrue(ref.contains("Categories"), "Reference should contain sheet name");
                assertTrue(ref.contains("$A$"), "Reference should contain column A");
            }
        }
    }

    // ========================================================================
    // 7. List Validation from Range (v0.7.2)
    // ========================================================================
    @Nested
    class ListValidationFromRange {

        record Option(String value) {}
        record DataEntry(String status) {}

        @Test
        void shouldCreateListValidationReferencingAnotherSheet() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook workbook = new ExcelWorkbook()) {
                // Sheet1: Options
                workbook.<Option>sheet("Options")
                        .column("Status", Option::value)
                        .write(Stream.of(
                                new Option("Active"),
                                new Option("Inactive"),
                                new Option("Pending")));

                // Sheet2: Data with listFromRange validation pointing to Options sheet
                workbook.<DataEntry>sheet("Data")
                        .column("Status", DataEntry::status, c ->
                                c.validation(ExcelValidation.listFromRange("Options!$A$2:$A$4")))
                        .write(Stream.of(new DataEntry("Active")));

                workbook.finish().consumeOutputStream(out);
            }

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet dataSheet = wb.getSheet("Data");
                assertNotNull(dataSheet, "Data sheet should exist");

                List<? extends DataValidation> validations = dataSheet.getDataValidations();
                assertFalse(validations.isEmpty(), "Data sheet should have validations");

                boolean hasListValidation = validations.stream()
                        .anyMatch(v -> v.getValidationConstraint().getValidationType()
                                == DataValidationConstraint.ValidationType.LIST);
                assertTrue(hasListValidation, "Data sheet should have a list validation");
            }
        }
    }

    // ========================================================================
    // 8. Mapping Mode with CellData.as() (v0.8.x)
    // ========================================================================
    @Nested
    class MappingModeWithCellDataAs {

        record TypedRow(String name, int price, double weight, UUID id) {}

        @Test
        void shouldWriteAndReadBackTypedDataWithMappingMode() throws IOException {
            UUID id1 = UUID.randomUUID();
            UUID id2 = UUID.randomUUID();

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<TypedRow>()
                    .column("Name", TypedRow::name)
                    .column("Price", r -> r.price()).type(ExcelDataType.INTEGER)
                    .column("Weight", r -> r.weight()).type(ExcelDataType.DOUBLE)
                    .column("Id", r -> r.id().toString())
                    .write(Stream.of(
                            new TypedRow("Alpha", 1500, 2.75, id1),
                            new TypedRow("Beta", 2500, 4.50, id2)))
                    .consumeOutputStream(out);

            List<TypedRow> results = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
                ExcelReader.<TypedRow>mapping(row -> new TypedRow(
                        row.get("Name").asString("Unknown"),
                        row.get("Price").asInt(0),
                        row.get("Weight").asDouble(0.0),
                        row.get("Id").as(UUID::fromString)
                )).build(is).read(r -> {
                    assertTrue(r.success(), "Row should be read successfully");
                    results.add(r.data());
                });
            }

            assertEquals(2, results.size());

            assertEquals("Alpha", results.get(0).name());
            assertEquals(1500, results.get(0).price());
            assertEquals(2.75, results.get(0).weight(), 0.01);
            assertEquals(id1, results.get(0).id());

            assertEquals("Beta", results.get(1).name());
            assertEquals(2500, results.get(1).price());
            assertEquals(4.50, results.get(1).weight(), 0.01);
            assertEquals(id2, results.get(1).id());
        }
    }

    // ========================================================================
    // 9. Mapping Mode with asOrDefault on Empty Cells (v0.8.x)
    // ========================================================================
    @Nested
    class MappingModeWithDefaults {

        record SparseRow(String name, int qty, double price) {}

        @Test
        void shouldApplyDefaultsForEmptyAndNullCells() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<SparseRow>()
                    .column("Name", SparseRow::name)
                    .column("Qty", r -> r.qty() == 0 ? null : r.qty()).type(ExcelDataType.INTEGER)
                    .column("Price", r -> r.price() == 0.0 ? null : r.price()).type(ExcelDataType.DOUBLE)
                    .write(Stream.of(
                            new SparseRow("Full", 10, 99.99),
                            new SparseRow("", 0, 0.0),
                            new SparseRow("Partial", 5, 0.0)))
                    .consumeOutputStream(out);

            List<SparseRow> results = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
                ExcelReader.<SparseRow>mapping(row -> new SparseRow(
                        row.get("Name").asString("N/A"),
                        row.get("Qty").asInt(-1),
                        row.get("Price").asDouble(-1.0)
                )).build(is).read(r -> {
                    assertTrue(r.success());
                    results.add(r.data());
                });
            }

            assertEquals(3, results.size());

            // Row 1: all values present
            assertEquals("Full", results.get(0).name());
            assertEquals(10, results.get(0).qty());
            assertEquals(99.99, results.get(0).price(), 0.01);

            // Row 2: name empty -> default, qty/price null -> defaults
            assertEquals("N/A", results.get(1).name());
            assertEquals(-1, results.get(1).qty());
            assertEquals(-1.0, results.get(1).price(), 0.01);

            // Row 3: name present, qty present, price null -> default
            assertEquals("Partial", results.get(2).name());
            assertEquals(5, results.get(2).qty());
            assertEquals(-1.0, results.get(2).price(), 0.01);
        }
    }

    // ========================================================================
    // 10. CSV Injection Defense Toggle (v0.8.x)
    // ========================================================================
    @Nested
    class CsvInjectionDefense {

        record DangerousRow(String value) {}

        @Test
        void shouldPreserveValuesWhenDefenseIsOff() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<DangerousRow>()
                    .csvInjectionDefense(false)
                    .bom(false)
                    .column("Value", DangerousRow::value)
                    .write(Stream.of(
                            new DangerousRow("=SUM(A1)"),
                            new DangerousRow("-100"),
                            new DangerousRow("+44")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8);

            // With defense OFF, values should be preserved as-is
            assertTrue(csv.contains("=SUM(A1)"), "=SUM(A1) should be preserved without prefix");
            assertTrue(csv.contains("-100"), "-100 should be preserved");
            assertTrue(csv.contains("+44"), "+44 should be preserved");
            assertFalse(csv.contains("'=SUM(A1)"), "Should NOT have quote prefix on =SUM(A1)");
            assertFalse(csv.contains("'+44"), "Should NOT have quote prefix on +44");

            // Read back with CsvReader and verify
            List<DangerousRow> results = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
                CsvReader.<DangerousRow>mapping(row -> new DangerousRow(
                        row.get("Value").asString()
                )).build(is).read(r -> {
                    assertTrue(r.success());
                    results.add(r.data());
                });
            }

            assertEquals(3, results.size());
            assertEquals("=SUM(A1)", results.get(0).value());
            assertEquals("-100", results.get(1).value());
            assertEquals("+44", results.get(2).value());
        }

        @Test
        void shouldAddQuotePrefixWhenDefenseIsOn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<DangerousRow>()
                    .csvInjectionDefense(true)
                    .bom(false)
                    .column("Value", DangerousRow::value)
                    .write(Stream.of(
                            new DangerousRow("=SUM(A1)"),
                            new DangerousRow("-100"),
                            new DangerousRow("+44")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8);

            // With defense ON, dangerous values should be prefixed with single quote
            assertTrue(csv.contains("'=SUM(A1)"), "=SUM(A1) should be prefixed with quote");
            assertTrue(csv.contains("'-100"), "-100 should be prefixed with quote");
            assertTrue(csv.contains("'+44"), "+44 should be prefixed with quote");

            // Read back: values will have the quote prefix
            List<DangerousRow> results = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
                CsvReader.<DangerousRow>mapping(row -> new DangerousRow(
                        row.get("Value").asString()
                )).build(is).read(r -> {
                    assertTrue(r.success());
                    results.add(r.data());
                });
            }

            assertEquals(3, results.size());
            assertEquals("'=SUM(A1)", results.get(0).value(), "Defended value should have quote prefix");
            assertEquals("'-100", results.get(1).value(), "Defended value should have quote prefix");
            assertEquals("'+44", results.get(2).value(), "Defended value should have quote prefix");
        }
    }

    // ========================================================================
    // 11. ExcelWorkbook Multi-Sheet with Mixed Features (v0.7.2)
    // ========================================================================
    @Nested
    class MultiSheetMixedFeatures {

        record SaleItem(String name, int amount) {}
        record ScoreEntry(String student, int score) {}

        @Test
        void shouldWriteMultiSheetWithSummaryAndConditionalFormatting() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook workbook = new ExcelWorkbook()) {
                // Sheet 1: Sales with summary
                workbook.<SaleItem>sheet("Sales")
                        .column("Name", SaleItem::name)
                        .column("Amount", s -> s.amount(), c -> c.type(ExcelDataType.INTEGER))
                        .summary(s -> s.label("Total").sum("Amount"))
                        .write(Stream.of(
                                new SaleItem("A", 100),
                                new SaleItem("B", 200),
                                new SaleItem("C", 300)));

                // Sheet 2: Scores with conditional formatting
                workbook.<ScoreEntry>sheet("Scores")
                        .column("Student", ScoreEntry::student)
                        .column("Score", s -> s.score(), c -> c.type(ExcelDataType.INTEGER))
                        .conditionalFormatting(cf -> cf
                                .columns(1)
                                .greaterThanOrEqual("90", ExcelColor.LIGHT_GREEN)
                                .lessThan("50", ExcelColor.LIGHT_RED))
                        .write(Stream.of(
                                new ScoreEntry("Alice", 95),
                                new ScoreEntry("Bob", 42),
                                new ScoreEntry("Charlie", 78)));

                workbook.finish().consumeOutputStream(out);
            }

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                // Verify Sheet 1: Sales
                Sheet salesSheet = wb.getSheet("Sales");
                assertNotNull(salesSheet, "Sales sheet should exist");

                // Read data back via ExcelReader
                List<SaleItem> sales = new ArrayList<>();
                try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
                    ExcelReader.<SaleItem>mapping(row -> new SaleItem(
                            row.get("Name").asString(),
                            row.get("Amount").asInt(0)
                    )).sheetIndex(0).build(is).read(r -> {
                        if (r.success() && r.data() != null) {
                            sales.add(r.data());
                        }
                    });
                }
                assertTrue(sales.size() >= 3, "Should read at least 3 sale items");
                assertEquals("A", sales.get(0).name());
                assertEquals(100, sales.get(0).amount());

                // Verify summary formula on Sales sheet
                Row summaryRow = salesSheet.getRow(4);
                assertNotNull(summaryRow, "Summary row should exist at row 4");
                Cell amountSummary = summaryRow.getCell(1);
                assertEquals(CellType.FORMULA, amountSummary.getCellType());
                assertEquals("SUM(B2:B4)", amountSummary.getCellFormula());

                // Verify Sheet 2: Scores
                Sheet scoresSheet = wb.getSheet("Scores");
                assertNotNull(scoresSheet, "Scores sheet should exist");

                // Verify conditional formatting exists on Scores sheet
                int cfCount = scoresSheet.getSheetConditionalFormatting().getNumConditionalFormattings();
                assertTrue(cfCount > 0, "Scores sheet should have conditional formatting rules");

                // Read Scores data back
                List<ScoreEntry> scores = new ArrayList<>();
                try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
                    ExcelReader.<ScoreEntry>mapping(row -> new ScoreEntry(
                            row.get("Student").asString(),
                            row.get("Score").asInt(0)
                    )).sheetIndex(1).build(is).read(r -> {
                        if (r.success() && r.data() != null) {
                            scores.add(r.data());
                        }
                    });
                }
                assertEquals(3, scores.size());
                assertEquals("Alice", scores.get(0).student());
                assertEquals(95, scores.get(0).score());
                assertEquals("Bob", scores.get(1).student());
                assertEquals(42, scores.get(1).score());
            }
        }
    }

    // ========================================================================
    // 12. ExcelKitSchema Mapping Mode Round-Trip
    // ========================================================================
    @Nested
    class SchemaRoundTrip {

        static class Product {
            String name;
            int price;
            boolean active;

            Product() {}

            Product(String name, int price, boolean active) {
                this.name = name;
                this.price = price;
                this.active = active;
            }
        }

        @Test
        void shouldWriteWithSchemaAndReadBackWithMappingMode() throws IOException {
            ExcelKitSchema<Product> schema = ExcelKitSchema.<Product>builder()
                    .column("Name", p -> p.name, (p, cell) -> p.name = cell.asString())
                    .column("Price", p -> p.price, (p, cell) -> p.price = cell.asInt(),
                            c -> c.type(ExcelDataType.INTEGER))
                    .column("Active", p -> p.active, (p, cell) -> p.active = "Y".equals(cell.asString()),
                            c -> c.type(ExcelDataType.BOOLEAN_TO_YN))
                    .build();

            // Write using schema
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            schema.excelWriter()
                    .write(Stream.of(
                            new Product("Widget", 1000, true),
                            new Product("Gadget", 2500, false),
                            new Product("Doohickey", 750, true)))
                    .consumeOutputStream(out);

            // Read back using schema's mapping mode reader
            List<Product> results = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
                schema.excelReader(
                        (row) -> {
                            Product p = new Product();
                            p.name = row.get("Name").asString();
                            p.price = row.get("Price").asInt(0);
                            p.active = "Y".equals(row.get("Active").asString());
                            return p;
                        },
                        null
                ).build(is).read(r -> {
                    assertTrue(r.success(), "Row should be read successfully");
                    results.add(r.data());
                });
            }

            assertEquals(3, results.size());

            assertEquals("Widget", results.get(0).name);
            assertEquals(1000, results.get(0).price);
            assertTrue(results.get(0).active);

            assertEquals("Gadget", results.get(1).name);
            assertEquals(2500, results.get(1).price);
            assertFalse(results.get(1).active);

            assertEquals("Doohickey", results.get(2).name);
            assertEquals(750, results.get(2).price);
            assertTrue(results.get(2).active);
        }

        @Test
        void shouldWriteWithSchemaAndReadBackWithSetterMode() throws IOException {
            ExcelKitSchema<Product> schema = ExcelKitSchema.<Product>builder()
                    .column("Name", p -> p.name, (p, cell) -> p.name = cell.asString())
                    .column("Price", p -> p.price, (p, cell) -> p.price = cell.asInt(),
                            c -> c.type(ExcelDataType.INTEGER))
                    .column("Active", p -> p.active, (p, cell) -> p.active = "Y".equals(cell.asString()),
                            c -> c.type(ExcelDataType.BOOLEAN_TO_YN))
                    .build();

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            schema.excelWriter()
                    .write(Stream.of(
                            new Product("Alpha", 500, true),
                            new Product("Bravo", 1200, false)))
                    .consumeOutputStream(out);

            // Read back using schema's setter mode reader
            List<Product> results = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
                schema.excelReader(Product::new, null)
                        .build(is)
                        .read(r -> {
                            assertTrue(r.success(), "Row should be read successfully");
                            results.add(r.data());
                        });
            }

            assertEquals(2, results.size());

            assertEquals("Alpha", results.get(0).name);
            assertEquals(500, results.get(0).price);
            assertTrue(results.get(0).active);

            assertEquals("Bravo", results.get(1).name);
            assertEquals(1200, results.get(1).price);
            assertFalse(results.get(1).active);
        }
    }
}
