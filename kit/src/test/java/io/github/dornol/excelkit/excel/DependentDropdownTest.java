package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class DependentDropdownTest {

    private record Row(String category, String item) {}

    private static Map<String, List<String>> options() {
        var map = new LinkedHashMap<String, List<String>>();
        map.put("가", List.of("1", "2", "3"));
        map.put("나", List.of("4", "5"));
        map.put("다", List.of("6"));
        return map;
    }

    private static byte[] write() throws IOException {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create()
                .column("Category", Row::category, c -> c.dropdown("가", "나", "다"))
                .column("Item", Row::item, c -> c.dependentDropdown("Category", options()))
                .write(Stream.of(new Row("가", "1"), new Row("나", "4")))
                .writeTo(out);
        return out.toByteArray();
    }

    @Nested
    class Output {
        @Test
        void writesAHiddenOptionsSheetAndAListValidationOverTheWholeColumn() throws IOException {
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(write()))) {
                assertEquals(2, wb.getNumberOfSheets());
                assertEquals(SheetVisibility.HIDDEN, wb.getSheetVisibility(1));

                var helper = wb.getSheetAt(1);
                assertEquals("가", helper.getRow(0).getCell(0).getStringCellValue());
                assertEquals("나", helper.getRow(0).getCell(1).getStringCellValue());
                assertEquals("다", helper.getRow(0).getCell(2).getStringCellValue());
                assertEquals(3.0, helper.getRow(3).getCell(0).getNumericCellValue(), "a plain-number option is a number");
                assertNull(helper.getRow(3).getCell(1), "a parent with fewer options leaves the rest empty");

                var validations = wb.getSheetAt(0).getDataValidations();
                assertEquals(2, validations.size());
                var dependent = validations.stream()
                        .filter(v -> v.getRegions().getCellRangeAddress(0).getFirstColumn() == 1)
                        .findFirst().orElseThrow();
                assertEquals(1, dependent.getRegions().getCellRangeAddress(0).getFirstRow(), "starts below the header");
                assertTrue(dependent.getValidationConstraint().getFormula1().contains("$A2"),
                        "reads the parent cell of its own row: " + dependent.getValidationConstraint().getFormula1());
            }
        }

        @Test
        void theFormulaOffersTheOptionsOfTheRowsParentValue() throws IOException {
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(write()))) {
                String formula = wb.getSheetAt(0).getDataValidations().stream()
                        .filter(v -> v.getRegions().getCellRangeAddress(0).getFirstColumn() == 1)
                        .findFirst().orElseThrow().getValidationConstraint().getFormula1();
                var sheet = wb.getSheetAt(0);
                var evaluator = wb.getCreationHelper().createFormulaEvaluator();
                // Evaluate the validation formula the way Excel does for a cell of row 2, 3 and 4 (parent in column A)
                var row = sheet.createRow(10);
                for (var expected : Map.of("가", 3, "나", 2, "다", 1).entrySet()) {
                    sheet.getRow(1).createCell(0).setCellValue(expected.getKey()); // a new cell: POI 5 ignores setCellValue on a cell it has read from a file
                    var cell = row.createCell(0);
                    cell.setCellFormula("COUNTA(" + formula + ")");
                    evaluator = wb.getCreationHelper().createFormulaEvaluator(); // a new one: POI 5 keeps the value of a changed cell it has already read
                    var value = evaluator.evaluate(cell);
                    assertEquals(CellType.NUMERIC, value.getCellType());
                    assertEquals(expected.getValue(), (int) value.getNumberValue(), expected.getKey());
                }
            }
        }

        @Test
        void aNumericParentValueFindsItsOptions() throws IOException {
            var numeric = new LinkedHashMap<String, List<String>>();
            numeric.put("1", List.of("a", "b", "c"));
            numeric.put("2", List.of("d"));
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Row>create()
                    .column("Code", Row::category)
                    .column("Item", Row::item, c -> c.dependentDropdown("Code", numeric))
                    .write(Stream.of(new Row("1", "a"))).writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                String formula = wb.getSheetAt(0).getDataValidations().get(0).getValidationConstraint().getFormula1();
                var sheet = wb.getSheetAt(0);
                var evaluator = wb.getCreationHelper().createFormulaEvaluator();
                var row = sheet.createRow(10);
                // the parent cell holds a number, as a numeric code column does
                for (var expected : Map.of(1, 3, 2, 1).entrySet()) {
                    sheet.getRow(1).createCell(0).setCellValue((double) expected.getKey()); // a new cell: POI 5 ignores setCellValue on a cell it has read from a file
                    var cell = row.createCell(0);
                    cell.setCellFormula("COUNTA(" + formula + ")");
                    evaluator = wb.getCreationHelper().createFormulaEvaluator(); // a new one: POI 5 keeps the value of a changed cell it has already read
                    assertEquals(expected.getValue(), (int) evaluator.evaluate(cell).getNumberValue(), "code " + expected.getKey());
                }
            }
        }

        @Test
        void aWriterWithRolloverAndAChartKeepsTheHelperSheetOutOfTheProcessing() throws IOException {
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Row>create().maxRows(2)
                    .column("Category", Row::category)
                    .column("Item", Row::item, c -> c.dependentDropdown("Category", options()))
                    .table("Items")
                    .write(Stream.of(new Row("가", "1"), new Row("나", "4"), new Row("다", "6")))
                    .writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(3, wb.getNumberOfSheets(), "two data sheets share one helper");
                assertEquals(SheetVisibility.HIDDEN, wb.getSheetVisibility(2));
            }
        }

        @Test
        void sheetsOfOneWorkbookShareAHelperForEqualOptionsAndGetOwnOtherwise() throws IOException {
            var out = new ByteArrayOutputStream();
            try (var workbook = ExcelWorkbook.create()) {
                int n = 0;
                for (String name : List.of("A", "B", "C")) {
                    var opts = n++ < 2 ? options() : Map.of("x", List.of("y"));
                    workbook.<Row>sheet(name)
                            .column("Category", Row::category)
                            .column("Item", Row::item, c -> c.dependentDropdown("Category", opts))
                            .write(Stream.of(new Row("가", "1")));
                }
                workbook.finish().writeTo(out);
            }
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(5, wb.getNumberOfSheets(), "three data sheets and two helpers");
                assertNotNull(wb.getSheet("_excelkit_dep_1"));
                assertNotNull(wb.getSheet("_excelkit_dep_2"));
                assertNull(wb.getSheet("_excelkit_dep_3"));
            }
        }
    }

    @Nested
    class Rejections {
        @Test
        void aMissingParentColumnFailsOnWrite() {
            assertThrows(ExcelWriteException.class, () -> ExcelWriter.<Row>create()
                    .column("Item", Row::item, c -> c.dependentDropdown("Nope", options()))
                    .write(Stream.of(new Row("가", "1")))
                    .writeTo(new ByteArrayOutputStream()));
        }

        @Test
        void theColumnItselfIsNotItsParent() {
            assertThrows(ExcelWriteException.class, () -> ExcelWriter.<Row>create()
                    .column("Item", Row::item, c -> c.dependentDropdown("Item", options()))
                    .write(Stream.of(new Row("가", "1")))
                    .writeTo(new ByteArrayOutputStream()));
        }

        @Test
        void aDropdownOnTheSameColumnIsRejected() {
            assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<Row>create()
                    .column("Category", Row::category)
                    .column("Item", Row::item, c -> c.dropdown("x").dependentDropdown("Category", options())));
        }

        @Test
        void aValidationSetAfterTheDependentDropdownReplacesIt() throws IOException {
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Row>create()
                    .column("Category", Row::category)
                    .column("Item", Row::item, c -> c
                            .dependentDropdown("Category", options()).validation(ExcelValidation.textLength(1, 5)))
                    .write(Stream.of(new Row("가", "1"))).writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(1, wb.getNumberOfSheets(), "no helper sheet: the dependent dropdown was replaced");
            }
        }

        @Test
        void invalidArgumentsAreRejected() {
            var c = new ColumnStyleConfig.DefaultStyleConfig<Row>();
            assertThrows(NullPointerException.class, () -> c.dependentDropdown(null, options()));
            assertThrows(NullPointerException.class, () -> c.dependentDropdown("P", null));
            assertThrows(IllegalArgumentException.class, () -> c.dependentDropdown(" ", options()));
            assertThrows(IllegalArgumentException.class, () -> c.dependentDropdown("P", Map.of()));
            assertThrows(IllegalArgumentException.class, () -> c.dependentDropdown("P", Map.of("a", List.of())));
            assertThrows(IllegalArgumentException.class, () -> c.dependentDropdown("P", Map.of("", List.of("x"))));
        }

        @Test
        void moreParentValuesThanColumnsAreRejected() {
            var map = new LinkedHashMap<String, List<String>>();
            for (int i = 0; i < 20_000; i++) map.put("k" + i, List.of("x"));
            var tooMany = assertThrows(IllegalArgumentException.class,
                    () -> new ColumnStyleConfig.DefaultStyleConfig<Row>().dependentDropdown("P", map));
            assertTrue(tooMany.getMessage().contains("columns"));
        }
    }

    @Nested
    class Limits {
        @Test
        void moreOptionsThanASheetHasRowsAreRejectedAndTheLimitIsExact() {
            int max = 1_048_575; // the rows below the header
            var c = new ColumnStyleConfig.DefaultStyleConfig<Row>();
            assertDoesNotThrow(() -> c.dependentDropdown("P", Map.of("a", java.util.Collections.nCopies(max, "x"))));
            var e = assertThrows(IllegalArgumentException.class,
                    () -> c.dependentDropdown("P", Map.of("a", java.util.Collections.nCopies(max + 1, "x"))));
            assertTrue(e.getMessage().contains("options"), e.getMessage());
        }

        @Test
        void aNullOptionOrAnEmptyListIsRejected() {
            var c = new ColumnStyleConfig.DefaultStyleConfig<Row>();
            assertThrows(NullPointerException.class,
                    () -> c.dependentDropdown("P", Map.of("a", java.util.Arrays.asList("x", null))));
            assertThrows(IllegalArgumentException.class, () -> c.dependentDropdown("P", Map.of("a", List.of())));
        }

        @Test
        void theParentValuesKeepTheOrderGiven() throws IOException {
            var ordered = new LinkedHashMap<String, List<String>>();
            for (String key : List.of("z", "m", "a")) ordered.put(key, List.of("1"));
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Row>create()
                    .column("Category", Row::category)
                    .column("Item", Row::item, c -> c.dependentDropdown("Category", ordered))
                    .write(Stream.of(new Row("z", "1"))).writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var header = wb.getSheet("_excelkit_dep_1").getRow(0);
                assertEquals("z", header.getCell(0).getStringCellValue());
                assertEquals("m", header.getCell(1).getStringCellValue());
                assertEquals("a", header.getCell(2).getStringCellValue());
            }
        }
    }

    @Nested
    class OptionTypes {
        @Test
        void optionsThatArePlainNumbersAreWrittenAsNumbersAndTheRestAsText() throws IOException {
            var options = new LinkedHashMap<String, List<String>>();
            options.put("k", List.of("1", "-2", "0", "3.5", "007", "1.50", "+4", "-0", "12abc", "1234567890123456", "x"));
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Row>create()
                    .column("Category", Row::category)
                    .column("Item", Row::item, c -> c.dependentDropdown("Category", options))
                    .write(Stream.of(new Row("k", "1"))).writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var helper = wb.getSheet("_excelkit_dep_1");
                var kinds = new java.util.ArrayList<String>();
                for (int r = 1; r <= 11; r++) kinds.add(helper.getRow(r).getCell(0).getCellType().name());
                assertEquals(List.of("NUMERIC", "NUMERIC", "NUMERIC", "NUMERIC",   // 1, -2, 0, 3.5
                        "STRING", "STRING", "STRING", "STRING", "STRING", "STRING", "STRING"), kinds,
                        "007, 1.50, +4, -0, 12abc, 16 digits and x stay text");
                assertEquals(3.5, helper.getRow(4).getCell(0).getNumericCellValue());
                assertEquals("007", helper.getRow(5).getCell(0).getStringCellValue());
            }
        }

        @Test
        void theParentKeysStayTextWhateverTheyLookLike() throws IOException {
            var options = new LinkedHashMap<String, List<String>>();
            options.put("1", List.of("a"));
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Row>create()
                    .column("Category", Row::category)
                    .column("Item", Row::item, c -> c.dependentDropdown("Category", options))
                    .write(Stream.of(new Row("1", "a"))).writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(org.apache.poi.ss.usermodel.CellType.STRING,
                        wb.getSheet("_excelkit_dep_1").getRow(0).getCell(0).getCellType());
            }
        }

        @Test
        void theNumbersAreStillCountedAsOptions() throws IOException {
            var options = new LinkedHashMap<String, List<String>>();
            options.put("가", List.of("1", "2", "3"));
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Row>create()
                    .column("Category", Row::category)
                    .column("Item", Row::item, c -> c.dependentDropdown("Category", options))
                    .write(Stream.of(new Row("가", "1"))).writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                String formula = wb.getSheetAt(0).getDataValidations().get(0).getValidationConstraint().getFormula1();
                var evaluator = wb.getCreationHelper().createFormulaEvaluator();
                var cell = wb.getSheetAt(0).createRow(10).createCell(0);
                cell.setCellFormula("COUNT(" + formula + ")");   // COUNT counts numbers only
                assertEquals(3, (int) evaluator.evaluate(cell).getNumberValue());
            }
        }
    }
}
