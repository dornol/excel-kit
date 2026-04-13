package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.core.CellData;
import io.github.dornol.excelkit.core.ReadResult;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for v0.11.0 additions that otherwise have only happy-path coverage:
 * <ul>
 *   <li>{@code ExcelWriter.forMap(String[], Consumer...)} — edge cases around the
 *       configurers array (length validation, partial configurer arrays)</li>
 *   <li>{@code ExcelWriter.forMap(String...)} — verifying the returned writer is
 *       a full ExcelWriter (not a wrapper) and exposes all fluent methods</li>
 *   <li>{@code CsvWriter.forMap(String...)} — same invariant</li>
 *   <li>Reader {@code column(...)} — verifying the new return type is the Reader
 *       itself (so chaining works the way the Writer side has worked since v0.10.0)</li>
 * </ul>
 */
class ForMapAndReaderColumnTest {

    @Nested
    @DisplayName("ExcelWriter.forMap with configurers")
    class ExcelForMapConfigurers {

        @Test
        @DisplayName("configurers longer than columnNames throws IllegalArgumentException")
        void configurers_tooMany_throws() {
            String[] names = {"Name"};
            @SuppressWarnings("unchecked")
            Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>[] cfgs =
                    new Consumer[]{
                            (Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>) c -> c.bold(true),
                            (Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>) c -> c.bold(true)
                    };
            assertThrows(IllegalArgumentException.class,
                    () -> ExcelWriter.forMap(names, cfgs));
        }

        @Test
        @DisplayName("configurers shorter than columnNames: extra columns use no configurer")
        void configurers_partialArray_extraColumnsPlain() throws IOException {
            // 3 columns, only 1 configurer → columns 2 and 3 get null configurer
            String[] names = {"Name", "Age", "City"};
            @SuppressWarnings("unchecked")
            Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>[] cfgs =
                    new Consumer[]{
                            (Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>) c -> c.bold(true)
                    };

            var out = new ByteArrayOutputStream();
            ExcelWriter.forMap(names, cfgs)
                    .write(Stream.of(Map.of("Name", "Alice", "Age", 30, "City", "Seoul")))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("Age", sheet.getRow(0).getCell(1).getStringCellValue());
                assertEquals("City", sheet.getRow(0).getCell(2).getStringCellValue());
                assertEquals("Alice", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("Seoul", sheet.getRow(1).getCell(2).getStringCellValue());
            }
        }

        @Test
        @DisplayName("configurers exactly matching columnNames: all columns configured")
        void configurers_exactLength_allConfigured() throws IOException {
            String[] names = {"Name", "Age"};
            @SuppressWarnings("unchecked")
            Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>[] cfgs =
                    new Consumer[]{
                            (Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>) c -> c.bold(true),
                            (Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>) c -> c.type(ExcelDataType.INTEGER)
                    };

            var out = new ByteArrayOutputStream();
            ExcelWriter.forMap(names, cfgs)
                    .write(Stream.of(Map.of("Name", "Alice", "Age", 30)))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Integer column: Age should be numeric 30, not string
                assertEquals(30.0, sheet.getRow(1).getCell(1).getNumericCellValue());
            }
        }

        @Test
        @DisplayName("empty configurers with non-empty columnNames: all columns plain")
        void emptyConfigurers_allPlain() throws IOException {
            String[] names = {"A", "B"};
            @SuppressWarnings("unchecked")
            Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>[] cfgs = new Consumer[]{};

            var out = new ByteArrayOutputStream();
            ExcelWriter.forMap(names, cfgs)
                    .write(Stream.of(Map.of("A", "1", "B", "2")))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("A", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("B", sheet.getRow(0).getCell(1).getStringCellValue());
                assertEquals("1", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("2", sheet.getRow(1).getCell(1).getStringCellValue());
            }
        }
    }

    @Nested
    @DisplayName("ExcelWriter.forMap returns a full ExcelWriter")
    class ExcelForMapReturnsFullWriter {

        @Test
        @DisplayName("fluent configuration (sheetName, rowHeight, autoFilter) is available on forMap result")
        void fluentMethodsAvailable() throws IOException {
            var out = new ByteArrayOutputStream();
            ExcelWriter.forMap("Name", "Age")
                    .sheetName("Users")
                    .rowHeight(30)
                    .autoFilter(true)
                    .write(Stream.of(Map.of("Name", "Alice", "Age", 30)))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals("Users", wb.getSheetAt(0).getSheetName());
                // auto filter present
                assertNotNull(wb.getSheetAt(0).getCTWorksheet().getAutoFilter());
            }
        }

        @Test
        @DisplayName("protectSheet on forMap result")
        void protectSheetWorks() throws IOException {
            var out = new ByteArrayOutputStream();
            ExcelWriter.forMap("Name")
                    .protectSheet("secret")
                    .write(Stream.of(Map.of("Name", "Alice")))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertNotNull(wb.getSheetAt(0).getCTWorksheet().getSheetProtection());
            }
        }
    }

    @Nested
    @DisplayName("CsvWriter.forMap returns a full CsvWriter")
    class CsvForMapReturnsFullWriter {

        @Test
        @DisplayName("all fluent methods available on forMap result (not just shortcut subset)")
        void fluentMethodsAvailable() {
            var out = new ByteArrayOutputStream();
            // The pre-v0.11.0 CsvMapWriter only exposed dialect/delimiter/charset/bom.
            // CsvWriter.forMap() must expose the full CsvWriter API.
            CsvWriter.forMap("Name", "Age")
                    .delimiter('|')
                    .bom(false)
                    .csvInjectionDefense(false)
                    .write(Stream.of(Map.of("Name", "Alice", "Age", 30)))
                    .write(out);

            String csv = out.toString(StandardCharsets.UTF_8);
            assertFalse(csv.startsWith("\uFEFF"), "bom(false) should suppress BOM");
            assertTrue(csv.contains("Name|Age"));
            assertTrue(csv.contains("Alice|30"));
        }
    }

    static class Person {
        String name;
        int age;
        String city;
    }

    @Nested
    @DisplayName("Reader column() returns Reader (not a builder)")
    class ReaderColumnReturnsReader {

        @Test
        @DisplayName("ExcelReader.column chain produces a single reader instance end-to-end")
        void excelReader_chainReturnsReader() throws IOException {
            // Write a test file
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Person>create()
                    .column("Name", p -> p.name)
                    .column("Age", p -> p.age)
                    .column("City", p -> p.city)
                    .write(Stream.of(makePerson("Alice", 30, "Seoul")))
                    .write(out);

            // The point of this test: `.column()` must return `ExcelReader<Person>`,
            // so we can use a single chained expression without intermediate variables.
            // If column() returned a Builder, this expression wouldn't compile because
            // Builder doesn't have build(InputStream).
            List<ReadResult<Person>> results = new ArrayList<>();
            new ExcelReader<>(Person::new, null)
                    .column((p, cell) -> p.name = cell.asString())
                    .column((p, cell) -> p.age = cell.asInt())
                    .column((p, cell) -> p.city = cell.asString())
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).data().name);
            assertEquals(30, results.get(0).data().age);
            assertEquals("Seoul", results.get(0).data().city);
        }

        @Test
        @DisplayName("ExcelReader: column(name) mixed with columnAt(idx) in one chain")
        void excelReader_mixedColumnKinds() throws IOException {
            // Header: Name, Age, City, Note
            var out = new ByteArrayOutputStream();
            String[] row = {"Alice", "30", "Seoul", "SKIP"};
            ExcelWriter.<String[]>create()
                    .column("Name", a -> a[0])
                    .column("Age", a -> a[1])
                    .column("City", a -> a[2])
                    .column("Note", a -> a[3])
                    .write(Stream.<String[]>of(row))
                    .write(out);

            // Read via: name (by header), skip, positional age, explicit index 3 (Note)
            Person result = new Person();
            new ExcelReader<>(() -> result, null)
                    .column("Name", (p, cell) -> p.name = cell.asString())
                    .column("Age", (p, cell) -> p.age = cell.asInt())
                    .columnAt(3, (p, cell) -> p.city = cell.asString())  // jump to Note, store as city
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(r -> {});

            assertEquals("Alice", result.name);
            assertEquals(30, result.age);
            assertEquals("SKIP", result.city);
        }

        @Test
        @DisplayName("CsvReader.column chain produces a single reader instance end-to-end")
        void csvReader_chainReturnsReader() throws IOException {
            String csv = "Name,Age,City\nAlice,30,Seoul\n";
            List<ReadResult<Person>> results = new ArrayList<>();
            new CsvReader<>(Person::new, null)
                    .column((p, cell) -> p.name = cell.asString())
                    .column((p, cell) -> p.age = cell.asInt())
                    .column((p, cell) -> p.city = cell.asString())
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(results::add);

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).data().name);
            assertEquals(30, results.get(0).data().age);
            assertEquals("Seoul", results.get(0).data().city);
        }

        @Test
        @DisplayName("Reader skipColumn works mid-chain after column(setter)")
        void skipColumn_midChain() throws IOException {
            String csv = "A,B,C\nkeep,skip,keep2\n";
            Person result = new Person();
            new CsvReader<>(() -> result, null)
                    .column((p, cell) -> p.name = cell.asString())
                    .skipColumn()
                    .column((p, cell) -> p.city = cell.asString())
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> {});

            assertEquals("keep", result.name);
            assertEquals("keep2", result.city);
        }
    }

    @Nested
    @DisplayName("ExcelReader.forMap(String...) column selection")
    class ExcelReaderForMapColumnSelection {

        @Test
        @DisplayName("forMap with selected columns returns only those columns")
        void forMap_selectedColumns_filtersOthers() throws IOException {
            // Write a 3-column Excel file
            var out = new ByteArrayOutputStream();
            ExcelWriter.forMap("Name", "Age", "City")
                    .write(Stream.of(
                            Map.of("Name", "Alice", "Age", 30, "City", "Seoul"),
                            Map.of("Name", "Bob", "Age", 25, "City", "Tokyo")))
                    .write(out);

            // Read with only "Name" and "City" selected
            List<ReadResult<Map<String, String>>> results = new ArrayList<>();
            ExcelReader.forMap("Name", "City")
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(results::add);

            assertEquals(2, results.size());
            Map<String, String> row1 = results.get(0).data();
            assertEquals("Alice", row1.get("Name"));
            assertEquals("Seoul", row1.get("City"));
            assertNull(row1.get("Age"), "Age should be filtered out");
            assertEquals(2, row1.size());
        }

        @Test
        @DisplayName("forMap with no args returns all columns")
        void forMap_noArgs_returnsAll() throws IOException {
            var out = new ByteArrayOutputStream();
            ExcelWriter.forMap("Name", "Age")
                    .write(Stream.of(Map.of("Name", "Alice", "Age", 30)))
                    .write(out);

            List<ReadResult<Map<String, String>>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertEquals(2, results.get(0).data().size());
        }

        @Test
        @DisplayName("CsvReader.forMap(String...) column selection")
        void csvReader_forMap_selectedColumns() {
            String csv = "Name,Age,City\nAlice,30,Seoul\nBob,25,Tokyo\n";
            List<ReadResult<Map<String, String>>> results = new ArrayList<>();
            CsvReader.forMap("Name", "City")
                    .build(new ByteArrayInputStream(csv.getBytes(java.nio.charset.StandardCharsets.UTF_8)))
                    .read(results::add);

            assertEquals(2, results.size());
            Map<String, String> row1 = results.get(0).data();
            assertEquals("Alice", row1.get("Name"));
            assertEquals("Seoul", row1.get("City"));
            assertNull(row1.get("Age"), "Age should be filtered out");
        }
    }

    private static Person makePerson(String name, int age, String city) {
        Person p = new Person();
        p.name = name;
        p.age = age;
        p.city = city;
        return p;
    }
}
