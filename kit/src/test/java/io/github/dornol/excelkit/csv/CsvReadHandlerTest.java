package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadResult;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.constraints.Max;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class CsvReadHandlerTest {

    private Validator validator;

    @BeforeEach
    void setUp() {
        validator = Validation.byDefaultProvider()
                .configure()
                .messageInterpolator(new ParameterMessageInterpolator())
                .buildValidatorFactory()
                .getValidator();
    }

    @Test
    void read_shouldReadCsvCorrectly() {
        String csv = "Name,Age\nAlice,30\nBob,25\nCharlie,35\n";
        List<TestPerson> results = new ArrayList<>();

        buildHandler(csv, validator).read(result -> {
            if (result.success()) {
                results.add(result.data());
            }
        });

        assertEquals(3, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(30, results.get(0).age);
        assertEquals("Bob", results.get(1).name);
        assertEquals("Charlie", results.get(2).name);
        assertEquals(35, results.get(2).age);
    }

    @Test
    void read_shouldValidateData() {
        String csv = "Name,Age\nValid,30\n,25\nTooOld,150\n";
        List<TestPerson> valid = new ArrayList<>();
        List<ReadResult<TestPerson>> invalid = new ArrayList<>();

        buildHandler(csv, validator).read(result -> {
            if (result.success()) {
                valid.add(result.data());
            } else {
                invalid.add(result);
            }
        });

        assertEquals(1, valid.size());
        assertEquals("Valid", valid.get(0).name);
        assertEquals(2, invalid.size());
        assertFalse(invalid.get(0).messages().isEmpty());
        assertFalse(invalid.get(1).messages().isEmpty());
    }

    @Test
    void read_shouldWorkWithoutValidator() {
        String csv = "Name,Age\nAlice,30\n";
        List<TestPerson> results = new ArrayList<>();

        buildHandler(csv, null).read(result -> {
            if (result.success()) {
                results.add(result.data());
            }
        });

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name);
    }

    @Test
    void read_shouldHandleQuotedFields() {
        String csv = "Name,Age\n\"Alice, Jr.\",30\n\"Bob \"\"The Builder\"\"\",25\n";
        List<TestPerson> results = new ArrayList<>();

        buildHandler(csv, null).read(result -> {
            if (result.success()) {
                results.add(result.data());
            }
        });

        assertEquals(2, results.size());
        assertEquals("Alice, Jr.", results.get(0).name);
        assertEquals("Bob \"The Builder\"", results.get(1).name);
    }

    @Test
    void read_shouldHandleEmptyCsv() {
        String csv = "Name,Age\n";
        List<ReadResult<TestPerson>> results = new ArrayList<>();

        buildHandler(csv, null).read(results::add);

        assertTrue(results.isEmpty());
    }

    @Test
    void read_shouldHandleMissingColumns() {
        String csv = "Name,Age\nAlice\n";
        List<TestPerson> results = new ArrayList<>();

        buildHandler(csv, null).read(result -> results.add(result.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(0, results.get(0).age);
    }

    @Test
    void read_viaReaderBuilderApi() {
        String csv = "Name,Age\nAlice,30\nBob,25\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestPerson> results = new ArrayList<>();

        new CsvReader<>(TestPerson::new, null)
                .column((p, cell) -> p.name = cell.asString())
                .column((p, cell) -> p.age = cell.asInt())
                .build(is)
                .read(result -> {
                    if (result.success()) {
                        results.add(result.data());
                    }
                });

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(25, results.get(1).age);
    }

    @Test
    void read_shouldThrowForBomOnlyHeader() {
        // The first header column contains only the BOM character
        String csv = "\uFEFF,Age\nAlice,30\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        CsvReadHandler<TestPerson> handler = new CsvReader<>(TestPerson::new, null)
                .column((p, cell) -> p.name = cell.asString())
                .column((p, cell) -> p.age = cell.asInt())
                .build(is);

        assertThrows(CsvReadException.class, () -> handler.read(r -> {}),
                "Should throw CsvReadException when first header is BOM-only");
    }

    @Test
    void constructor_shouldThrowForNullColumns() {
        InputStream is = new ByteArrayInputStream("a\n".getBytes());
        assertThrows(IllegalArgumentException.class,
                () -> new CsvReadHandler<>(is, null, TestPerson::new, null));
    }

    @Test
    void constructor_shouldThrowForEmptyColumns() {
        InputStream is = new ByteArrayInputStream("a\n".getBytes());
        assertThrows(IllegalArgumentException.class,
                () -> new CsvReadHandler<>(is, List.of(), TestPerson::new, null));
    }

    @Test
    void readAsStream_shouldReturnStreamOfResults() {
        String csv = "Name,Age\nAlice,30\nBob,25\nCharlie,35\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        Stream<ReadResult<TestPerson>> stream = new CsvReader<>(TestPerson::new, validator)
                .column((p, cell) -> p.name = cell.asString())
                .column((p, cell) -> p.age = cell.asInt())
                .build(is)
                .readAsStream();

        List<String> names = stream
                .filter(ReadResult::success)
                .map(r -> r.data().name)
                .toList();

        assertEquals(3, names.size());
        assertEquals("Alice", names.get(0));
        assertEquals("Bob", names.get(1));
        assertEquals("Charlie", names.get(2));
    }

    @Test
    void skipColumn_shouldSkipMiddleColumn() {
        String csv = "Col1,Col2,Col3\nA1,B1,C1\nA2,B2,C2\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestPersonThreeCol> results = new ArrayList<>();

        new CsvReader<>(TestPersonThreeCol::new, null)
                .column((p, cell) -> p.first = cell.asString())
                .skipColumn()
                .column((p, cell) -> p.third = cell.asString())
                .build(is)
                .read(result -> results.add(result.data()));

        assertEquals(2, results.size());
        assertEquals("A1", results.get(0).first);
        assertNull(results.get(0).second);
        assertEquals("C1", results.get(0).third);
        assertEquals("A2", results.get(1).first);
        assertEquals("C2", results.get(1).third);
    }

    @Test
    void skipColumns_shouldSkipMultipleColumns() {
        String csv = "Col1,Col2,Col3\nA1,B1,C1\nA2,B2,C2\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestPersonThreeCol> results = new ArrayList<>();

        new CsvReader<>(TestPersonThreeCol::new, null)
                .skipColumns(2)
                .column((p, cell) -> p.third = cell.asString())
                .build(is)
                .read(result -> results.add(result.data()));

        assertEquals(2, results.size());
        assertNull(results.get(0).first);
        assertEquals("C1", results.get(0).third);
        assertEquals("C2", results.get(1).third);
    }

    @Test
    void headerRowIndex_shouldSkipMetadataRows() {
        String csv = "Report Title,,\nGenerated: 2025-01-01,,\nName,Age\nAlice,30\nBob,25\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestPerson> results = new ArrayList<>();

        new CsvReader<>(TestPerson::new, null)
                .headerRowIndex(2)
                .column((p, cell) -> p.name = cell.asString())
                .column((p, cell) -> p.age = cell.asInt())
                .build(is)
                .read(result -> {
                    if (result.success()) {
                        results.add(result.data());
                    }
                });

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(30, results.get(0).age);
        assertEquals("Bob", results.get(1).name);
        assertEquals(25, results.get(1).age);
    }

    @Test
    void headerRowIndex_shouldThrowForInsufficientRows() {
        String csv = "only one row\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        CsvReadHandler<TestPerson> handler = new CsvReader<>(TestPerson::new, null)
                .headerRowIndex(5)
                .column((p, cell) -> p.name = cell.asString())
                .build(is);

        assertThrows(CsvReadException.class, () -> handler.read(r -> {}));
    }

    @Test
    void constructor_shouldThrowForNegativeHeaderRowIndex() {
        InputStream is = new ByteArrayInputStream("a\n".getBytes());
        assertThrows(IllegalArgumentException.class,
                () -> new CsvReadHandler<>(is, List.of(new CsvReadColumn<>((p, c) -> {})), TestPerson::new, null, -1));
    }

    private CsvReadHandler<TestPerson> buildHandler(String csv, Validator validator) {
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));
        return new CsvReader<>(TestPerson::new, validator)
                .column((p, cell) -> p.name = cell.asString())
                .column((p, cell) -> p.age = cell.asInt())
                .build(is);
    }

    public static class TestPerson {
        @NotBlank
        String name;

        @Min(1)
        @Max(100)
        int age;
    }

    public static class TestPersonThreeCol {
        String first;
        String second;
        String third;
    }
}
