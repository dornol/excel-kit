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
}
