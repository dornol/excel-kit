package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.ReadAbortException;
import io.github.dornol.excelkit.shared.ReadResult;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.constraints.Max;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for mapping-based (immutable object) CSV reading.
 */
class CsvMappingReadTest {

    @TempDir
    Path tempDir;

    record PersonRecord(String name, Integer age, String city) {}

    // --- Basic functionality ---

    @Test
    void mapping_shouldCreateImmutableObjects() {
        String csv = "Name,Age,City\nAlice,30,Seoul\nBob,25,Busan";

        List<PersonRecord> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                row.get("City").asString()
        )).build(toInputStream(csv))
          .read(r -> {
              assertTrue(r.success());
              results.add(r.data());
          });

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
        assertEquals("Bob", results.get(1).name());
        assertEquals(25, results.get(1).age());
        assertEquals("Busan", results.get(1).city());
    }

    @Test
    void mapping_shouldWorkWithDifferentColumnOrder() {
        String csv = "City,Age,Name\nSeoul,30,Alice\nBusan,25,Bob";

        List<PersonRecord> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                row.get("City").asString()
        )).build(toInputStream(csv))
          .read(r -> results.add(r.data()));

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals("Seoul", results.get(0).city());
    }

    @Test
    void mapping_shouldReadSubsetOfColumns() {
        String csv = "Name,Age,City,Email\nAlice,30,Seoul,a@test.com";

        List<PersonRecord> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                null,
                row.get("City").asString()
        )).build(toInputStream(csv))
          .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
        assertNull(results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
    }

    // --- Error handling ---

    @Test
    void mapping_shouldThrowOnMissingHeader() {
        String csv = "Name,Age\nAlice,30";

        List<ReadResult<PersonRecord>> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                row.get("NonExistent").asString()
        )).build(toInputStream(csv))
          .read(results::add);

        assertEquals(1, results.size());
        assertFalse(results.get(0).success());
        assertNull(results.get(0).data());
        assertNotNull(results.get(0).messages());
        assertTrue(results.get(0).messages().get(0).contains("NonExistent"));
    }

    @Test
    void mapping_shouldHandleConversionError() {
        String csv = "Name,Age\nAlice,not-a-number";

        List<ReadResult<PersonRecord>> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                null
        )).build(toInputStream(csv))
          .read(results::add);

        assertEquals(1, results.size());
        assertFalse(results.get(0).success());
    }

    @Test
    void mapping_shouldContinueAfterRowError() {
        String csv = "Name,Age\nAlice,30\nBob,bad\nCharlie,35";

        List<ReadResult<PersonRecord>> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                null
        )).build(toInputStream(csv))
          .read(results::add);

        assertEquals(3, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data().name());
        assertFalse(results.get(1).success());
        assertTrue(results.get(2).success());
        assertEquals("Charlie", results.get(2).data().name());
    }

    // --- Read modes ---

    @Test
    void mapping_shouldWorkWithReadStrict() {
        String csv = "Name,Age,City\nAlice,30,Seoul";

        List<PersonRecord> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                row.get("City").asString()
        )).build(toInputStream(csv))
          .readStrict(results::add);

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
    }

    @Test
    void mapping_readStrict_shouldThrowOnError() {
        String csv = "Name,Age\nAlice,30\nBob,bad";

        var handler = CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                null
        )).build(toInputStream(csv));

        assertThrows(ReadAbortException.class, () -> handler.readStrict(r -> {}));
    }

    @Test
    void mapping_shouldWorkWithReadAsStream() {
        String csv = "Name,Age,City\nAlice,30,Seoul\nBob,25,Busan";

        List<PersonRecord> results;
        try (var stream = CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                row.get("City").asString()
        )).build(toInputStream(csv)).readAsStream()) {
            results = stream
                    .filter(ReadResult::success)
                    .map(ReadResult::data)
                    .toList();
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals("Bob", results.get(1).name());
    }

    // --- Configuration ---

    @Test
    void mapping_shouldWorkWithCustomDelimiter() {
        String tsv = "Name\tAge\tCity\nAlice\t30\tSeoul";

        List<PersonRecord> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                row.get("City").asString()
        )).delimiter('\t')
          .build(toInputStream(tsv))
          .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
    }

    @Test
    void mapping_shouldWorkWithHeaderRowIndex() {
        String csv = "METADATA\nSKIP\nName,Age\nAlice,30";

        List<PersonRecord> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                null
        )).headerRowIndex(2)
          .build(toInputStream(csv))
          .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
    }

    @Test
    void mapping_shouldWorkWithProgressCallback() {
        String csv = "Name,Age\nA,1\nB,2\nC,3\nD,4\nE,5";

        AtomicLong lastProgress = new AtomicLong(0);
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                null
        )).onProgress(2, (count, cursor) -> lastProgress.set(count))
          .build(toInputStream(csv))
          .read(r -> {});

        assertEquals(4, lastProgress.get());
    }

    // --- RowData access ---

    @Test
    void mapping_shouldSupportIndexAccess() {
        String csv = "Name,Age,City\nAlice,30,Seoul";

        List<PersonRecord> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get(0).asString(),
                row.get(1).asInt(),
                row.get(2).asString()
        )).build(toInputStream(csv))
          .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
    }

    @Test
    void mapping_rowDataHasShouldWork() {
        String csv = "Name,Age\nAlice,30";

        CsvReader.<PersonRecord>mapping(row -> {
            assertTrue(row.has("Name"));
            assertTrue(row.has("Age"));
            assertFalse(row.has("City"));
            assertEquals(List.of("Name", "Age"), row.headerNames());
            return new PersonRecord(row.get("Name").asString(), row.get("Age").asInt(), null);
        }).build(toInputStream(csv))
          .read(r -> {});
    }

    // --- Bean Validation ---

    @Test
    void mapping_shouldWorkWithBeanValidation() {
        String csv = "Name,Age\nAlice,30\n,25\nCharlie,150";

        Validator validator = Validation.byDefaultProvider()
                .configure()
                .messageInterpolator(new ParameterMessageInterpolator())
                .buildValidatorFactory()
                .getValidator();

        List<ReadResult<ValidatedPerson>> results = new ArrayList<>();
        CsvReader.<ValidatedPerson>mapping(row -> {
            ValidatedPerson p = new ValidatedPerson();
            p.name = row.get("Name").asString();
            p.age = row.get("Age").asInt();
            return p;
        }, validator).build(toInputStream(csv))
          .read(results::add);

        assertEquals(3, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data().name);
        assertFalse(results.get(1).success());  // blank name
        assertFalse(results.get(2).success());  // age > 100
    }

    // --- Round-trip ---

    @Test
    void roundTrip_writeWithCsvWriter_readWithMapping() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<PersonRecord>()
                .column("Name", PersonRecord::name)
                .column("Age", p -> p.age())
                .column("City", PersonRecord::city)
                .bom(false)
                .write(Stream.of(
                        new PersonRecord("Alice", 30, "Seoul"),
                        new PersonRecord("Bob", 25, "Busan")))
                .consumeOutputStream(out);

        List<PersonRecord> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                row.get("City").asString()
        )).build(new ByteArrayInputStream(out.toByteArray()))
          .read(r -> {
              assertTrue(r.success());
              results.add(r.data());
          });

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
    }

    // --- Edge cases ---

    @Test
    void mapping_shouldHandleEmptyFile() {
        String csv = "Name,Age";  // only header, no data

        List<ReadResult<PersonRecord>> results = new ArrayList<>();
        CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
                row.get("Name").asString(), null, null
        )).build(toInputStream(csv))
          .read(results::add);

        assertTrue(results.isEmpty());
    }

    // --- Helper ---

    public static class ValidatedPerson {
        @NotBlank
        String name;
        @Min(1) @Max(100)
        int age;
    }

    private InputStream toInputStream(String content) {
        return new ByteArrayInputStream(content.getBytes(StandardCharsets.UTF_8));
    }
}
