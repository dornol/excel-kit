package io.github.dornol.excelkit.shared;

import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelHandler;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.constraints.Max;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ExcelKitSchema}
 */
class ExcelKitSchemaTest {

    private ExcelKitSchema<TestPerson> schema;
    private Validator validator;

    @BeforeEach
    void setUp() {
        schema = ExcelKitSchema.<TestPerson>builder()
                .column("Name", TestPerson::getName, (p, cell) -> p.setName(cell.asString()))
                .column("Age", TestPerson::getAge, (p, cell) -> p.setAge(cell.asInt()))
                .build();

        validator = Validation.byDefaultProvider()
                .configure()
                .messageInterpolator(new ParameterMessageInterpolator())
                .buildValidatorFactory()
                .getValidator();
    }

    @Test
    void builder_shouldBuildSchemaWithColumns() {
        assertNotNull(schema);
        assertEquals(2, schema.getColumns().size());
        assertEquals("Name", schema.getColumns().get(0).name());
        assertEquals("Age", schema.getColumns().get(1).name());
    }

    @Test
    void builder_shouldThrowWhenNoColumns() {
        assertThrows(IllegalArgumentException.class, () ->
                ExcelKitSchema.<TestPerson>builder().build());
    }

    @Test
    void excelWriter_shouldCreateWriterAndWrite() throws IOException {
        // Act
        ExcelHandler handler = schema.excelWriter()
                .write(Stream.of(new TestPerson("Alice", 30), new TestPerson("Bob", 25)));

        // Assert
        assertNotNull(handler);
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void excelWriter_shouldSupportAdditionalOptions() throws IOException {
        // Act
        ExcelHandler handler = schema.excelWriter()
                .title("Employee List")
                .autoFilter(true)
                .freezePane(1)
                .write(Stream.of(new TestPerson("Alice", 30)));

        // Assert
        assertNotNull(handler);
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void excelWriter_shouldSupportAdditionalColumns() throws IOException {
        // Act - schema columns + additional column via column() builder API
        ExcelHandler handler = schema.excelWriter()
                .column("Doubled Age", p -> p.getAge() * 2)
                .write(Stream.of(new TestPerson("Alice", 30)));

        // Assert
        assertNotNull(handler);
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void csvWriter_shouldCreateWriterAndWrite() {
        // Act
        CsvHandler handler = schema.csvWriter()
                .write(Stream.of(new TestPerson("Alice", 30), new TestPerson("Bob", 25)));

        // Assert
        assertNotNull(handler);
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        handler.consumeOutputStream(bos);
        String csv = bos.toString();
        String[] lines = csv.split("\\r?\\n");
        assertEquals(3, lines.length);
        assertEquals("\uFEFFName,Age", lines[0]);
        assertEquals("Alice,30", lines[1]);
        assertEquals("Bob,25", lines[2]);
    }

    @Test
    void csvWriter_shouldSupportDelimiter() {
        // Act
        CsvHandler handler = schema.csvWriter()
                .delimiter('\t')
                .write(Stream.of(new TestPerson("Alice", 30)));

        // Assert
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        handler.consumeOutputStream(bos);
        String csv = bos.toString();
        String[] lines = csv.split("\\r?\\n");
        assertEquals("\uFEFFName\tAge", lines[0]);
        assertEquals("Alice\t30", lines[1]);
    }

    @Test
    void csvWriter_shouldSupportAdditionalColumns() {
        // Act
        CsvHandler handler = schema.csvWriter()
                .column("Doubled Age", p -> p.getAge() * 2)
                .write(Stream.of(new TestPerson("Alice", 30)));

        // Assert
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        handler.consumeOutputStream(bos);
        String csv = bos.toString();
        String[] lines = csv.split("\\r?\\n");
        assertEquals("\uFEFFName,Age,Doubled Age", lines[0]);
        assertEquals("Alice,30,60", lines[1]);
    }

    @Test
    void excelReader_shouldReadExcelFile() throws IOException {
        // Arrange - write Excel first
        ByteArrayOutputStream excelOut = new ByteArrayOutputStream();
        schema.excelWriter()
                .write(Stream.of(
                        new TestPerson("Alice", 30),
                        new TestPerson("Bob", 25),
                        new TestPerson("Charlie", 35)
                ))
                .consumeOutputStream(excelOut);

        // Act - read back
        List<TestPerson> results = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(excelOut.toByteArray())) {
            schema.excelReader(TestPerson::new, null)
                    .build(is)
                    .read(result -> {
                        if (result.success()) {
                            results.add(result.data());
                        }
                    });
        }

        // Assert
        assertEquals(3, results.size());
        assertEquals("Alice", results.get(0).getName());
        assertEquals(30, results.get(0).getAge());
        assertEquals("Bob", results.get(1).getName());
        assertEquals(25, results.get(1).getAge());
        assertEquals("Charlie", results.get(2).getName());
        assertEquals(35, results.get(2).getAge());
    }

    @Test
    void excelReader_shouldSupportSheetIndex() throws IOException {
        // Arrange
        ByteArrayOutputStream excelOut = new ByteArrayOutputStream();
        schema.excelWriter()
                .write(Stream.of(new TestPerson("Alice", 30)))
                .consumeOutputStream(excelOut);

        // Act
        List<TestPerson> results = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(excelOut.toByteArray())) {
            schema.excelReader(TestPerson::new, null)
                    .sheetIndex(0)
                    .build(is)
                    .read(result -> {
                        if (result.success()) {
                            results.add(result.data());
                        }
                    });
        }

        // Assert
        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).getName());
    }

    @Test
    void csvReader_shouldReadCsvData() {
        // Arrange
        String csv = "Name,Age\nAlice,30\nBob,25\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        // Act
        List<TestPerson> results = new ArrayList<>();
        schema.csvReader(TestPerson::new, null)
                .build(is)
                .read(result -> {
                    if (result.success()) {
                        results.add(result.data());
                    }
                });

        // Assert
        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).getName());
        assertEquals(30, results.get(0).getAge());
        assertEquals("Bob", results.get(1).getName());
        assertEquals(25, results.get(1).getAge());
    }

    @Test
    void csvReader_shouldSupportDelimiter() {
        // Arrange
        String csv = "Name\tAge\nAlice\t30\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        // Act
        List<TestPerson> results = new ArrayList<>();
        schema.csvReader(TestPerson::new, null)
                .delimiter('\t')
                .build(is)
                .read(result -> {
                    if (result.success()) {
                        results.add(result.data());
                    }
                });

        // Assert
        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).getName());
        assertEquals(30, results.get(0).getAge());
    }

    @Test
    void roundTrip_excel_shouldPreserveData() throws IOException {
        // Arrange
        List<TestPerson> original = List.of(
                new TestPerson("Alice", 30),
                new TestPerson("Bob", 25),
                new TestPerson("Charlie", 35)
        );

        // Write
        ByteArrayOutputStream excelOut = new ByteArrayOutputStream();
        schema.excelWriter()
                .write(original.stream())
                .consumeOutputStream(excelOut);

        // Read back
        List<TestPerson> results = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(excelOut.toByteArray())) {
            schema.excelReader(TestPerson::new, null)
                    .build(is)
                    .read(result -> {
                        if (result.success()) {
                            results.add(result.data());
                        }
                    });
        }

        // Assert round-trip
        assertEquals(original.size(), results.size());
        for (int i = 0; i < original.size(); i++) {
            assertEquals(original.get(i).getName(), results.get(i).getName());
            assertEquals(original.get(i).getAge(), results.get(i).getAge());
        }
    }

    @Test
    void roundTrip_csv_shouldPreserveData() {
        // Arrange
        List<TestPerson> original = List.of(
                new TestPerson("Alice", 30),
                new TestPerson("Bob", 25),
                new TestPerson("Charlie", 35)
        );

        // Write
        ByteArrayOutputStream csvOut = new ByteArrayOutputStream();
        schema.csvWriter()
                .bom(false)
                .write(original.stream())
                .consumeOutputStream(csvOut);

        // Read back
        InputStream is = new ByteArrayInputStream(csvOut.toByteArray());
        List<TestPerson> results = new ArrayList<>();
        schema.csvReader(TestPerson::new, null)
                .build(is)
                .read(result -> {
                    if (result.success()) {
                        results.add(result.data());
                    }
                });

        // Assert round-trip
        assertEquals(original.size(), results.size());
        for (int i = 0; i < original.size(); i++) {
            assertEquals(original.get(i).getName(), results.get(i).getName());
            assertEquals(original.get(i).getAge(), results.get(i).getAge());
        }
    }

    @Test
    void excelReader_shouldSupportValidation() throws IOException {
        // Arrange - write Excel with invalid data
        ByteArrayOutputStream excelOut = new ByteArrayOutputStream();
        schema.excelWriter()
                .write(Stream.of(
                        new TestPerson("Valid", 30),
                        new TestPerson("", 25),         // blank name
                        new TestPerson("TooOld", 150)   // age > 100
                ))
                .consumeOutputStream(excelOut);

        // Act
        List<TestPerson> valid = new ArrayList<>();
        List<ReadResult<TestPerson>> invalid = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(excelOut.toByteArray())) {
            schema.excelReader(TestPerson::new, validator)
                    .build(is)
                    .read(result -> {
                        if (result.success()) {
                            valid.add(result.data());
                        } else {
                            invalid.add(result);
                        }
                    });
        }

        // Assert
        assertEquals(1, valid.size());
        assertEquals("Valid", valid.get(0).getName());
        assertEquals(2, invalid.size());
    }

    @Test
    void schemaIsImmutable_multipleWritersShouldBeIndependent() throws IOException {
        // Act - create two writers from same schema
        ExcelWriter<TestPerson> writer1 = schema.excelWriter().title("Writer 1");
        ExcelWriter<TestPerson> writer2 = schema.excelWriter().autoFilter(true);

        // Both should be able to write independently
        ExcelHandler handler1 = writer1.write(Stream.of(new TestPerson("Alice", 30)));
        ExcelHandler handler2 = writer2.write(Stream.of(new TestPerson("Bob", 25)));

        assertNotNull(handler1);
        assertNotNull(handler2);

        try (ByteArrayOutputStream bos1 = new ByteArrayOutputStream();
             ByteArrayOutputStream bos2 = new ByteArrayOutputStream()) {
            handler1.consumeOutputStream(bos1);
            handler2.consumeOutputStream(bos2);
            assertTrue(bos1.toByteArray().length > 0);
            assertTrue(bos2.toByteArray().length > 0);
        }
    }

    public static class TestPerson {
        @NotBlank
        private String name;

        @Min(1)
        @Max(100)
        private int age;

        public TestPerson() {}

        public TestPerson(String name, int age) {
            this.name = name;
            this.age = age;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public int getAge() {
            return age;
        }

        public void setAge(int age) {
            this.age = age;
        }
    }
}
