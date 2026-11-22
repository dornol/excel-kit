package io.github.dornol.excelkit.schema;

import io.github.dornol.excelkit.core.read.CellConversionConfig;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** Every way a schema column can be declared: with read aliases, as required, with a write configuration. */
class SchemaBuilderOverloadsTest {

    public static class Person {
        String name;
        Integer age;
        Double price;
    }

    private static List<Person> readCsv(ExcelKitSchema<Person> schema, String csv) {
        List<Person> people = new ArrayList<>();
        schema.csvReader(Person::new).read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), r -> {
            if (r.success()) people.add(r.data());
        });
        return people;
    }

    private static List<ReadResult<Person>> readCsvResults(ExcelKitSchema<Person> schema, String csv) {
        List<ReadResult<Person>> results = new ArrayList<>();
        schema.csvReader(Person::new).read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), results::add);
        return results;
    }

    private static Person person(String name, Integer age) {
        Person p = new Person();
        p.name = name;
        p.age = age;
        return p;
    }

    // ---- read aliases ----

    @Test
    void aColumnWithAliasesIsWrittenUnderItsNameAndReadUnderAnyOfThem() throws Exception {
        ExcelKitSchema<Person> schema = ExcelKitSchema.<Person>builder()
                .column("Name", List.of("Full Name", "Employee"), p -> p.name, (p, c) -> p.name = c.asString())
                .build();

        // written under its name
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        schema.excelWriter().write(List.of(person("Ann", 1))).writeTo(out);
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("Name", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }

        // read under the name and under each alias
        assertEquals("Ann", readCsv(schema, "Name\nAnn\n").get(0).name);
        assertEquals("Bob", readCsv(schema, "Full Name\nBob\n").get(0).name);
        assertEquals("Cy", readCsv(schema, "Employee\nCy\n").get(0).name);
    }

    @Test
    void anAliasListWithANullIsRefused() {
        var builder = ExcelKitSchema.<Person>builder();

        IllegalArgumentException e = assertThrows(IllegalArgumentException.class, () ->
                builder.column("Name", Arrays.asList("Full Name", null), p -> p.name, (p, c) -> p.name = c.asString()).build());

        assertEquals("readHeaderNames cannot contain null", e.getMessage());
    }

    // ---- required columns, in each of their declarations ----

    private interface RequiredDeclaration {
        ExcelKitSchema<Person> schema();
    }

    private static List<RequiredDeclaration> requiredDeclarations() {
        return List.of(
                () -> ExcelKitSchema.<Person>builder()
                        .requiredColumn("Name", p -> p.name, (p, c) -> p.name = c.asString()).build(),
                () -> ExcelKitSchema.<Person>builder()
                        .requiredColumn("Name", p -> p.name, (p, c) -> p.name = c.asString(), c -> c.bold(true)).build(),
                () -> ExcelKitSchema.<Person>builder()
                        .requiredColumn("Name", List.of("Full Name"), p -> p.name, (p, c) -> p.name = c.asString()).build(),
                () -> ExcelKitSchema.<Person>builder()
                        .requiredColumn("Name", List.of("Full Name"), p -> p.name, (p, c) -> p.name = c.asString(), c -> c.bold(true)).build());
    }

    @Test
    void aRequiredColumnFailsTheRowsWhoseCellIsBlank() {
        for (int i = 0; i < requiredDeclarations().size(); i++) {
            ExcelKitSchema<Person> schema = requiredDeclarations().get(i).schema();

            List<ReadResult<Person>> results = readCsvResults(schema, "Name\nAnn\n\"\"\n");

            assertEquals(2, results.size(), "declaration " + i);
            assertEquals(true, results.get(0).success(), "declaration " + i + " non-blank");
            assertEquals("Ann", results.get(0).data().name);
            assertEquals(false, results.get(1).success(), "declaration " + i + " blank");
            assertEquals(io.github.dornol.excelkit.core.read.RowError.Type.VALIDATION, results.get(1).errorType());
            assertEquals(true, results.get(1).messages().stream().anyMatch(m -> m.contains("Required")),
                    "declaration " + i + ": " + results.get(1).messages());
        }
    }

    @Test
    void aRequiredColumnWithAliasesIsFoundUnderAnAlias() {
        for (int declaration : new int[]{2, 3}) {
            ExcelKitSchema<Person> schema = requiredDeclarations().get(declaration).schema();

            assertEquals("Ann", readCsv(schema, "Full Name\nAnn\n").get(0).name, "declaration " + declaration);
            assertEquals(false, readCsvResults(schema, "Full Name\n\"\"\n").get(0).success(), "declaration " + declaration);
        }
    }

    @Test
    void everyColumnNeedsItsHeaderInTheFileRequiredOrNot() {
        for (RequiredDeclaration declaration : requiredDeclarations()) {
            assertThrows(io.github.dornol.excelkit.csv.CsvReadException.class,
                    () -> readCsv(declaration.schema(), "Other\nAnn\n"));
        }
        ExcelKitSchema<Person> optional = ExcelKitSchema.<Person>builder()
                .column("Name", p -> p.name, (p, c) -> p.name = c.asString())
                .column("Age", List.of("Years"), p -> p.age, (p, c) -> p.age = c.asInt(), c -> c.type(ExcelDataType.INTEGER))
                .build();

        io.github.dornol.excelkit.csv.CsvReadException e = assertThrows(io.github.dornol.excelkit.csv.CsvReadException.class,
                () -> readCsv(optional, "Name\nAnn\n"));

        assertEquals(true, e.getMessage().contains("Header aliases [Age, Years] not found"), e.getMessage());
    }

    @Test
    void anOptionalColumnMayHaveABlankCellAndIsLeftEmpty() {
        ExcelKitSchema<Person> schema = ExcelKitSchema.<Person>builder()
                .column("Name", p -> p.name, (p, c) -> p.name = c.asString())
                .column("Age", List.of("Years"), p -> p.age, (p, c) -> p.age = c.asInt(0), c -> c.type(ExcelDataType.INTEGER))
                .build();

        List<ReadResult<Person>> results = readCsvResults(schema, "Name,Years\nAnn,\nBob,30\n");

        assertEquals(true, results.get(0).success(), String.valueOf(results.get(0).messages()));
        assertEquals(0, results.get(0).data().age);
        assertEquals(30, results.get(1).data().age);
    }

    // ---- write configuration ----

    @Test
    void theWriteConfigurationOfEveryDeclarationAppliesToTheWrittenColumn() throws Exception {
        ExcelKitSchema<Person> schema = ExcelKitSchema.<Person>builder()
                .column("A", List.of("a1"), p -> p.age, (p, c) -> p.age = c.asInt(), c -> c.type(ExcelDataType.INTEGER))
                .requiredColumn("B", p -> p.age, (p, c) -> p.age = c.asInt(), c -> c.type(ExcelDataType.INTEGER))
                .requiredColumn("C", List.of("c1"), p -> p.age, (p, c) -> p.age = c.asInt(), c -> c.type(ExcelDataType.INTEGER))
                .column("D", p -> p.age, (p, c) -> p.age = c.asInt())
                .build();

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        schema.excelWriter().write(List.of(person("x", 1234567))).writeTo(out);

        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var row = wb.getSheetAt(0).getRow(1);
            for (int column = 0; column < 3; column++) {
                assertEquals("#,##0", row.getCell(column).getCellStyle().getDataFormatString(), "column " + column);
            }
            // the column without a configuration keeps the default for its value
            assertEquals(1234567, (int) row.getCell(3).getNumericCellValue());
        }
    }

    // ---- reading with a conversion configuration ----

    @Test
    void theConversionConfigurationOfASchemaAppliesToItsReaders() {
        ExcelKitSchema<Person> german = ExcelKitSchema.<Person>builder()
                .column("Name", p -> p.name, (p, c) -> p.name = c.asString())
                .column("Price", p -> p.price, (p, c) -> p.price = c.asDouble())
                .cellConversion(CellConversionConfig.builder().locale(Locale.GERMANY).build())
                .build();

        List<Person> read = readCsv(german, "Name,Price\nAnn,\"2,5\"\n");

        assertEquals(2.5, read.get(0).price);
    }

    // ---- the error policy of the writer ----

    @Test
    void theWriteErrorPolicyOfASchemaIsTheOneItsWritersUse() throws Exception {
        java.util.function.Function<Person, Object> failsForBob = p -> {
            if ("Bob".equals(p.name)) throw new IllegalStateException("cannot read Bob");
            return p.name;
        };
        ExcelKitSchema<Person> failFast = ExcelKitSchema.<Person>builder()
                .column("Name", failsForBob, (p, c) -> p.name = c.asString()).build();
        ExcelKitSchema<Person> lenient = ExcelKitSchema.<Person>builder()
                .column("Name", failsForBob, (p, c) -> p.name = c.asString())
                .writeErrorPolicy(io.github.dornol.excelkit.excel.ExcelWriteErrorPolicy.LENIENT).build();
        List<Person> people = List.of(person("Ann", 1), person("Bob", 2), person("Cy", 3));

        assertThrows(io.github.dornol.excelkit.excel.ExcelWriteException.class,
                () -> failFast.excelWriter().write(people).writeTo(java.io.OutputStream.nullOutputStream()));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        lenient.excelWriter().write(people).writeTo(out);
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals("Ann", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(org.apache.poi.ss.usermodel.CellType.BLANK, sheet.getRow(2).getCell(0).getCellType(), "the failing row is blank");
            assertEquals("Cy", sheet.getRow(3).getCell(0).getStringCellValue());
        }
    }

    @Test
    void aNullWriteErrorPolicyIsRefused() {
        assertThrows(NullPointerException.class, () -> ExcelKitSchema.<Person>builder().writeErrorPolicy(null));
    }
}
