package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.*;

class NextApiTest {
    record Person(String name) {}

    @Test void headerPolicyMatchesNormalizedNames() {
        var values = new java.util.ArrayList<Person>();
        CsvReader.<Person>mapping(row -> new Person(row.get("user name").asString()))
                .headerPolicy(HeaderPolicy.NORMALIZED_CASE_INSENSITIVE)
                .read(csv(" User   Name \nA\n"), result -> values.add(result.data()));
        assertEquals(List.of(new Person("A")), values);
    }

    @Test void inputLimitRejectsOversizedInput() {
        assertThrows(ExcelKitException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(3, -1, -1, -1))
                .read(csv("A\n123\n"), result -> {}));
    }

    @Test void columnAndCellLimitsAbortCsvRead() {
        assertThrows(ExcelKitException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(-1, -1, 1, -1))
                .read(csv("A,B\n1,2\n"), result -> fail("must fail before rows")));
        assertThrows(io.github.dornol.excelkit.csv.CsvReadException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(-1, -1, -1, 2))
                .read(csv("A\nlong\n"), result -> fail("must not deliver an oversized cell")));
    }

    @Test void excelSheetLimitAbortsBeforeRows() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            workbook.createSheet("One").createRow(0).createCell(0).setCellValue("A");
            workbook.createSheet("Two").createRow(0).createCell(0).setCellValue("A");
            workbook.write(output);
        }
        assertThrows(io.github.dornol.excelkit.excel.ExcelReadException.class, () ->
                ExcelReader.forMap().limits(new ReadLimits(-1, 1, -1, -1))
                        .read(new ByteArrayInputStream(output.toByteArray()), result -> fail()));
    }

    @Test void cancellationStopsBetweenRows() {
        AtomicBoolean cancelled = new AtomicBoolean();
        var values = new java.util.ArrayList<Person>();
        ReadSummary summary = CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .cancellationToken(cancelled::get)
                .readWithSummary(csv("Name\nA\nB\n"), result -> {
                    values.add(result.data());
                    cancelled.set(true);
                });
        assertEquals(List.of(new Person("A")), values);
        assertTrue(summary.stoppedEarly());
    }

    @Test void detectorRecognizesCsvAndXlsxWithoutConsuming() {
        ByteArrayInputStream csv = csv("A,B\n1,2\n");
        assertEquals(TabularFileType.CSV, TabularFileDetector.detect(csv));
        assertEquals('A', csv.read());

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ExcelWriter.<Person>create().column("Name", Person::name)
                .write(List.of(new Person("A"))).writeTo(output);
        assertEquals(TabularFileType.XLSX,
                TabularFileDetector.detect(new ByteArrayInputStream(output.toByteArray())));
    }

    @Test void summaryAndBoundedReportCountRows() {
        ReadSummary summary = CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .readWithSummary(csv("Name\nA\nB\n"), result -> {});
        assertEquals(2, summary.totalRows());
        assertEquals(2, summary.successRows());
        assertEquals(0, summary.errorRows());
        assertFalse(summary.stoppedEarly());
        assertFalse(summary.duration().isNegative());

        ReadReport report = CsvReader.<Person>mapping(row -> { throw new IllegalArgumentException("bad"); })
                .readReport(csv("Name\nA\nB\n"), 1);
        assertEquals(2, report.summary().errorRows());
        assertEquals(1, report.errors().size());
        assertTrue(report.errorsTruncated());
    }

    @Test void detailedProgressReportsCounts() {
        AtomicReference<ReadProgress> progress = new AtomicReference<>();
        CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .onReadProgress(1, progress::set)
                .read(csv("Name\nA\n"), result -> {});
        assertEquals(1, progress.get().processedRows());
        assertEquals(1, progress.get().successRows());
    }

    @Test void writerSupportsStreamingOptionsAndStructuredTable() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ExcelWriter.<Person>create(options -> options.rowAccessWindowSize(1)
                        .compressTempFiles(true).useSharedStrings(true))
                .column("Name", Person::name)
                .table("PeopleTable")
                .write(List.of(new Person("A"), new Person("B"))).writeTo(output);
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("PeopleTable", workbook.getSheetAt(0).getTables().get(0).getName());
        }
    }

    @Test void structuredTableRejectsInvalidAndCellReferenceNames() {
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.create().table("bad name"));
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.create().table("A1"));
    }

    @Test void valueTypesValidateInvalidArguments() {
        assertThrows(IllegalArgumentException.class, () -> new ReadLimits(-2, -1, -1, -1));
        assertThrows(IllegalArgumentException.class, () ->
                new ReadSummary(-1, 0, 0, false, java.time.Duration.ZERO));
        assertThrows(IllegalArgumentException.class, () -> CsvReader.forMap().readReport(csv("A\n1\n"), -1));
        assertThrows(IllegalArgumentException.class, () ->
                TabularFileDetector.detect(java.io.InputStream.nullInputStream()));
    }

    private static ByteArrayInputStream csv(String value) {
        return new ByteArrayInputStream(value.getBytes(StandardCharsets.UTF_8));
    }
}
