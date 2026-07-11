package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
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
        ReadLimitExceededException error = assertThrows(ReadLimitExceededException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(3, -1, -1, -1))
                .read(csv("A\n123\n"), result -> {}));
        assertEquals(ReadLimitExceededException.Limit.INPUT_BYTES, error.limit());
        assertEquals(3, error.configured());
        assertTrue(error.actual() > 3);
    }

    @Test void inputLimitStopsBeforeConsumingEntireStream() {
        byte[] content = new byte[1_000_000];
        java.util.concurrent.atomic.AtomicInteger consumed = new java.util.concurrent.atomic.AtomicInteger();
        var input = new ByteArrayInputStream(content) {
            @Override public synchronized int read(byte[] bytes, int off, int len) {
                int read = super.read(bytes, off, len);
                if (read > 0) consumed.addAndGet(read);
                return read;
            }
        };
        assertThrows(ReadLimitExceededException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(100, -1, -1, -1)).read(input, result -> {}));
        assertTrue(consumed.get() < content.length);
    }

    @Test void columnAndCellLimitsAbortCsvRead() {
        assertThrows(ExcelKitException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(-1, -1, 1, -1))
                .read(csv("A,B\n1,2\n"), result -> fail("must fail before rows")));
        assertThrows(ReadLimitExceededException.class, () -> CsvReader.forMap()
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
        assertThrows(ReadLimitExceededException.class, () ->
                ExcelReader.forMap().limits(new ReadLimits(-1, 1, -1, -1))
                        .read(new ByteArrayInputStream(output.toByteArray()), result -> fail()));
    }

    @Test void strictSecurityPolicyRejectsFormulas() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            var sheet = workbook.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue("Value");
            sheet.createRow(1).createCell(0).setCellFormula("1+1");
            workbook.write(output);
        }
        assertThrows(io.github.dornol.excelkit.excel.ExcelReadException.class, () ->
                ExcelReader.forMap().securityPolicy(ReadSecurityPolicy.STRICT)
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
        var progress = new java.util.ArrayList<ReadProgress>();
        CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .onReadProgress(10, progress::add)
                .read(csv("Name\nA\n"), result -> {});
        assertEquals(1, progress.size(), "completion event is required even below interval");
        assertTrue(progress.get(0).completed());
        assertEquals(1, progress.get(0).successRows());
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

    @Test void summaryReportAndWhileSupportPathAndSource() throws Exception {
        var path = Files.createTempFile("excel-kit-summary", ".csv");
        try {
            Files.writeString(path, "Name\nA\nB\n");
            var reader = CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()));
            assertEquals(2, reader.readWithSummary(path, result -> {}).totalRows());
            assertEquals(0, reader.readReport((InputStreamSource) () -> Files.newInputStream(path), 1)
                    .summary().errorRows());
            ReadSummary stopped = reader.readWhile(path, result -> false);
            assertEquals(1, stopped.totalRows());
            assertTrue(stopped.stoppedEarly());
        } finally { Files.deleteIfExists(path); }
    }

    @Test void detailedDetectionReportsCharsetDelimiterAndReadSupport() throws Exception {
        byte[] utf16 = "A;B\n1;2\n".getBytes(StandardCharsets.UTF_16LE);
        byte[] bom = new byte[utf16.length + 2];
        bom[0] = (byte) 0xff; bom[1] = (byte) 0xfe;
        System.arraycopy(utf16, 0, bom, 2, utf16.length);
        var result = TabularFileDetector.detectDetailed(new ByteArrayInputStream(bom));
        assertEquals(StandardCharsets.UTF_16LE, result.charset());
        assertEquals(';', result.delimiter());
        assertTrue(TabularFileType.XLSX.isReadable());
        assertFalse(TabularFileType.XLS.isReadable());
    }

    @Test void rolloverCreatesUniqueStructuredTables() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ExcelWriter.<Person>create().column("Name", Person::name).maxRows(2)
                .table(new io.github.dornol.excelkit.excel.TableOptions(
                        "People", "TableStyleMedium3", true, true))
                .write(List.of(new Person("A"), new Person("B"), new Person("C"))).writeTo(output);
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals(2, workbook.getNumberOfSheets());
            assertEquals("People_1", workbook.getSheetAt(0).getTables().get(0).getName());
            assertEquals("People_2", workbook.getSheetAt(1).getTables().get(0).getName());
        }
    }

    @Test void workbookSheetSupportsSharedStreamingOptionsAndTables() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var workbook = io.github.dornol.excelkit.excel.ExcelWorkbook.create(options ->
                options.streaming(new io.github.dornol.excelkit.excel.StreamingOptions(1, true, true)))) {
            workbook.<Person>sheet("People").column("Name", Person::name)
                    .table("WorkbookPeople").write(List.of(new Person("A"), new Person("B")));
            workbook.finish().writeTo(output);
        }
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("WorkbookPeople", workbook.getSheet("People").getTables().get(0).getName());
        }
    }

    @Test void templateListSupportsStructuredTable() throws Exception {
        ByteArrayOutputStream template = new ByteArrayOutputStream();
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("Name");
            workbook.write(template);
        }
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var writer = new io.github.dornol.excelkit.excel.ExcelTemplateWriter(
                new ByteArrayInputStream(template.toByteArray()))) {
            writer.<Person>list(1).column("Name", Person::name).table("TemplatePeople")
                    .write(List.of(new Person("A")));
            writer.finish().writeTo(output);
        }
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("TemplatePeople", workbook.getSheet("Data").getTables().get(0).getName());
        }
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
