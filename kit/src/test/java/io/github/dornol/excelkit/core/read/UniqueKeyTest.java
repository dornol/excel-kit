package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelErrorReport;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class UniqueKeyTest {

    static class Emp {
        String dept;
        String id;
        String name;
    }

    private static byte[] excel(String[]... rows) throws IOException {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<String[]>create()
                .column("Dept", r -> r[0])
                .column("ID", r -> r[1])
                .column("Name", r -> r[2])
                .write(Stream.of(rows)).writeTo(out);
        return out.toByteArray();
    }

    private static ExcelReader<Emp> excelReader() {
        return ExcelReader.setter(Emp::new)
                .column("Dept", (e, c) -> e.dept = c.asString())
                .column("ID", (e, c) -> e.id = c.asString())
                .column("Name", (e, c) -> e.name = c.asString());
    }

    private static List<ReadResult<Emp>> read(ExcelReader<Emp> reader, byte[] file) {
        var results = new ArrayList<ReadResult<Emp>>();
        reader.read(new ByteArrayInputStream(file), results::add);
        return results;
    }

    @Nested
    class Excel {
        @Test
        void aRepeatedKeyFailsFromItsSecondRowOn() throws IOException {
            var results = read(excelReader().uniqueKey("ID"), excel(
                    new String[]{"A", "1001", "Ann"}, new String[]{"A", "1002", "Bob"},
                    new String[]{"B", "1001", "Cy"}, new String[]{"B", "1001", "Di"}));
            assertEquals(List.of(true, true, false, false), results.stream().map(r -> r.success()).toList());
            var dup = results.get(2);
            assertEquals(RowError.Type.VALIDATION, dup.errorType());
            assertEquals("Duplicate key [ID=1001]; first seen at row 2", dup.messages().get(0));
            assertEquals(4, dup.fileRowNum(), "header is row 1, Cy is the fourth row");
            assertEquals("first seen at row 2", dup.messages().get(0).substring(dup.messages().get(0).indexOf("first")));
            assertEquals("Cy", dup.data().name, "the mapped row is still delivered with the error");
            assertEquals(List.of("B", "1001", "Cy"), dup.rawValues());
            assertEquals(1, dup.cellErrors().size());
            assertEquals(1, dup.cellErrors().get(0).columnIndex());
            assertEquals("ID", dup.cellErrors().get(0).headerName());
            assertEquals("first seen at row 2", results.get(3).messages().get(0)
                    .substring(results.get(3).messages().get(0).indexOf("first")), "always the first row, not the previous one");
        }

        @Test
        void aKeyOfSeveralColumnsIsComparedAsAWhole() throws IOException {
            var results = read(excelReader().uniqueKey("Dept", "ID"), excel(
                    new String[]{"A", "1", "Ann"}, new String[]{"B", "1", "Bob"}, new String[]{"A", "1", "Cy"}));
            assertEquals(List.of(true, true, false), results.stream().map(r -> r.success()).toList());
            assertTrue(results.get(2).messages().get(0).startsWith("Duplicate key [Dept=A, ID=1]"));
            assertEquals(2, results.get(2).cellErrors().size());
        }

        @Test
        void valuesAreTrimmedAndCaseSensitiveAndBlankKeysAreNotChecked() throws IOException {
            var results = read(excelReader().uniqueKey("ID"), excel(
                    new String[]{"A", "x1", "1"}, new String[]{"A", " x1 ", "2"}, new String[]{"A", "X1", "3"},
                    new String[]{"A", "", "4"}, new String[]{"A", "", "5"}));
            assertEquals(List.of(true, false, true, true, true), results.stream().map(r -> r.success()).toList());
        }

        @Test
        void aNumericCellIsComparedByItsStoredNumber() throws IOException {
            var out = new ByteArrayOutputStream();
            try (var wb = new XSSFWorkbook()) {
                var sheet = wb.createSheet("S");
                sheet.createRow(0).createCell(0).setCellValue("ID");
                sheet.createRow(1).createCell(0).setCellValue(1001);
                sheet.createRow(2).createCell(0).setCellValue(1001);
                wb.write(out);
            }
            var results = new ArrayList<ReadResult<Emp>>();
            ExcelReader.setter(Emp::new).column("ID", (e, c) -> e.id = c.asString()).uniqueKey("ID")
                    .read(new ByteArrayInputStream(out.toByteArray()), results::add);
            assertEquals(List.of(true, false), results.stream().map(r -> r.success()).toList());
        }

        @Test
        void aRowThatAlreadyFailsReportsTheDuplicateTogetherWithItsOtherErrors() throws IOException {
            var reader = ExcelReader.setter(Emp::new)
                    .column("ID", (e, c) -> e.id = c.asString())
                    .column("Name", (e, c) -> { throw new IllegalStateException("bad name"); })
                    .uniqueKey("ID");
            var results = new ArrayList<ReadResult<Emp>>();
            reader.read(new ByteArrayInputStream(excel(
                    new String[]{"A", "1", "x"}, new String[]{"A", "1", "y"})), results::add);
            var second = results.get(1);
            assertFalse(second.success());
            assertEquals(RowError.Type.MAPPING, second.errorType(), "the kind of the first failure is kept");
            assertTrue(second.messages().stream().anyMatch(m -> m.startsWith("Duplicate key")));
            assertTrue(second.messages().size() >= 2);
        }

        @Test
        void worksInMappingAndMapMode() throws IOException {
            var file = excel(new String[]{"A", "1", "Ann"}, new String[]{"A", "1", "Bob"});
            var mapped = new ArrayList<Boolean>();
            ExcelReader.<String>mapping(row -> row.get("Name").asString()).uniqueKey("ID")
                    .read(new ByteArrayInputStream(file), r -> mapped.add(r.success()));
            assertEquals(List.of(true, false), mapped);

            var maps = new ArrayList<Boolean>();
            ExcelReader.forMap().uniqueKey("ID").read(new ByteArrayInputStream(file), r -> maps.add(r.success()));
            assertEquals(List.of(true, false), maps);
        }

        @Test
        void splitCallbacksGetTheDuplicateAsARowError() throws IOException {
            var ok = new ArrayList<Emp>();
            var errors = new ArrayList<RowError>();
            excelReader().uniqueKey("ID").read(new ByteArrayInputStream(excel(
                    new String[]{"A", "1", "Ann"}, new String[]{"A", "1", "Bob"})), ok::add, errors::add);
            assertEquals(1, ok.size());
            assertEquals(1, errors.size());
            assertEquals(3, errors.get(0).fileRowNum());
        }

        @Test
        void aKeyHeaderThatIsNotInTheFileFailsBeforeTheFirstRow() {
            var ex = assertThrows(RuntimeException.class, () -> read(excelReader().uniqueKey("Nope"),
                    excel(new String[]{"A", "1", "Ann"})));
            assertTrue(ex.getMessage().contains("Nope"), ex.getMessage());
        }

        @Test
        void countsTowardsMaxErrors() throws IOException {
            var reader = excelReader().uniqueKey("ID").maxErrors(0);
            assertThrows(ReadAbortException.class, () -> read(reader, excel(
                    new String[]{"A", "1", "Ann"}, new String[]{"A", "1", "Bob"})));
        }

        @Test
        void feedsTheErrorReport() throws IOException {
            var errors = new ArrayList<RowError>();
            excelReader().uniqueKey("ID").read(new ByteArrayInputStream(excel(
                    new String[]{"A", "1", "Ann"}, new String[]{"A", "1", "Bob"})), e -> {}, errors::add);
            var out = new ByteArrayOutputStream();
            ExcelErrorReport.create(List.of("Dept", "ID", "Name")).write(errors).writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var row = wb.getSheet("Errors").getRow(1);
                assertEquals("Bob", row.getCell(3).getStringCellValue());
                assertTrue(row.getCell(4).getStringCellValue().startsWith("Duplicate key [ID=1]"));
                assertNotNull(row.getCell(2).getCellComment(), "the key cell carries the message");
            }
        }
    }

    @Nested
    class Csv {
        private final String csv = "Dept,ID,Name\nA,1,Ann\nA,2,Bob\nB,1,Cy\n";

        private List<ReadResult<Emp>> read(CsvReader<Emp> reader, String text) {
            var results = new ArrayList<ReadResult<Emp>>();
            reader.read(new ByteArrayInputStream(text.getBytes(StandardCharsets.UTF_8)), results::add);
            return results;
        }

        private CsvReader<Emp> reader() {
            return CsvReader.setter(Emp::new)
                    .column("ID", (e, c) -> e.id = c.asString())
                    .column("Name", (e, c) -> e.name = c.asString());
        }

        @Test
        void aRepeatedKeyFails() {
            var results = read(reader().uniqueKey("ID"), csv);
            assertEquals(List.of(true, true, false), results.stream().map(r -> r.success()).toList());
            assertEquals("Duplicate key [ID=1]; first seen at row 2", results.get(2).messages().get(0));
        }

        @Test
        void worksInMappingMode() {
            var results = new ArrayList<Boolean>();
            CsvReader.<String>mapping(row -> row.get("Name").asString()).uniqueKey("Dept", "ID")
                    .read(new ByteArrayInputStream("Dept,ID,Name\nA,1,x\nA,1,y\nB,1,z\n".getBytes(StandardCharsets.UTF_8)),
                            r -> results.add(r.success()));
            assertEquals(List.of(true, false, true), results);
        }

        @Test
        void aMissingKeyHeaderFails() {
            assertThrows(ExcelKitException.class, () -> read(reader().uniqueKey("Nope"), csv));
        }
    }

    @Nested
    class Argument {
        @Test
        void rejectsBlankAndNullHeadersAndCanBeReplaced() {
            var reader = ExcelReader.setter(Emp::new);
            assertThrows(IllegalArgumentException.class, () -> reader.uniqueKey(" "));
            assertThrows(NullPointerException.class, () -> reader.uniqueKey((String[]) null));
            assertThrows(NullPointerException.class, () -> reader.uniqueKey("A", null));
            assertDoesNotThrow(() -> reader.uniqueKey("A").uniqueKey());
        }
    }

    @Nested
    class Reading {
        private byte[] twoSheets() throws IOException {
            try (var wb = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
                for (String name : new String[]{"First", "Second"}) {
                    var sheet = wb.createSheet(name);
                    sheet.createRow(0).createCell(0).setCellValue("ID");
                    sheet.createRow(1).createCell(0).setCellValue("1");
                    sheet.createRow(2).createCell(0).setCellValue(name.equals("Second") ? "1" : "2");
                }
                wb.write(out);
                return out.toByteArray();
            }
        }

        @Test
        void theHeadersAreMatchedLikeOtherHeadersThroughTheNormalizer() throws IOException {
            var out = new ByteArrayOutputStream();
            try (var wb = new XSSFWorkbook()) {
                var sheet = wb.createSheet("S");
                sheet.createRow(0).createCell(0).setCellValue("  Emp ID ");
                sheet.createRow(1).createCell(0).setCellValue("7");
                sheet.createRow(2).createCell(0).setCellValue("7");
                wb.write(out);
            }
            var results = new ArrayList<ReadResult<Emp>>();
            ExcelReader.setter(Emp::new).headerNormalizer(h -> h.trim().toLowerCase(java.util.Locale.ROOT))
                    .column("emp id", (e, c) -> e.id = c.asString())
                    .uniqueKey("EMP ID")
                    .read(new ByteArrayInputStream(out.toByteArray()), results::add);
            assertEquals(List.of(true, false), results.stream().map(r -> r.success()).toList());
        }

        @Test
        void onlyTheSelectedSheetIsRead() throws IOException {
            var results = new ArrayList<ReadResult<Emp>>();
            ExcelReader.setter(Emp::new).column("ID", (e, c) -> e.id = c.asString()).sheetIndex(1).uniqueKey("ID")
                    .read(new ByteArrayInputStream(twoSheets()), results::add);
            assertEquals(List.of(true, false), results.stream().map(r -> r.success()).toList());
        }

        @Test
        void eachSheetOfAWorkbookReaderKeepsItsOwnKeys() throws IOException {
            var first = new ArrayList<ReadResult<Emp>>();
            var second = new ArrayList<ReadResult<Emp>>();
            io.github.dornol.excelkit.excel.read.ExcelWorkbookReader.create()
                    .sheet("First", ExcelReader.setter(Emp::new).column("ID", (e, c) -> e.id = c.asString()).uniqueKey("ID"), first::add)
                    .sheet("Second", ExcelReader.setter(Emp::new).column("ID", (e, c) -> e.id = c.asString()).uniqueKey("ID"), second::add)
                    .read(new ByteArrayInputStream(twoSheets()));
            assertEquals(List.of(true, true), first.stream().map(r -> r.success()).toList(), "1 and 2");
            assertEquals(List.of(true, false), second.stream().map(r -> r.success()).toList(),
                    "1 and 1; the 1 of the first sheet does not count");
        }

        @Test
        void theFileRowNumbersFollowAHeaderThatIsNotOnTheFirstRow() throws IOException {
            var out = new ByteArrayOutputStream();
            try (var wb = new XSSFWorkbook()) {
                var sheet = wb.createSheet("S");
                sheet.createRow(0).createCell(0).setCellValue("Report title");
                sheet.createRow(2).createCell(0).setCellValue("ID");
                sheet.createRow(3).createCell(0).setCellValue("9");
                sheet.createRow(4).createCell(0).setCellValue("8");
                sheet.createRow(5).createCell(0).setCellValue("9");
                wb.write(out);
            }
            var results = new ArrayList<ReadResult<Emp>>();
            ExcelReader.setter(Emp::new).headerRowIndex(2).column("ID", (e, c) -> e.id = c.asString()).uniqueKey("ID")
                    .read(new ByteArrayInputStream(out.toByteArray()), results::add);
            assertEquals(3, results.size());
            assertEquals("Duplicate key [ID=9]; first seen at row 4", results.get(2).messages().get(0));
            assertEquals(6, results.get(2).fileRowNum());
        }

        @Test
        void blankRowsSkippedByTheReaderDoNotTakePartInTheKeys() throws IOException {
            var csv = "ID,Name\n1,a\n,\n,\n2,b\n";
            var results = new ArrayList<ReadResult<Emp>>();
            CsvReader.setter(Emp::new).skipBlankRows().column("ID", (e, c) -> e.id = c.asString()).uniqueKey("ID")
                    .read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), results::add);
            assertEquals(List.of(true, true), results.stream().map(r -> r.success()).toList());
        }

        @Test
        void theSplitCallbacksSeeNoDuplicateOnASuccessAndTheSummaryCountsIt() throws IOException {
            var summary = excelReader().uniqueKey("ID").readWithSummary(new ByteArrayInputStream(excel(
                    new String[]{"A", "1", "x"}, new String[]{"A", "1", "y"}, new String[]{"A", "2", "z"})), r -> { });
            assertEquals(3, summary.totalRows());
            assertEquals(2, summary.successRows());
            assertEquals(1, summary.errorRows());
        }
    }

    @Nested
    class Composite {
        @Test
        void keysThatOnlyLookAlikeWhenJoinedAreNotDuplicates() throws IOException {
            var results = read(excelReader().uniqueKey("Dept", "ID"), excel(
                    new String[]{"ab", "c", "1"}, new String[]{"a", "bc", "2"}, new String[]{"ab", "c", "3"}));
            assertEquals(List.of(true, true, false), results.stream().map(r -> r.success()).toList());
        }
    }
}
