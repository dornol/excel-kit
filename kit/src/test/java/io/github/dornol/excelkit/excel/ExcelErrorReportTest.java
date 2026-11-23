package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.CellError;
import io.github.dornol.excelkit.core.read.RowError;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ExcelErrorReportTest {

    static class Person {
        String name;
        int age;
    }

    private static RowError error(long fileRow, List<String> messages, List<CellError> cells, List<String> raw) {
        return new RowError(fileRow - 1, fileRow, RowError.Type.MAPPING, messages, null, cells, raw);
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    @Nested
    class Layout {
        @Test
        void holdsRowNumberUploadedValuesAndMessages() throws IOException {
            var errors = List.of(
                    error(3, List.of("age is not a number"),
                            List.of(new CellError(1, "Age", "abc", "not a number")), List.of("Bob", "abc")),
                    error(7, List.of(), List.of(new CellError(0, "Name", "", "required")), List.of("", "20")));

            try (var wb = open(ExcelErrorReport.create(List.of("Name", "Age")).write(errors))) {
                var sheet = wb.getSheet("Errors");
                assertNotNull(sheet);
                var header = sheet.getRow(0);
                assertEquals("Row", header.getCell(0).getStringCellValue());
                assertEquals("Name", header.getCell(1).getStringCellValue());
                assertEquals("Age", header.getCell(2).getStringCellValue());
                assertEquals("Error", header.getCell(3).getStringCellValue());

                var first = sheet.getRow(1);
                assertEquals(3, (long) first.getCell(0).getNumericCellValue());
                assertEquals("Bob", first.getCell(1).getStringCellValue());
                assertEquals("abc", first.getCell(2).getStringCellValue());
                assertEquals("age is not a number", first.getCell(3).getStringCellValue());

                var second = sheet.getRow(2);
                assertEquals(7, (long) second.getCell(0).getNumericCellValue());
                assertEquals("Name: required", second.getCell(3).getStringCellValue(),
                        "no row message: the cell messages, with their header");
            }
        }

        @Test
        void shadesTheFailedCellsAndCommentsThem() throws IOException {
            var errors = List.of(error(3, List.of("bad"),
                    List.of(new CellError(1, "Age", "abc", "not a number")), List.of("Bob", "abc")));
            try (var wb = open(ExcelErrorReport.create(List.of("Name", "Age")).write(errors))) {
                var row = wb.getSheet("Errors").getRow(1);
                XSSFCell failed = (XSSFCell) row.getCell(2);
                XSSFCell fine = (XSSFCell) row.getCell(1);
                assertEquals(ExcelColor.LIGHT_RED.getR(), failed.getCellStyle().getFillForegroundColorColor().getRGB()[0] & 0xFF);
                assertNotEquals(failed.getCellStyle().getFillForegroundColor(), fine.getCellStyle().getFillForegroundColor());
                assertNotNull(failed.getCellComment());
                assertEquals("not a number", failed.getCellComment().getString().getString());
                assertNull(fine.getCellComment());
            }
        }

        @Test
        void aShortOrMissingRawRowLeavesTheCellsEmpty() throws IOException {
            var errors = List.of(error(2, List.of("bad row"), List.of(), List.of("only")));
            try (var wb = open(ExcelErrorReport.create(List.of("A", "B", "C")).write(errors))) {
                var row = wb.getSheet("Errors").getRow(1);
                assertEquals("only", row.getCell(1).getStringCellValue());
                assertTrue(row.getCell(2) == null || row.getCell(2).getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK);
                assertEquals("bad row", row.getCell(4).getStringCellValue());
            }
        }

        @Test
        void aUnknownFileRowLeavesTheRowCellEmpty() throws IOException {
            var errors = List.of(error(-1, List.of("bad"), List.of(), List.of("x")));
            try (var wb = open(ExcelErrorReport.create(List.of("A")).write(errors))) {
                var cell = wb.getSheet("Errors").getRow(1).getCell(0);
                assertTrue(cell == null || cell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK);
            }
        }

        @Test
        void blankAndRepeatedHeadersAreRenamed() throws IOException {
            try (var wb = open(ExcelErrorReport.create(List.of("Name", "name", " ", "Error", "Row"))
                    .write(List.<RowError>of()))) {
                var header = wb.getSheet("Errors").getRow(0);
                var names = new ArrayList<String>();
                header.forEach(c -> names.add(c.getStringCellValue()));
                assertEquals(7, names.size());
                assertEquals(7, names.stream().map(String::toLowerCase).distinct().count(), names.toString());
            }
        }

        @Test
        void namesColorAndStreamAreConfigurable() throws IOException {
            var report = ExcelErrorReport.create(List.of("A")).rowColumnName("No.").errorColumnName("Reason")
                    .sheetName("Rejected").errorColor(ExcelColor.LIGHT_YELLOW);
            try (var wb = open(report.write(Stream.of(error(2, List.of("bad"), List.of(), List.of("x")))))) {
                var sheet = wb.getSheet("Rejected");
                assertEquals("No.", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("Reason", sheet.getRow(0).getCell(2).getStringCellValue());
            }
        }
    }

    @Nested
    class FromARealRead {
        @Test
        void turnsTheErrorsOfAnUploadIntoAReport() throws IOException {
            var upload = new ByteArrayOutputStream();
            ExcelWriter.<String[]>create()
                    .column("Name", r -> r[0])
                    .column("Age", r -> r[1])
                    .write(Stream.of(new String[]{"Alice", "30"}, new String[]{"Bob", "abc"}, new String[]{"Carol", "41"}))
                    .writeTo(upload);

            var ok = new ArrayList<Person>();
            var errors = new ArrayList<RowError>();
            ExcelReader.setter(Person::new)
                    .column((p, c) -> p.name = c.asString())
                    .column((p, c) -> p.age = c.asInt())
                    .read(new ByteArrayInputStream(upload.toByteArray()), ok::add, errors::add);

            assertEquals(2, ok.size());
            assertEquals(1, errors.size());

            try (var wb = open(ExcelErrorReport.create(List.of("Name", "Age")).write(errors))) {
                var row = wb.getSheet("Errors").getRow(1);
                assertEquals(3, (long) row.getCell(0).getNumericCellValue(), "Bob is row 3 of the uploaded file");
                assertEquals("Bob", row.getCell(1).getStringCellValue());
                assertEquals("abc", row.getCell(2).getStringCellValue());
                assertFalse(row.getCell(3).getStringCellValue().isBlank(), "the row says what went wrong");
                assertEquals(2, wb.getSheet("Errors").getPhysicalNumberOfRows());
            }
        }
    }

    @Nested
    class Rejections {
        @Test
        void invalidArgumentsAreRejected() {
            assertThrows(NullPointerException.class, () -> ExcelErrorReport.create(null));
            assertThrows(IllegalArgumentException.class, () -> ExcelErrorReport.create(List.of()));
            var report = ExcelErrorReport.create(List.of("A"));
            assertThrows(NullPointerException.class, () -> report.write((Iterable<RowError>) null));
            assertThrows(NullPointerException.class, () -> report.write((Stream<RowError>) null));
            assertThrows(NullPointerException.class, () -> report.write((io.github.dornol.excelkit.core.read.ReadReport) null));
            assertThrows(IllegalArgumentException.class, () -> report.sheetName(" "));
            assertThrows(IllegalArgumentException.class, () -> report.rowColumnName(""));
            assertThrows(IllegalArgumentException.class, () -> report.errorColumnName(" "));
            assertThrows(NullPointerException.class, () -> report.errorColor(null));
        }
    }

    @Nested
    class MoreCases {
        @Test
        void writesTheErrorsOfAReadReportInOrder() throws IOException {
            var errors = List.of(
                    error(3, List.of("first"), List.of(), List.of("a")),
                    error(9, List.of("second"), List.of(), List.of("b")));
            var report = new io.github.dornol.excelkit.core.read.ReadReport(
                    new io.github.dornol.excelkit.core.read.ReadSummary(10, 8, 2, false, java.time.Duration.ZERO), errors, true);
            try (var wb = open(ExcelErrorReport.create(List.of("A")).write(report))) {
                var sheet = wb.getSheet("Errors");
                assertEquals(3, sheet.getPhysicalNumberOfRows(), "header and the two errors the report kept");
                assertEquals(3, (long) sheet.getRow(1).getCell(0).getNumericCellValue());
                assertEquals("second", sheet.getRow(2).getCell(2).getStringCellValue());
            }
        }

        @Test
        void aNullHeaderIsNamedByItsPosition() throws IOException {
            try (var wb = open(ExcelErrorReport.create(java.util.Arrays.asList("Name", null)).write(List.<RowError>of()))) {
                var header = wb.getSheet("Errors").getRow(0);
                assertEquals("Column 2", header.getCell(2).getStringCellValue());
            }
        }

        @Test
        void aCellMessageWithoutAHeaderIsWrittenBare() throws IOException {
            var errors = List.of(error(2, List.of(), List.of(new CellError(0, null, "x", "bad value"),
                    new CellError(1, "Age", "y", "not a number")), List.of("x", "y")));
            try (var wb = open(ExcelErrorReport.create(List.of("A", "Age")).write(errors))) {
                assertEquals("bad value\nAge: not a number", wb.getSheet("Errors").getRow(1).getCell(3).getStringCellValue());
            }
        }

        @Test
        void aReportWithoutErrorsHoldsOnlyTheHeader() throws IOException {
            try (var wb = open(ExcelErrorReport.create(List.of("A")).write(List.<RowError>of()))) {
                assertEquals(1, wb.getSheet("Errors").getPhysicalNumberOfRows());
            }
        }
    }
}
