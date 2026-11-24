package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * What the one-sheet workbook ({@code ExcelWriter}) and the sheet of a workbook ({@code ExcelSheetWriter}) must do
 * the same way, and the edges of that path: failures, names that collide, settings that are taken at write time.
 */
class WriterFacadeGapsTest {
    record Row(int n) {}

    private static List<Row> rows(int count) {
        return IntStream.range(0, count).mapToObj(Row::new).toList();
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static byte[] rgb(XSSFWorkbook wb, int sheet) {
        return ((XSSFColor) wb.getSheetAt(sheet).getRow(0).getCell(0).getCellStyle().getFillForegroundColorColor()).getRGB();
    }

    @Nested
    class CollidingSheetNames {
        @Test
        void aNameFunctionThatRepeatsItselfFailsAtTheRolloverOfAWriter() {
            var writer = ExcelWriter.<Row>create().column("N", Row::n).maxRows(2).sheetName(i -> "Same");
            var e = assertThrows(ExcelWriteException.class, () -> writer.write(rows(5)));
            assertTrue(e.getMessage().contains("Duplicate sheet name: 'Same'"), e.getMessage());
        }

        @Test
        void aRolloverSheetCannotTakeTheNameOfAnotherSheetOfTheWorkbook() {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("A (2)").column("N", Row::n).write(rows(1));
                var e = assertThrows(ExcelWriteException.class,
                        () -> workbook.<Row>sheet("A").column("N", Row::n).maxRows(2).write(rows(5)));
                assertTrue(e.getMessage().contains("Duplicate sheet name: 'A (2)'"), e.getMessage());
            }
        }

        @Test
        void anInvalidNameFromTheFunctionIsRejected() {
            var writer = ExcelWriter.<Row>create().column("N", Row::n).sheetName(i -> "a/b");
            assertThrows(RuntimeException.class, () -> writer.write(rows(1)));
        }
    }

    @Nested
    class FailuresOnBothWriters {
        @Test
        void aColumnThatFailsIsAnExcelWriteExceptionOnASheetOfAWorkbookToo() {
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("A").column("N", r -> {
                    throw new IllegalStateException("boom");
                });
                var e = assertThrows(ExcelWriteException.class, () -> sheet.write(rows(1)));
                assertInstanceOf(IllegalStateException.class, e.getCause());
            }
        }

        @Test
        void aFailureBeforeTheFirstRowKeepsTheCallbacksOwnException() {
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("A").column("N", Row::n).beforeHeader(ctx -> {
                    throw new IllegalStateException("early");
                });
                assertThrows(IllegalStateException.class, () -> sheet.write(rows(1)));
            }
        }

        @Test
        void aFailingAfterAllReleasesTheWorkbookOfAWriter() {
            var writer = ExcelWriter.<Row>create().column("N", Row::n).afterAll(ctx -> {
                throw new IllegalStateException("late");
            });
            assertThrows(ExcelWriteException.class, () -> writer.write(rows(2)));
        }

        @Test
        void aWriterWithoutColumnsFailsAndCanThenBeUsed() throws IOException {
            var writer = ExcelWriter.<Row>create();
            assertThrows(RuntimeException.class, () -> writer.write(rows(1)));
            writer.column("N", Row::n);
            try (var wb = open(writer.write(rows(2)))) {
                assertEquals(3, wb.getSheetAt(0).getPhysicalNumberOfRows());
            }
        }

        @Test
        void aWriterWrittenTwiceSaysToCreateANewOne() {
            var writer = ExcelWriter.<Row>create().column("N", Row::n);
            writer.write(rows(1)).discard();
            var e = assertThrows(ExcelWriteException.class, () -> writer.write(rows(1)));
            assertTrue(e.getMessage().contains("create a new ExcelWriter"), e.getMessage());
        }

        @Test
        void aSheetWithoutColumnsFailsAndCanThenBeWritten() throws IOException {
            ExcelHandler handler;
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("A");
                assertThrows(RuntimeException.class, () -> sheet.write(rows(1)));
                sheet.column("N", Row::n);
                sheet.write(rows(2));
                handler = workbook.finish();
            }
            try (var wb = open(handler)) {
                assertEquals(3, wb.getSheetAt(0).getPhysicalNumberOfRows());
            }
        }

        @Test
        void aSheetWrittenTwiceFails() {
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("A").column("N", Row::n);
                sheet.write(rows(1));
                assertThrows(ExcelWriteException.class, () -> sheet.write(rows(1)));
                workbook.finish().discard();
            }
        }

        @Test
        void nullArgumentsAreRejectedByBoth() {
            var writer = ExcelWriter.<Row>create().column("N", Row::n);
            assertThrows(NullPointerException.class, () -> writer.write((java.util.stream.Stream<Row>) null));
            assertThrows(NullPointerException.class, () -> writer.write(rows(1).stream(), null));
            assertThrows(NullPointerException.class, () -> writer.write((Iterable<Row>) null));
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("A").column("N", Row::n);
                assertThrows(NullPointerException.class, () -> sheet.write((java.util.stream.Stream<Row>) null));
                assertThrows(NullPointerException.class, () -> sheet.write(rows(1).stream(), null));
                assertThrows(NullPointerException.class, () -> sheet.write((Iterable<Row>) null));
                assertThrows(NullPointerException.class, () -> sheet.afterAll(null));
            }
        }
    }

    @Nested
    class SettingsAreTakenWhenTheWriteStarts {
        @Test
        void aSheetOfAWorkbookIgnoresSettingsChangedWhileItWrites() throws IOException {
            ExcelHandler handler;
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("A").maxRows(10);
                sheet.column("N", r -> {
                    sheet.maxRows(1);
                    sheet.table("Late");
                    return r.n();
                });
                sheet.write(rows(3));
                handler = workbook.finish();
            }
            try (var wb = open(handler)) {
                assertEquals(1, wb.getNumberOfSheets());
                assertEquals(3, wb.getSheetAt(0).getLastRowNum());
                assertTrue(wb.getSheetAt(0).getTables().isEmpty());
            }
        }

        @Test
        void aColumnAddedWhileASheetWritesIsIgnored() throws IOException {
            ExcelHandler handler;
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("A");
                sheet.column("N", r -> {
                    if (r.n() == 2) sheet.column("Late", Row::n);
                    return r.n();
                });
                sheet.write(rows(4));
                handler = workbook.finish();
            }
            try (var wb = open(handler)) {
                assertEquals(1, wb.getSheetAt(0).getRow(0).getPhysicalNumberOfCells());
            }
        }

        @Test
        void aHeaderColorOfAWorkbookCountsFromTheSheetsMadeAfterIt() throws IOException {
            ExcelHandler handler;
            try (var workbook = ExcelWorkbook.create()) {
                var first = workbook.<Row>sheet("First").column("N", Row::n);
                workbook.headerColor(ExcelColor.STEEL_BLUE);
                var second = workbook.<Row>sheet("Second").column("N", Row::n);
                first.write(rows(1));
                second.write(rows(1));
                handler = workbook.finish();
            }
            try (var wb = open(handler)) {
                assertArrayEquals(new byte[]{(byte) 255, (byte) 255, (byte) 255}, rgb(wb, 0), "made before the color");
                assertArrayEquals(new byte[]{(byte) ExcelColor.STEEL_BLUE.getR(), (byte) ExcelColor.STEEL_BLUE.getG(),
                        (byte) ExcelColor.STEEL_BLUE.getB()}, rgb(wb, 1));
            }
        }
    }

    @Nested
    class CallbacksOnBothWriters {
        @Test
        void theCursorCountsTheRowsOnAWriterAndOnASheet() {
            List<Long> onWriter = new ArrayList<>();
            ExcelWriter.<Row>create().column("N", Row::n).write(rows(3).stream(), (r, c) -> onWriter.add((long) c.getCurrentTotal())).discard();
            List<Long> onSheet = new ArrayList<>();
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("A").column("N", Row::n).write(rows(3).stream(), (r, c) -> onSheet.add((long) c.getCurrentTotal()));
                workbook.finish().discard();
            }
            assertEquals(onWriter, onSheet);
            assertEquals(3, onSheet.size());
            assertEquals(onSheet.stream().sorted().toList(), onSheet, "the count only grows");
        }

        @Test
        void afterAllRunsAfterAfterDataOnTheLastSheetOfAWriter() throws IOException {
            var writer = ExcelWriter.<Row>create().column("N", Row::n).maxRows(2)
                    .afterData(ctx -> {
                        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("DATA-END");
                        return ctx.getCurrentRow() + 1;
                    })
                    .afterAll(ctx -> {
                        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("ALL-END");
                        return ctx.getCurrentRow() + 1;
                    });
            try (var wb = open(writer.write(rows(3)))) {
                assertEquals(2, wb.getNumberOfSheets());
                assertEquals("DATA-END", wb.getSheetAt(0).getRow(3).getCell(0).getStringCellValue());
                assertEquals("DATA-END", wb.getSheetAt(1).getRow(2).getCell(0).getStringCellValue());
                assertEquals("ALL-END", wb.getSheetAt(1).getRow(3).getCell(0).getStringCellValue());
                assertNull(wb.getSheetAt(0).getRow(4));
            }
        }

        @Test
        void aParallelStreamIsWrittenInItsOrderByBoth() throws IOException {
            var data = rows(50);
            try (var wb = open(ExcelWriter.<Row>create().column("N", Row::n, c -> c.type(ExcelDataType.INTEGER)).write(data.parallelStream()))) {
                for (int i = 0; i < 50; i++) {
                    assertEquals(i, (int) wb.getSheetAt(0).getRow(i + 1).getCell(0).getNumericCellValue());
                }
            }
        }
    }

    @Nested
    class SheetSettingValidation {
        @Test
        void invalidValuesAreRejectedOnBothWriters() {
            var writer = ExcelWriter.<Row>create().column("N", Row::n);
            assertThrows(IllegalArgumentException.class, () -> writer.rowHeight(0));
            assertThrows(IllegalArgumentException.class, () -> writer.groupComment("x"));
            assertThrows(IllegalArgumentException.class, () -> writer.namedRange(" ", 0));
            assertThrows(IllegalArgumentException.class, () -> writer.namedRange("Ok", -1));
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("A");
                assertThrows(IllegalArgumentException.class, () -> sheet.rowHeight(-1));
                assertThrows(IllegalArgumentException.class, () -> sheet.groupComment("x"));
                assertThrows(IllegalArgumentException.class, () -> sheet.namedRange("", 0));
                assertThrows(IllegalArgumentException.class, () -> sheet.namedRange("Ok", -1));
            }
        }
    }
}
