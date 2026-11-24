package io.github.dornol.excelkit.excel;

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
 * {@code ExcelWriter} is a workbook with one sheet: it behaves as the same sheet written through an
 * {@code ExcelWorkbook} (names, rollover, settings), so the two do not differ in what they write.
 */
class OneSheetWorkbookTest {
    record Row(int n) {}

    private static List<Row> rows(int count) {
        return IntStream.range(0, count).mapToObj(Row::new).toList();
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static List<String> names(XSSFWorkbook wb) {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) names.add(wb.getSheetName(i));
        return names;
    }

    @Nested
    class SheetNames {
        @Test
        void writerAndWorkbookNameTheirRolledOverSheetsTheSame() throws IOException {
            List<String> viaWriter;
            try (var wb = open(ExcelWriter.<Row>create().column("N", Row::n).maxRows(2).write(rows(5)))) {
                viaWriter = names(wb);
            }
            ExcelHandler handler;
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("Sheet").column("N", Row::n).maxRows(2).write(rows(5));
                handler = workbook.finish();
            }
            try (var wb = open(handler)) {
                assertEquals(viaWriter, names(wb));
            }
            assertEquals(List.of("Sheet", "Sheet (2)", "Sheet (3)"), viaWriter);
        }

        @Test
        void aNameGivenToTheWriterNamesTheFirstSheetAndTheRolledOverOnes() throws IOException {
            try (var wb = open(ExcelWriter.<Row>create().sheetName("Data").column("N", Row::n).maxRows(2).write(rows(3)))) {
                assertEquals(List.of("Data", "Data (2)"), names(wb));
            }
        }

        @Test
        void aNameFunctionNamesTheFirstSheetToo() throws IOException {
            try (var wb = open(ExcelWriter.<Row>create().sheetName(i -> "Part " + (i + 1)).column("N", Row::n).maxRows(2).write(rows(3)))) {
                assertEquals(List.of("Part 1", "Part 2"), names(wb));
            }
        }
    }

    @Nested
    class SettingsOfTheWriterAndTheWorkbook {
        @Test
        void aHeaderColorSetAfterTheColumnsStillColorsTheHeader() throws IOException {
            try (var wb = open(ExcelWriter.<Row>create().column("N", Row::n)
                    .headerColor(io.github.dornol.excelkit.excel.style.ExcelColor.STEEL_BLUE).write(rows(1)))) {
                var color = (org.apache.poi.xssf.usermodel.XSSFColor) wb.getSheetAt(0).getRow(0).getCell(0)
                        .getCellStyle().getFillForegroundColorColor();
                assertArrayEquals(new byte[]{(byte) io.github.dornol.excelkit.excel.style.ExcelColor.STEEL_BLUE.getR(),
                        (byte) io.github.dornol.excelkit.excel.style.ExcelColor.STEEL_BLUE.getG(),
                        (byte) io.github.dornol.excelkit.excel.style.ExcelColor.STEEL_BLUE.getB()}, color.getRGB());
            }
        }

        @Test
        void aWriterWrittenTwiceFails() {
            var writer = ExcelWriter.<Row>create().column("N", Row::n);
            ExcelHandler first = writer.write(rows(1));
            first.discard();
            assertThrows(ExcelWriteException.class, () -> writer.write(rows(1)));
        }
    }

    @Nested
    class AfterAllAndRowCallbackOnASheetOfAWorkbook {
        @Test
        void afterAllIsWrittenOnTheLastSheetOnly() throws IOException {
            ExcelHandler handler;
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("A").column("N", Row::n).maxRows(2)
                        .afterAll(ctx -> {
                            ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("END");
                            return ctx.getCurrentRow() + 1;
                        })
                        .write(rows(3));
                handler = workbook.finish();
            }
            try (var wb = open(handler)) {
                assertEquals(2, wb.getNumberOfSheets());
                assertEquals(3, wb.getSheetAt(0).getLastRowNum() + 1, "header and two rows: no END on the first sheet");
                assertEquals("END", wb.getSheetAt(1).getRow(2).getCell(0).getStringCellValue());
            }
        }

        @Test
        void theRowCallbackSeesEveryRow() {
            List<Integer> seen = new ArrayList<>();
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("A").column("N", Row::n).write(rows(4).stream(), (row, cursor) -> seen.add(row.n()));
                workbook.finish().discard();
            }
            assertEquals(List.of(0, 1, 2, 3), seen);
        }
    }
}
