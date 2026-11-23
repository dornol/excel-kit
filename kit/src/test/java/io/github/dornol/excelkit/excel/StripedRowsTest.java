package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class StripedRowsTest {

    private static List<Integer> rows(int count) {
        return IntStream.range(0, count).boxed().toList();
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    /** The fill of every cell in column 0 below the header: its RGB as a string, or "-" without a solid fill. */
    private static List<String> fills(XSSFWorkbook wb, int sheet) {
        var result = new ArrayList<String>();
        var s = wb.getSheetAt(sheet);
        for (int r = 1; r <= s.getLastRowNum(); r++) {
            var cell = (XSSFCell) s.getRow(r).getCell(0);
            var style = cell.getCellStyle();
            if (style.getFillPattern() != FillPatternType.SOLID_FOREGROUND) {
                result.add("-");
                continue;
            }
            byte[] rgb = style.getFillForegroundColorColor().getRGB();
            result.add(String.format("%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
        }
        return result;
    }

    @Nested
    class Striping {
        @Test
        void shadesEverySecondDataRowAndNotTheFirst() throws IOException {
            try (var wb = open(ExcelWriter.<Integer>create().column("N", n -> n).stripedRows().write(rows(5)))) {
                assertEquals(List.of("-", "F2F2F2", "-", "F2F2F2", "-"), fills(wb, 0));
            }
        }

        @Test
        void takesAColor() throws IOException {
            try (var wb = open(ExcelWriter.<Integer>create().column("N", n -> n)
                    .stripedRows(ExcelColor.LIGHT_YELLOW).write(rows(3)))) {
                var fill = fills(wb, 0);
                assertEquals("-", fill.get(0));
                assertNotEquals("-", fill.get(1));
                assertNotEquals("F2F2F2", fill.get(1));
            }
        }

        @Test
        void theHeaderIsNotShaded() throws IOException {
            try (var wb = open(ExcelWriter.<Integer>create().headerColor(ExcelColor.STEEL_BLUE).column("N", n -> n)
                    .stripedRows().write(rows(2)))) {
                var header = (XSSFCell) wb.getSheetAt(0).getRow(0).getCell(0);
                byte[] rgb = header.getCellStyle().getFillForegroundColorColor().getRGB();
                assertNotEquals(0xF2, rgb[0] & 0xFF, "the header keeps its own color");
            }
        }

        @Test
        void aRolloverSheetStartsThePatternAgain() throws IOException {
            try (var wb = open(ExcelWriter.<Integer>create().maxRows(3).column("N", n -> n).stripedRows().write(rows(7)))) {
                assertEquals(3, wb.getNumberOfSheets());
                assertEquals(List.of("-", "F2F2F2", "-"), fills(wb, 0));
                assertEquals(List.of("-", "F2F2F2", "-"), fills(wb, 1));
                assertEquals(List.of("-"), fills(wb, 2));
            }
        }

        @Test
        void aWorkbookSheetStripesToo() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Integer>sheet("S").column("N", n -> n).stripedRows().write(rows(3).stream());
                try (var wb = open(workbook.finish())) {
                    assertEquals(List.of("-", "F2F2F2", "-"), fills(wb, 0));
                }
            }
        }

        @Test
        void noStripesByDefault() throws IOException {
            try (var wb = open(ExcelWriter.<Integer>create().column("N", n -> n).write(rows(3)))) {
                assertEquals(List.of("-", "-", "-"), fills(wb, 0));
            }
        }
    }

    @Nested
    class Priority {
        @Test
        void rowColorAndCellColorWinOverTheStripe() throws IOException {
            try (var wb = open(ExcelWriter.<Integer>create()
                    .rowColor(n -> n == 1 ? ExcelColor.LIGHT_RED : null)
                    .column("N", n -> n, c -> c.cellColor((v, n) -> n == 3 ? ExcelColor.LIGHT_GREEN : null))
                    .stripedRows().write(rows(5)))) {
                var fill = fills(wb, 0);
                assertEquals("-", fill.get(0));
                assertEquals(String.format("%02X%02X%02X", 255, 199, 206), fill.get(1), "rowColor");
                assertEquals("-", fill.get(2));
                assertNotEquals("F2F2F2", fill.get(3), "cellColor");
                assertNotEquals("-", fill.get(3));
                assertEquals("-", fill.get(4));
            }
        }

        @Test
        void nullColorIsRejected() {
            assertThrows(NullPointerException.class, () -> ExcelWriter.<Integer>create().stripedRows(null));
        }
    }
}
