package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class SheetVisibilityZoomTest {

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static ExcelSheetWriter<String> sheet(ExcelWorkbook workbook, String name) {
        return workbook.<String>sheet(name).column("V", s -> s);
    }

    @Nested
    class HideSheet {
        @Test
        void hidesAndVeryHidesSheetsAndKeepsTheRestVisible() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                sheet(workbook, "Main").write(Stream.of("a"));
                sheet(workbook, "Codes").hideSheet().write(Stream.of("x"));
                sheet(workbook, "Secret").veryHideSheet().write(Stream.of("y"));
                try (var wb = open(workbook.finish())) {
                    assertEquals(SheetVisibility.VISIBLE, wb.getSheetVisibility(0));
                    assertEquals(SheetVisibility.HIDDEN, wb.getSheetVisibility(1));
                    assertEquals(SheetVisibility.VERY_HIDDEN, wb.getSheetVisibility(2));
                    assertEquals("x", wb.getSheet("Codes").getRow(1).getCell(0).getStringCellValue(),
                            "a hidden sheet still holds its data");
                }
            }
        }

        @Test
        void aHiddenFirstSheetIsNotWhatTheWorkbookOpensOn() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                sheet(workbook, "Codes").hideSheet().write(Stream.of("x"));
                sheet(workbook, "Main").write(Stream.of("a"));
                try (var wb = open(workbook.finish())) {
                    assertEquals(1, wb.getActiveSheetIndex());
                    assertFalse(wb.getSheetAt(0).isSelected(), "a hidden sheet is not selected");
                    assertTrue(wb.getSheetAt(1).isSelected());
                }
            }
        }

        @Test
        void aWorkbookWithoutAVisibleSheetIsRejectedAndCanStillBeClosed() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                sheet(workbook, "Codes").hideSheet().write(Stream.of("x"));
                var e = assertThrows(ExcelWriteException.class, workbook::finish);
                assertTrue(e.getMessage().contains("visible"));
            }
        }

        @Test
        void everySheetOfARolloverIsHidden() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                sheet(workbook, "Main").write(Stream.of("a"));
                workbook.<String>sheet("Codes").column("V", s -> s).maxRows(2).hideSheet()
                        .write(Stream.of("1", "2", "3", "4", "5"));
                try (var wb = open(workbook.finish())) {
                    assertTrue(wb.getNumberOfSheets() > 2, "the codes rolled over");
                    for (int i = 1; i < wb.getNumberOfSheets(); i++) {
                        assertEquals(SheetVisibility.HIDDEN, wb.getSheetVisibility(i), wb.getSheetName(i));
                    }
                }
            }
        }
    }

    @Nested
    class Zoom {
        @Test
        void setsTheZoomOfASingleWriter() throws IOException {
            try (var wb = open(ExcelWriter.<String>create().column("V", s -> s).zoom(80).write(Stream.of("a")))) {
                assertEquals(80, wb.getSheetAt(0).getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale());
            }
        }

        @Test
        void setsTheZoomOfEverySheetOfARollover() throws IOException {
            try (var wb = open(ExcelWriter.<String>create().column("V", s -> s).maxRows(2).zoom(120)
                    .write(Stream.of("1", "2", "3", "4", "5")))) {
                assertTrue(wb.getNumberOfSheets() > 1);
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    assertEquals(120, wb.getSheetAt(i).getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale());
                }
            }
        }

        @Test
        void aWorkbookSheetTakesItToo() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                sheet(workbook, "A").zoom(50).write(Stream.of("a"));
                try (var wb = open(workbook.finish())) {
                    assertEquals(50, wb.getSheetAt(0).getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale());
                }
            }
        }

        @Test
        void onlyTheRangeExcelAcceptsIsAllowed() {
            Consumer<Integer> zoom = p -> ExcelWriter.<String>create().zoom(p);
            assertThrows(IllegalArgumentException.class, () -> zoom.accept(9));
            assertThrows(IllegalArgumentException.class, () -> zoom.accept(401));
            assertDoesNotThrow(() -> zoom.accept(10));
            assertDoesNotThrow(() -> zoom.accept(400));
        }
    }

    @Nested
    class MoreCases {
        @Test
        void aWorkbookWithoutSheetsIsNotRejectedForHavingNoVisibleSheet() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                assertDoesNotThrow(workbook::finish);
            }
        }

        @Test
        void aHiddenSheetAfterTheVisibleOneKeepsTheFirstAsTheActiveSheet() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                sheet(workbook, "Main").write(Stream.of("a"));
                sheet(workbook, "Codes").hideSheet().write(Stream.of("x"));
                try (var wb = open(workbook.finish())) {
                    assertEquals(0, wb.getActiveSheetIndex());
                    assertTrue(wb.getSheetAt(0).isSelected());
                    assertFalse(wb.getSheetAt(1).isSelected());
                }
            }
        }
    }
}
