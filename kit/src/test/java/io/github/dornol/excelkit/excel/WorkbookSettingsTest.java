package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/** {@link ExcelWriter} and {@link ExcelWorkbook} share their workbook settings, so they accept and refuse the same. */
class WorkbookSettingsTest {

    private static void both(Consumer<ExcelWriter<String>> onWriter, Consumer<ExcelWorkbook> onWorkbook) {
        onWriter.accept(ExcelWriter.create());
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            onWorkbook.accept(workbook);
        }
    }

    @Test
    void bothRefuseTheSameInvalidSettings() {
        both(w -> assertThrows(IllegalArgumentException.class, () -> w.headerColor(null)),
                w -> assertThrows(IllegalArgumentException.class, () -> w.headerColor(null)));
        both(w -> assertThrows(IllegalArgumentException.class, () -> w.headerFontSize(0)),
                w -> assertThrows(IllegalArgumentException.class, () -> w.headerFontSize(0)));
        both(w -> assertThrows(NullPointerException.class, () -> w.headerFontName(null)),
                w -> assertThrows(NullPointerException.class, () -> w.headerFontName(null)));
        both(w -> assertThrows(IllegalArgumentException.class, () -> w.password("  ")),
                w -> assertThrows(IllegalArgumentException.class, () -> w.password("  ")));
        both(w -> assertThrows(IllegalArgumentException.class, () -> w.protectWorkbook("")),
                w -> assertThrows(IllegalArgumentException.class, () -> w.protectWorkbook("")));
    }

    @Test
    void theLastHeaderSettingsWinWhateverTheOrder() throws Exception {
        var writerOut = new ByteArrayOutputStream();
        ExcelWriter.<String>create().headerFontName("Arial").headerColor(ExcelColor.STEEL_BLUE).headerFontSize(14)
                .headerFontName("Courier New").column("v", s -> s).write(Stream.of("a")).writeTo(writerOut);
        var workbookOut = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create().headerFontName("Arial").headerColor(ExcelColor.STEEL_BLUE)
                .headerFontSize(14).headerFontName("Courier New")) {
            workbook.<String>sheet("S").column("v", s -> s).write(Stream.of("a"));
            workbook.finish().writeTo(workbookOut);
        }

        for (byte[] bytes : new byte[][]{writerOut.toByteArray(), workbookOut.toByteArray()}) {
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                var font = wb.getFontAt(wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getFontIndex());
                assertEquals("Courier New", font.getFontName());
                assertEquals(14, font.getFontHeightInPoints());
            }
        }
    }

    @Test
    void theFilePasswordAndTheStructureProtectionStillReachTheOutput() throws Exception {
        var out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create().password("secret")) {
            workbook.<String>sheet("S").column("v", s -> s).write(Stream.of("a"));
            workbook.finish().writeTo(out);
        }
        assertThrows(Exception.class, () -> new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray())).close(),
                "an encrypted file is not a plain workbook");
        try (var wb = org.apache.poi.ss.usermodel.WorkbookFactory.create(new ByteArrayInputStream(out.toByteArray()), "secret")) {
            assertEquals(1, wb.getNumberOfSheets());
        }
    }
}
