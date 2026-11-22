package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** A null value is an empty cell, not a cell holding an empty string: ISBLANK and COUNTA see the difference. */
class NullCellTest {

    private static byte[] write() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s)
                .column("Note", s -> s.equals("b") ? "has note" : null, c -> c.backgroundColor(200, 220, 255))
                .write(List.of("a", "b")).writeTo(out);
        return out.toByteArray();
    }

    @Test
    void nullValue_isABlankCellThatKeepsTheColumnStyle() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(write()))) {
            Cell empty = wb.getSheetAt(0).getRow(1).getCell(1);
            Cell filled = wb.getSheetAt(0).getRow(2).getCell(1);

            assertEquals(CellType.BLANK, empty.getCellType());
            assertEquals(CellType.STRING, filled.getCellType());
            assertEquals(filled.getCellStyle().getFillForegroundColorColor(),
                    empty.getCellStyle().getFillForegroundColorColor(), "the blank cell keeps the column's fill");
            assertNotEquals(0, empty.getCellStyle().getIndex());
        }
    }

    @Test
    void nullValue_readsBackAsAnEmptyValue() {
        List<Map<String, String>> rows = ExcelReader.forMap().readAll(new ByteArrayInputStream(write()));

        assertEquals(2, rows.size());
        String note = rows.get(0).get("Note");
        assertTrue(note == null || note.isEmpty(), "was: " + note);
        assertEquals("has note", rows.get(1).get("Note"));
    }
}
