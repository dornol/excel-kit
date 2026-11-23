package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

/** Subtotal rows next to the rest of the writer: streaming options, other writers, headers, callbacks, formulas, limits. */
class GroupSubtotalCombinationsTest {
    record Row(String dept, String name, int amount) {}

    private static final List<Row> ROWS = List.of(
            new Row("A", "x", 10), new Row("A", "y", 20), new Row("B", "z", 5), new Row("B", "w", 7));

    private static ExcelWriter<Row> writer(ExcelWriter<Row> w) {
        return w.column("Dept", Row::dept).column("Name", Row::name).column("Amount", Row::amount)
                .groupSubtotal(Row::dept, s -> s.label("Name", k -> k + " total").sum("Amount"));
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    /** Every row of every sheet as text: formulas as formulas. */
    private static List<String> dump(XSSFWorkbook wb) {
        List<String> out = new ArrayList<>();
        for (int s = 0; s < wb.getNumberOfSheets(); s++) {
            for (var row : wb.getSheetAt(s)) {
                var sb = new StringBuilder(wb.getSheetName(s) + "#" + row.getRowNum() + ":");
                for (var cell : row) {
                    sb.append(' ').append(switch (cell.getCellType()) {
                        case FORMULA -> "=" + cell.getCellFormula();
                        case NUMERIC -> String.valueOf((long) cell.getNumericCellValue());
                        case STRING -> cell.getStringCellValue();
                        default -> "";
                    });
                }
                out.add(sb.toString());
            }
        }
        return out;
    }

    @Nested
    class StreamingOptions {
        @Test
        void aWindowOfOneRowSharedStringsAndCompressedTempFilesWriteTheSameSheet() throws IOException {
            List<Row> rows = IntStream.range(0, 300).mapToObj(i -> new Row("K" + i / 7, "n" + i, i)).toList();
            List<String> plain;
            try (var wb = open(writer(ExcelWriter.<Row>create()).maxRows(100).write(rows))) {
                plain = dump(wb);
            }
            try (var wb = open(writer(ExcelWriter.<Row>create(o -> o.rowAccessWindowSize(1).useSharedStrings(true)
                    .compressTempFiles(true))).maxRows(100).write(rows))) {
                assertEquals(plain, dump(wb));
            }
        }
    }

    @Nested
    class OtherWriters {
        @Test
        void aMapWriterTakesSubtotalsOfItsKeys() throws IOException {
            var w = ExcelWriter.forMap("Region", "Amount")
                    .groupSubtotal(m -> m.get("Region"), s -> s.label("Region", k -> k + " total").sum("Amount"));
            List<Map<String, Object>> rows = List.of(Map.of("Region", "East", "Amount", 1), Map.of("Region", "East", "Amount", 2),
                    Map.of("Region", "West", "Amount", 4));
            try (var wb = open(w.write(rows))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("East total", sheet.getRow(3).getCell(0).getStringCellValue());
                assertEquals("SUBTOTAL(9,B2:B3)", sheet.getRow(3).getCell(1).getCellFormula());
                assertEquals("West total", sheet.getRow(5).getCell(0).getStringCellValue());
            }
        }

        @Test
        void writeToAPathAndAnIterableWork() throws IOException {
            Path file = Files.createTempFile("subtotal", ".xlsx");
            try {
                writer(ExcelWriter.<Row>create()).write((Iterable<Row>) ROWS).writeTo(file);
                try (var wb = new XSSFWorkbook(Files.newInputStream(file))) {
                    assertEquals("SUBTOTAL(9,C2:C3)", wb.getSheetAt(0).getRow(3).getCell(2).getCellFormula());
                }
            } finally {
                Files.deleteIfExists(file);
            }
        }

        @Test
        void aHiddenWorkbookSheetKeepsItsSubtotals() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("Visible").column("Dept", Row::dept).column("Amount", Row::amount)
                        .groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount")).write(ROWS);
                workbook.<Row>sheet("Hidden").column("Dept", Row::dept).column("Amount", Row::amount)
                        .groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount")).hideSheet().write(ROWS);
                try (var wb = open(workbook.finish())) {
                    assertTrue(wb.isSheetHidden(1));
                    assertEquals("SUBTOTAL(9,B2:B3)", wb.getSheetAt(1).getRow(3).getCell(1).getCellFormula());
                }
            }
        }
    }

    @Nested
    class HeadersAndCallbacks {
        @Test
        void everyRolledOverSheetHasItsOwnPreambleAndGroupHeaderAboveItsRanges() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .beforeHeader(ctx -> {
                        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Title");
                        return ctx.getCurrentRow() + 2;
                    })
                    .column("Dept", Row::dept, c -> c.group("G"))
                    .column("Name", Row::name, c -> c.group("G"))
                    .column("Amount", Row::amount, c -> c.group("G"))
                    .maxRows(2)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> k + " total").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(2, wb.getNumberOfSheets());
                for (int sheet = 0; sheet < 2; sheet++) {
                    var s = wb.getSheetAt(sheet);
                    assertEquals("Title", s.getRow(0).getCell(0).getStringCellValue());
                    assertEquals("Dept", s.getRow(3).getCell(0).getStringCellValue(), "header below the preamble and the group row");
                    // data rows 5 and 6 (1-based), the subtotal row 7 covers exactly them
                    assertEquals("SUBTOTAL(9,C5:C6)", s.getRow(6).getCell(2).getCellFormula(), "sheet " + sheet);
                }
            }
        }

        @Test
        void afterDataComesBelowTheLastSubtotalOfEverySheetAndAboveItsSummary() throws IOException {
            var w = writer(ExcelWriter.<Row>create()).maxRows(2)
                    .afterData(ctx -> {
                        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("after");
                        return ctx.getCurrentRow() + 1;
                    })
                    .summary(s -> s.label("Name", "Total").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                for (int sheet = 0; sheet < 2; sheet++) {
                    var s = wb.getSheetAt(sheet);
                    assertEquals("after", s.getRow(4).getCell(0).getStringCellValue(), "sheet " + sheet);
                    assertEquals("Total", s.getRow(5).getCell(1).getStringCellValue());
                    assertEquals("SUBTOTAL(9,C2:C4)", s.getRow(5).getCell(2).getCellFormula(), "the summary covers the data and the subtotal, not the afterData row");
                }
            }
        }
    }

    @Nested
    class Formulas {
        @Test
        void aRowFormulaThatUsesTheCursorRowPointsAtItsOwnRowBelowASubtotalRow() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .column("Dept", Row::dept)
                    .column("Amount", Row::amount)
                    .column("Double", (row, cursor) -> "B" + cursor.getRowOfSheet() + "*2", c -> c.type(ExcelDataType.FORMULA))
                    .groupSubtotal(Row::dept, s -> s.label("Dept", k -> k + " total").sum("Amount").sum("Double"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                int[] dataRows = {1, 2, 4, 5}; // 0-based; the subtotal rows are 3 and 6
                for (int r : dataRows) {
                    assertEquals("B" + (r + 1) + "*2", sheet.getRow(r).getCell(2).getCellFormula(), "row " + (r + 1));
                }
                assertEquals("SUBTOTAL(9,C2:C3)", sheet.getRow(3).getCell(2).getCellFormula(), "a subtotal over a formula column");
                var evaluator = wb.getCreationHelper().createFormulaEvaluator();
                assertEquals(60, evaluator.evaluate(sheet.getRow(3).getCell(2)).getNumberValue());
                assertEquals(24, evaluator.evaluate(sheet.getRow(6).getCell(2)).getNumberValue());
            }
        }
    }

    @Nested
    class Rules {
        @Test
        void aDependentDropdownKeepsItsHelperSheetAndRunsDownThroughTheSubtotalRows() throws IOException {
            Map<String, List<String>> items = Map.of("A", List.of("x", "y"), "B", List.of("z", "w"));
            var w = ExcelWriter.<Row>create()
                    .column("Dept", Row::dept)
                    .column("Name", Row::name, c -> c.dependentDropdown("Dept", items))
                    .column("Amount", Row::amount)
                    .maxRows(2)
                    .groupSubtotal(Row::dept, s -> s.label("Dept", k -> k + " total").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(3, wb.getNumberOfSheets(), "two data sheets and the helper sheet");
                assertTrue(wb.isSheetHidden(2));
                for (int sheet = 0; sheet < 2; sheet++) {
                    assertFalse(wb.getSheetAt(sheet).getDataValidations().isEmpty(), "sheet " + sheet);
                }
            }
        }

        @Test
        void aDataBarAndAnIconSetCoverTheSubtotalRowsToo() throws IOException {
            var w = writer(ExcelWriter.<Row>create())
                    .conditionalFormatting(c -> c.columns(2).dataBar(ExcelColor.STEEL_BLUE));
            try (var wb = open(w.write(ROWS))) {
                var range = wb.getSheetAt(0).getSheetConditionalFormatting().getConditionalFormattingAt(0).getFormattingRanges()[0];
                assertEquals("C2:C7", range.formatAsString(), "the bars scale over the subtotals as well (documented)");
            }
        }

        @Test
        void aLockedByDefaultStyleStillLocksTheSubtotalRowOfAnEditableColumn() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .defaultStyle(d -> d.locked(false))
                    .protectSheet("pw")
                    .column("Dept", Row::dept).column("Name", Row::name).column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> "t").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                assertFalse(sheet.getRow(1).getCell(2).getCellStyle().getLocked());
                for (int col = 0; col < 3; col++) assertTrue(sheet.getRow(3).getCell(col).getCellStyle().getLocked());
            }
        }
    }

    @Nested
    class Scale {
        @Test
        void manyGroupsKeepEveryRangeRight() throws IOException {
            int rows = 30_000;
            var list = IntStream.range(0, rows).mapToObj(i -> new Row("K" + i / 7, "n" + i, 1)).toList();
            try (var wb = open(writer(ExcelWriter.<Row>create()).write(list))) {
                var sheet = wb.getSheetAt(0);
                int groups = (rows + 6) / 7;
                assertEquals(rows + groups, sheet.getLastRowNum(), "header excluded: every data row and every subtotal row");
                int row = 1; // 0-based index of the first data row of the next group
                for (int g = 0; g < groups; g++) {
                    int size = Math.min(7, rows - g * 7);
                    int subtotalRow = row + size;
                    assertEquals("SUBTOTAL(9,C" + (row + 1) + ":C" + (subtotalRow) + ")",
                            sheet.getRow(subtotalRow).getCell(2).getCellFormula(), "group " + g);
                    row = subtotalRow + 1;
                }
            }
        }

        @Test
        void moreThan1026PageBreaksFailLikeWithoutSubtotals() {
            var list = IntStream.range(0, 1_100).mapToObj(i -> new Row("K" + i, "n", 1)).toList();
            var w = writer(ExcelWriter.<Row>create()).pageBreakOnChange(Row::dept);
            var e = assertThrows(ExcelWriteException.class, () -> w.write(list));
            assertTrue(e.getMessage().contains("page breaks"), e.getMessage());
        }
    }

    @Nested
    class Rejections {
        @Test
        void aLabelColumnLeftOutByColumnIfIsNamed() {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).columnIf("Name", false, Row::name)
                    .column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> "t").sum("Amount"));
            var e = assertThrows(ExcelWriteException.class, () -> w.write(ROWS));
            assertTrue(e.getMessage().contains("Name"), e.getMessage());
        }

        @Test
        void aSubtotalColumnLeftOutByColumnIfIsNamed() {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).columnIf("Amount", false, Row::amount)
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            var e = assertThrows(ExcelWriteException.class, () -> w.write(ROWS));
            assertTrue(e.getMessage().contains("Amount"), e.getMessage());
        }
    }
}
