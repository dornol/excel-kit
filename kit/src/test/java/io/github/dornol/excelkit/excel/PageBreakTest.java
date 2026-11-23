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
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class PageBreakTest {

    private record Emp(String dept, int no) {}

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static List<Emp> emps(String dept, int count) {
        return IntStream.range(0, count).mapToObj(i -> new Emp(dept, i)).toList();
    }

    @Nested
    class RowBreakEvery {
        @Test
        void breaksAfterEveryNDataRows() throws IOException {
            // header is row 0, data rows 1..10; a break above row 4 and above row 7 and above row 10
            try (var wb = open(ExcelWriter.<Emp>create().column("No", Emp::no)
                    .printSetup(p -> p.rowBreakEvery(3))
                    .write(emps("a", 10)))) {
                assertArrayEquals(new int[]{4, 7, 10}, wb.getSheetAt(0).getRowBreaks());
            }
        }

        @Test
        void noBreakWhenTheRowsFitOnePage() throws IOException {
            try (var wb = open(ExcelWriter.<Emp>create().column("No", Emp::no)
                    .printSetup(p -> p.rowBreakEvery(5)).write(emps("a", 5)))) {
                assertEquals(0, wb.getSheetAt(0).getRowBreaks().length);
            }
        }

        @Test
        void eachRolloverSheetCountsItsOwnRows() throws IOException {
            try (var wb = open(ExcelWriter.<Emp>create().maxRows(4).column("No", Emp::no)
                    .printSetup(p -> p.rowBreakEvery(2)).write(emps("a", 8)))) {
                assertEquals(2, wb.getNumberOfSheets());
                for (int i = 0; i < 2; i++) assertArrayEquals(new int[]{3}, wb.getSheetAt(i).getRowBreaks(), "sheet " + i);
            }
        }

        @Test
        void rejectsLessThanOne() {
            assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().rowBreakEvery(0));
        }

        @Test
        void moreBreaksThanExcelKeepsIsRejected() {
            assertThrows(ExcelWriteException.class, () -> ExcelWriter.<Emp>create().column("No", Emp::no)
                    .printSetup(p -> p.rowBreakEvery(1)).write(emps("a", 1200)));
        }
    }

    @Nested
    class PageBreakOnChange {
        @Test
        void startsANewPageWhenTheKeyChanges() throws IOException {
            var rows = new ArrayList<Emp>();
            rows.addAll(emps("a", 3));   // rows 1..3
            rows.addAll(emps("b", 2));   // rows 4..5
            rows.addAll(emps("c", 1));   // row 6
            try (var wb = open(ExcelWriter.<Emp>create().column("Dept", Emp::dept)
                    .pageBreakOnChange(Emp::dept).write(rows))) {
                assertArrayEquals(new int[]{4, 6}, wb.getSheetAt(0).getRowBreaks());
            }
        }

        @Test
        void aRolloverSheetHasNoBreakBelowItsHeader() throws IOException {
            var rows = new ArrayList<Emp>();
            rows.addAll(emps("a", 2));
            rows.addAll(emps("b", 2));   // the second sheet starts with the first "b"... maxRows 2 puts it there
            try (var wb = open(ExcelWriter.<Emp>create().maxRows(2).column("Dept", Emp::dept)
                    .pageBreakOnChange(Emp::dept).write(rows))) {
                assertEquals(2, wb.getNumberOfSheets());
                assertEquals(0, wb.getSheetAt(0).getRowBreaks().length);
                assertEquals(0, wb.getSheetAt(1).getRowBreaks().length);
            }
        }

        @Test
        void aWorkbookSheetBreaksToo() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                var rows = new ArrayList<Emp>(emps("a", 2));
                rows.addAll(emps("b", 2));
                workbook.<Emp>sheet("S").column("Dept", Emp::dept).pageBreakOnChange(Emp::dept).write(rows);
                try (var wb = open(workbook.finish())) {
                    assertArrayEquals(new int[]{3}, wb.getSheetAt(0).getRowBreaks());
                }
            }
        }

        @Test
        void addsToRowBreakEveryWithoutDuplicates() throws IOException {
            var rows = new ArrayList<Emp>(emps("a", 3));
            rows.addAll(emps("b", 3));   // data rows 1..6; the key changes above row 4; every 3 rows also breaks above 4
            try (var wb = open(ExcelWriter.<Emp>create().column("Dept", Emp::dept)
                    .printSetup(p -> p.rowBreakEvery(3)).pageBreakOnChange(Emp::dept).write(rows))) {
                assertArrayEquals(new int[]{4}, wb.getSheetAt(0).getRowBreaks());
            }
        }

        @Test
        void pageBreakWhenGetsThePreviousAndTheCurrentRow() throws IOException {
            var seen = new ArrayList<String>();
            ExcelWriter.<Emp>create().column("No", Emp::no)
                    .pageBreakWhen((p, c) -> { seen.add(p.no() + ">" + c.no()); return false; })
                    .write(emps("a", 3)).writeTo(OutputStreamNull.INSTANCE);
            assertEquals(List.of("0>1", "1>2"), seen);
        }

        @Test
        void nullArgumentsAreRejected() {
            var writer = ExcelWriter.<Emp>create();
            assertThrows(NullPointerException.class, () -> writer.pageBreakWhen(null));
            assertThrows(NullPointerException.class, () -> writer.pageBreakOnChange(null));
        }
    }

    private static final class OutputStreamNull {
        static final java.io.OutputStream INSTANCE = java.io.OutputStream.nullOutputStream();
    }
}
