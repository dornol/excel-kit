package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Multi-level (N-depth) group headers via {@code group(String... levels)}.
 * Levels are top-aligned: index 0 = outermost, last = closest to column header.
 */
class MultiLevelGroupHeaderTest {

    private static byte[] write(List<String[]> rows, ColumnSpec... specs) throws Exception {
        var writer = ExcelWriter.<String[]>create();
        for (ColumnSpec s : specs) {
            writer.column(s.name, r -> r[s.idx], c -> c.group(s.groups));
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(rows.stream()).writeTo(out);
        return out.toByteArray();
    }

    private record ColumnSpec(String name, int idx, String... groups) {}

    private static Optional<CellRangeAddress> mergedRegionContaining(XSSFSheet sheet, int row, int col) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);
            if (r.isInRange(row, col)) return Optional.of(r);
        }
        return Optional.empty();
    }

    @Nested
    class SingleLevel {
        @Test
        void existingBehaviorPreserved() throws Exception {
            byte[] bytes = write(
                    List.<String[]>of(new String[]{"1", "2"}),
                    new ColumnSpec("Price", 0, "Financial"),
                    new ColumnSpec("Qty",   1, "Financial"));

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Financial", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("Price", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("Qty",   sheet.getRow(1).getCell(1).getStringCellValue());

                // Horizontal merge of Financial across cols 0-1 on row 0
                var merge = mergedRegionContaining(sheet, 0, 0).orElseThrow();
                assertEquals(0, merge.getFirstRow());
                assertEquals(0, merge.getLastRow());
                assertEquals(0, merge.getFirstColumn());
                assertEquals(1, merge.getLastColumn());
            }
        }
    }

    @Nested
    class TwoLevels {
        @Test
        void nestedLevelsMergeIndependently() throws Exception {
            byte[] bytes = write(
                    List.<String[]>of(new String[]{"1", "2", "3", "4"}),
                    new ColumnSpec("Q1", 0, "Financial", "Revenue"),
                    new ColumnSpec("Q2", 1, "Financial", "Revenue"),
                    new ColumnSpec("Q3", 2, "Financial", "Profit"),
                    new ColumnSpec("Q4", 3, "Financial", "Profit"));

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                var sheet = wb.getSheetAt(0);
                // Row 0: Financial spans 0-3
                assertEquals("Financial", sheet.getRow(0).getCell(0).getStringCellValue());
                var m0 = mergedRegionContaining(sheet, 0, 0).orElseThrow();
                assertEquals(0, m0.getFirstColumn());
                assertEquals(3, m0.getLastColumn());

                // Row 1: Revenue spans 0-1, Profit spans 2-3
                assertEquals("Revenue", sheet.getRow(1).getCell(0).getStringCellValue());
                var m1a = mergedRegionContaining(sheet, 1, 0).orElseThrow();
                assertEquals(0, m1a.getFirstColumn());
                assertEquals(1, m1a.getLastColumn());

                assertEquals("Profit", sheet.getRow(1).getCell(2).getStringCellValue());
                var m1b = mergedRegionContaining(sheet, 1, 2).orElseThrow();
                assertEquals(2, m1b.getFirstColumn());
                assertEquals(3, m1b.getLastColumn());

                // Row 2: column names
                assertEquals("Q1", sheet.getRow(2).getCell(0).getStringCellValue());
                assertEquals("Q4", sheet.getRow(2).getCell(3).getStringCellValue());

                // Data row is row 3
                assertEquals("1", sheet.getRow(3).getCell(0).getStringCellValue());
            }
        }
    }

    @Nested
    class MixedDepth {
        @Test
        void shallowerColumnVerticallyMergesWithColumnHeader() throws Exception {
            byte[] bytes = write(
                    List.<String[]>of(new String[]{"n", "1", "2", "3"}),
                    new ColumnSpec("Name",   0),                            // no group
                    new ColumnSpec("Q1",     1, "Financial", "Revenue"),
                    new ColumnSpec("Q2",     2, "Financial", "Revenue"),
                    new ColumnSpec("Profit", 3, "Financial"));               // 1 level only

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                var sheet = wb.getSheetAt(0);

                // Name: no groups → full vertical merge rows 0-2 at col 0
                var mName = mergedRegionContaining(sheet, 0, 0).orElseThrow();
                assertEquals(0, mName.getFirstRow());
                assertEquals(2, mName.getLastRow());
                assertEquals(0, mName.getFirstColumn());
                assertEquals(0, mName.getLastColumn());
                assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());

                // Financial on row 0 spans cols 1-3
                var mFin = mergedRegionContaining(sheet, 0, 1).orElseThrow();
                assertEquals(1, mFin.getFirstColumn());
                assertEquals(3, mFin.getLastColumn());
                assertEquals("Financial", sheet.getRow(0).getCell(1).getStringCellValue());

                // Revenue on row 1 spans cols 1-2 only
                var mRev = mergedRegionContaining(sheet, 1, 1).orElseThrow();
                assertEquals(1, mRev.getFirstColumn());
                assertEquals(2, mRev.getLastColumn());

                // Profit: shallow (1 level). Row 1 at col 3 is null → vertically merged with row 2.
                var mProfit = mergedRegionContaining(sheet, 1, 3).orElseThrow();
                assertEquals(1, mProfit.getFirstRow());
                assertEquals(2, mProfit.getLastRow());
                assertEquals(3, mProfit.getFirstColumn());
                assertEquals(3, mProfit.getLastColumn());
                assertEquals("Profit", sheet.getRow(1).getCell(3).getStringCellValue());

                // Column header row (row 2) column values
                assertEquals("Q1", sheet.getRow(2).getCell(1).getStringCellValue());
                assertEquals("Q2", sheet.getRow(2).getCell(2).getStringCellValue());
            }
        }
    }

    @Nested
    class ThreeLevels {
        @Test
        void threeLevelsRenderCorrectly() throws Exception {
            byte[] bytes = write(
                    List.<String[]>of(new String[]{"1"}),
                    new ColumnSpec("Q1", 0, "Financial", "Revenue", "2025"));

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Financial", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("Revenue",   sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("2025",      sheet.getRow(2).getCell(0).getStringCellValue());
                assertEquals("Q1",        sheet.getRow(3).getCell(0).getStringCellValue());
                assertEquals("1",         sheet.getRow(4).getCell(0).getStringCellValue());
            }
        }
    }

    @Nested
    class EdgeCases {
        @Test
        void emptyVarargsIsTreatedAsNoGroup() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .column("Name", s -> s, c -> c.group())
                    .write(Stream.of("x"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Only 1 header row — no extra group row
                assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("x",    sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals(0, sheet.getNumMergedRegions());
            }
        }

        @Test
        void nullVarargsIsTreatedAsNoGroup() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .column("Name", s -> s, c -> c.group((String[]) null))
                    .write(Stream.of("x"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
            }
        }

        @Test
        void headerCommentOnShallowColumn_attachesToTopOfMergedRegion() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .column("Name", s -> s, c -> c.headerComment("format hint"))
                    .column("Q1",   s -> s, c -> c.group("Financial", "Revenue"))
                    .write(Stream.of("x"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // "Name" is vertically merged 0-2 col 0; comment should be on row 0
                var comment = sheet.getRow(0).getCell(0).getCellComment();
                assertNotNull(comment);
                assertEquals("format hint", comment.getString().getString());
                // row 2 (column header row) cell should NOT carry the comment
                assertNull(sheet.getRow(2).getCell(0).getCellComment());
            }
        }
    }
}
