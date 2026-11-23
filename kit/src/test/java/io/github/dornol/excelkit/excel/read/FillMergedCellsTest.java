package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadResult;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

class FillMergedCellsTest {

    static class Row {
        String dept;
        String name;
        double amount;
    }

    /**
     * Header in row 1; Sales merged over rows 2-4 (A2:A4), Dev over rows 5-6 (A5:A6); the amount of Sales is
     * merged over B2:B3 as a number.
     */
    private static byte[] workbook(Consumer<org.apache.poi.ss.usermodel.Sheet> extra) throws IOException {
        try (var wb = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
            var sheet = wb.createSheet("S");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("Dept");
            header.createCell(1).setCellValue("Name");
            header.createCell(2).setCellValue("Amount");
            String[][] rows = {{"Sales", "Ann"}, {null, "Bob"}, {null, "Cy"}, {"Dev", "Di"}, {null, "Ed"}};
            for (int i = 0; i < rows.length; i++) {
                var row = sheet.createRow(i + 1);
                if (rows[i][0] != null) row.createCell(0).setCellValue(rows[i][0]);
                row.createCell(1).setCellValue(rows[i][1]);
            }
            sheet.getRow(1).createCell(2).setCellValue(12.5);
            sheet.addMergedRegion(CellRangeAddress.valueOf("A2:A4"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("A5:A6"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("C2:C3"));
            extra.accept(sheet);
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static ExcelReader<Row> reader() {
        return ExcelReader.setter(Row::new)
                .column((r, c) -> r.dept = c.asString())
                .column((r, c) -> r.name = c.asString())
                .column((r, c) -> r.amount = c.asDouble());
    }

    private static List<ReadResult<Row>> read(ExcelReader<Row> reader, byte[] file) {
        var results = new ArrayList<ReadResult<Row>>();
        reader.read(new ByteArrayInputStream(file), results::add);
        return results;
    }

    @Nested
    class Filling {
        @Test
        void withoutTheOptionOnlyTheTopLeftCellHasTheValue() throws IOException {
            var results = read(ExcelReader.setter(Row::new)
                    .column((r, c) -> r.dept = c.asString()), workbook(s -> {}));
            assertEquals("Sales", results.get(0).data().dept);
            assertEquals("", results.get(1).data().dept);
        }

        @Test
        void everyRowOfARegionReadsTheValue() throws IOException {
            var results = read(reader().fillMergedCells(), workbook(s -> {}));
            assertEquals(5, results.size());
            assertEquals(List.of("Sales", "Sales", "Sales", "Dev", "Dev"),
                    results.stream().map(r -> r.data().dept).toList());
            assertEquals(List.of("Ann", "Bob", "Cy", "Di", "Ed"), results.stream().map(r -> r.data().name).toList());
        }

        @Test
        void aMergedNumberKeepsItsStoredValue() throws IOException {
            var results = read(reader().fillMergedCells(), workbook(s -> {}));
            assertEquals(12.5, results.get(0).data().amount);
            assertEquals(12.5, results.get(1).data().amount);
        }

        @Test
        void aCellWithAValueOfItsOwnKeepsIt() throws IOException {
            var file = workbook(s -> s.getRow(2).createCell(0).setCellValue("Other"));
            var results = read(reader().fillMergedCells(), file);
            assertEquals("Other", results.get(1).data().dept);
            assertEquals("Sales", results.get(2).data().dept);
        }

        @Test
        void aRegionSpanningColumnsFillsThemToo() throws IOException {
            var file = workbook(s -> {
                s.getRow(6 - 1).createCell(2).setCellValue(7);   // row 6: amount of Ed
                s.addMergedRegion(CellRangeAddress.valueOf("B6:C6"));
            });
            var results = read(ExcelReader.setter(Row::new)
                    .column((r, c) -> r.dept = c.asString())
                    .column((r, c) -> r.name = c.asString())
                    .column((r, c) -> r.amount = c.asDouble()).fillMergedCells(), file);
            assertEquals("Ed", results.get(4).data().name);
            assertEquals(7, results.get(4).data().amount, "C6 has a value of its own, which the merge does not replace");
        }

        @Test
        void theHeaderRowsAreNotChanged() throws IOException {
            var file = workbook(s -> s.addMergedRegion(CellRangeAddress.valueOf("A1:B1")));
            var results = read(reader().fillMergedCells(), file);
            assertEquals(5, results.size());
            assertEquals("Sales", results.get(0).data().dept);
        }

        @Test
        void aSheetWithoutMergedRegionsReadsAsBefore() throws IOException {
            try (var wb = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
                var sheet = wb.createSheet("S");
                sheet.createRow(0).createCell(0).setCellValue("Dept");
                sheet.createRow(1).createCell(0).setCellValue("Sales");
                wb.write(out);
                var results = read(ExcelReader.setter(Row::new).column((r, c) -> r.dept = c.asString())
                        .fillMergedCells(), out.toByteArray());
                assertEquals(1, results.size());
                assertEquals("Sales", results.get(0).data().dept);
            }
        }

        @Test
        void worksInMappingAndMapMode() throws IOException {
            var file = workbook(s -> {});
            var mapped = new ArrayList<String>();
            ExcelReader.<String>mapping(row -> row.get("Dept").asString()).fillMergedCells()
                    .read(new ByteArrayInputStream(file), r -> mapped.add(r.data()));
            assertEquals(List.of("Sales", "Sales", "Sales", "Dev", "Dev"), mapped);

            var maps = new ArrayList<Object>();
            ExcelReader.forMap().fillMergedCells()
                    .read(new ByteArrayInputStream(file), r -> maps.add(r.data().get("Dept")));
            assertEquals(List.of("Sales", "Sales", "Sales", "Dev", "Dev"), maps);
        }

        @Test
        void aMergedRowIsNotSkippedAsBlank() throws IOException {
            // B and C are empty in row 3 of a sheet that only has A merged: with the fill the row has a value
            try (var wb = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
                var sheet = wb.createSheet("S");
                sheet.createRow(0).createCell(0).setCellValue("Dept");
                sheet.createRow(1).createCell(0).setCellValue("Sales");
                sheet.createRow(2).createCell(1);   // a styled-but-empty cell keeps the row in the file
                sheet.addMergedRegion(CellRangeAddress.valueOf("A2:A3"));
                wb.write(out);
                var results = read(ExcelReader.setter(Row::new).column((r, c) -> r.dept = c.asString())
                        .fillMergedCells(), out.toByteArray());
                assertEquals(List.of("Sales", "Sales"), results.stream().map(r -> r.data().dept).toList());
            }
        }
    }
}
