package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class MergeSameValuesTest {

    private record Emp(String dept, String team, String name) {}

    private static List<Emp> staff() {
        return List.of(
                new Emp("Sales", "East", "Ann"), new Emp("Sales", "East", "Bob"), new Emp("Sales", "West", "Cy"),
                new Emp("Dev", "West", "Di"), new Emp("Dev", "West", "Ed"), new Emp("Ops", "East", "Flo"));
    }

    private static byte[] bytes(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return out.toByteArray();
    }

    private static List<String> regions(byte[] file, int sheet) throws IOException {
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
            return wb.getSheetAt(sheet).getMergedRegions().stream().map(CellRangeAddress::formatAsString).sorted().toList();
        }
    }

    private static ExcelWriter<Emp> writer() {
        return ExcelWriter.<Emp>create()
                .column("Dept", Emp::dept, c -> c.mergeSameValues())
                .column("Team", Emp::team, c -> c.mergeSameValues())
                .column("Name", Emp::name);
    }

    @Nested
    class Merging {
        @Test
        void mergesRunsOfEqualValuesAndNestsTheSecondColumnInsideTheFirst() throws IOException {
            // rows: 2-4 Sales (East 2-3, West 4), 5-6 Dev (West 5-6), 7 Ops (East 7)
            var file = bytes(writer().write(staff()));
            assertEquals(List.of("A2:A4", "A5:A6", "B2:B3", "B5:B6"), regions(file, 0));
        }

        @Test
        void aSingleRowRunIsNotMerged() throws IOException {
            var file = bytes(writer().write(List.of(new Emp("Sales", "East", "Ann"), new Emp("Dev", "West", "Bob"))));
            assertEquals(List.of(), regions(file, 0));
        }

        @Test
        void theCellsBelowTheFirstKeepTheirValues() throws IOException {
            var file = bytes(writer().write(staff()));
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Sales", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("Sales", sheet.getRow(2).getCell(0).getStringCellValue());
                assertEquals("Sales", sheet.getRow(3).getCell(0).getStringCellValue());
            }
        }

        @Test
        void emptyValuesEndARunAndAreNotMerged() throws IOException {
            var rows = List.of(new Emp("A", "x", "1"), new Emp("A", "x", "2"), new Emp(null, "x", "3"),
                    new Emp(null, "x", "4"), new Emp("A", "x", "5"), new Emp("A", "x", "6"));
            var file = bytes(writer().write(rows));
            assertEquals(List.of("A2:A3", "A6:A7", "B2:B3", "B6:B7"), regions(file, 0));
        }

        @Test
        void aRowsTheSameValueAfterAnotherDoesNotJoinNonNeighbours() throws IOException {
            var rows = List.of(new Emp("A", "x", "1"), new Emp("B", "x", "2"), new Emp("A", "x", "3"), new Emp("A", "x", "4"));
            var file = bytes(writer().write(rows));
            assertEquals(List.of("A4:A5", "B4:B5"), regions(file, 0));
        }

        @Test
        void aRolloverSheetStartsItsRunsAgain() throws IOException {
            var rows = new ArrayList<Emp>();
            for (int i = 0; i < 6; i++) rows.add(new Emp("Sales", "East", "n" + i));
            var file = bytes(writer().maxRows(4).write(rows));
            assertEquals(List.of("A2:A5", "B2:B5"), regions(file, 0));
            assertEquals(List.of("A2:A3", "B2:B3"), regions(file, 1));
        }

        @Test
        void aWorkbookSheetMergesToo() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Emp>sheet("S")
                        .column("Dept", Emp::dept, c -> c.mergeSameValues())
                        .column("Name", Emp::name)
                        .write(Stream.of(new Emp("A", "", "1"), new Emp("A", "", "2")));
                assertEquals(List.of("A2:A3"), regions(bytes(workbook.finish()), 0));
            }
        }

        @Test
        void numbersMergeToo() throws IOException {
            var file = bytes(ExcelWriter.<int[]>create()
                    .column("Group", r -> r[0], c -> c.mergeSameValues())
                    .column("N", r -> r[1])
                    .write(List.of(new int[]{1, 1}, new int[]{1, 2}, new int[]{2, 3})));
            assertEquals(List.of("A2:A3"), regions(file, 0));
        }

        @Test
        void aFileReadsBackWithOrWithoutFillMergedCells() throws IOException {
            var file = bytes(writer().write(staff()));
            for (boolean fill : new boolean[]{false, true}) {
                var depts = new ArrayList<String>();
                var reader = ExcelReader.forMap();
                if (fill) reader.fillMergedCells();
                reader.read(new ByteArrayInputStream(file), r -> depts.add(r.data().get("Dept")));
                assertEquals(List.of("Sales", "Sales", "Sales", "Dev", "Dev", "Ops"), depts, "fill=" + fill);
            }
        }

        @Test
        void emptyDataWritesNothingToMerge() throws IOException {
            assertEquals(List.of(), regions(bytes(writer().write(List.<Emp>of())), 0));
        }
    }

    @Nested
    class Rejections {
        @Test
        void aTableCannotHoldMergedCells() {
            assertThrows(ExcelWriteException.class, () -> writer().table("People").write(staff()));
        }

        @Test
        void aTemplateListWriterDoesNotMerge() throws IOException {
            var template = new ByteArrayOutputStream();
            try (var wb = new XSSFWorkbook()) {
                wb.createSheet("S");
                wb.write(template);
            }
            try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template.toByteArray()))) {
                var list = writer.<Emp>list(0, 0);
                list.column("Dept", Emp::dept, c -> c.mergeSameValues());
                assertThrows(ExcelWriteException.class, () -> list.write(staff().stream()));
            }
        }
    }
}
