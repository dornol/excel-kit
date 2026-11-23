package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadResult;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/** {@code fillMergedCells()} with the sheet it reads, the header it skips and the readers it is combined with. */
class FillMergedCellsSheetsTest {

    static class Row {
        String dept;
        String name;
    }

    private static void rows(XSSFWorkbook wb, String sheetName, String dept, int count, int firstDataRow) {
        var sheet = wb.createSheet(sheetName);
        var header = sheet.createRow(firstDataRow - 1);
        header.createCell(0).setCellValue("Dept");
        header.createCell(1).setCellValue("Name");
        for (int i = 0; i < count; i++) {
            var row = sheet.createRow(firstDataRow + i);
            if (i == 0) row.createCell(0).setCellValue(dept);
            row.createCell(1).setCellValue(sheetName + i);
        }
        sheet.addMergedRegion(new CellRangeAddress(firstDataRow, firstDataRow + count - 1, 0, 0));
    }

    private static byte[] twoSheets() throws IOException {
        try (var wb = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
            rows(wb, "First", "S0", 2, 1);   // a different merge on the first sheet: A2:A3
            rows(wb, "Second", "S1", 4, 1);  // A2:A5
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static ExcelReader<Row> reader() {
        return ExcelReader.setter(Row::new)
                .column((r, c) -> r.dept = c.asString())
                .column((r, c) -> r.name = c.asString());
    }

    private static List<String> depts(ExcelReader<Row> reader, byte[] file) {
        var out = new ArrayList<String>();
        reader.read(new ByteArrayInputStream(file), r -> out.add(r.data().dept));
        return out;
    }

    @Test
    void theMergedRegionsAreThoseOfTheSelectedSheet() throws IOException {
        assertEquals(List.of("S0", "S0"), depts(reader().fillMergedCells(), twoSheets()));
        assertEquals(List.of("S1", "S1", "S1", "S1"), depts(reader().sheetIndex(1).fillMergedCells(), twoSheets()));
    }

    @Test
    void aWorkbookReaderFillsOnlyTheSheetsWhoseReaderAsksForIt() throws IOException {
        var filled = new ArrayList<String>();
        var plain = new ArrayList<String>();
        ExcelWorkbookReader.create()
                .sheet("First", reader().fillMergedCells(), r -> filled.add(r.data().dept))
                .sheet("Second", reader(), r -> plain.add(r.data().dept))
                .read(new ByteArrayInputStream(twoSheets()));
        assertEquals(List.of("S0", "S0"), filled);
        assertEquals(List.of("S1", "", "", ""), plain);
    }

    @Test
    void aWorkbookReaderFillsEachSheetFromItsOwnRegions() throws IOException {
        var first = new ArrayList<String>();
        var second = new ArrayList<String>();
        ExcelWorkbookReader.create()
                .sheet("First", reader().fillMergedCells(), r -> first.add(r.data().dept))
                .sheet("Second", reader().fillMergedCells(), r -> second.add(r.data().dept))
                .read(new ByteArrayInputStream(twoSheets()));
        assertEquals(List.of("S0", "S0"), first);
        assertEquals(List.of("S1", "S1", "S1", "S1"), second);
    }

    @Test
    void aSheetThatIsNotThereStillFailsTheWayItAlwaysDid() throws IOException {
        assertThrows(RuntimeException.class, () -> depts(reader().sheetIndex(5).fillMergedCells(), twoSheets()));
    }

    @Test
    void worksWithARowCountPreScanAndWithProgressReporting() throws IOException {
        var progress = new ArrayList<Long>();
        var depts = new ArrayList<String>();
        reader().sheetIndex(1).fillMergedCells().countRows()
                .onProgress(1, (processed, cursor) -> progress.add(cursor.getTotalRows()))
                .read(new ByteArrayInputStream(twoSheets()), r -> depts.add(r.data().dept));
        assertEquals(List.of("S1", "S1", "S1", "S1"), depts);
        assertEquals(List.of(4L, 4L, 4L, 4L), progress);
    }

    @Test
    void aRegionWhoseTopLeftCellIsEmptyGivesNothingAndDoesNotFail() throws IOException {
        try (var wb = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
            var sheet = wb.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("Dept");
            sheet.getRow(0).createCell(1).setCellValue("Name");
            sheet.createRow(1).createCell(1).setCellValue("a");
            sheet.createRow(2).createCell(1).setCellValue("b");
            sheet.addMergedRegion(CellRangeAddress.valueOf("A2:A3"));   // A2 holds nothing
            wb.write(out);
            assertEquals(List.of("", ""), depts(reader().fillMergedCells(), out.toByteArray()));
        }
    }

    @Test
    void aRegionStartingInTheHeaderAndReachingIntoTheDataIsFilledInTheDataRows() throws IOException {
        // two header rows (headerRows(2), headerRowIndex(1)); "Dept" is merged over A1:A4: header row 1, header row 2 and two data rows
        try (var wb = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
            var sheet = wb.createSheet("S");
            var h1 = sheet.createRow(0);
            h1.createCell(0).setCellValue("Dept");
            h1.createCell(1).setCellValue("Name");
            sheet.createRow(1);
            sheet.createRow(2).createCell(1).setCellValue("a");
            sheet.createRow(3).createCell(1).setCellValue("b");
            sheet.addMergedRegion(CellRangeAddress.valueOf("A1:A4"));
            wb.write(out);
            var results = new ArrayList<ReadResult<Row>>();
            reader().headerRows(2).headerRowIndex(1).fillMergedCells()
                    .read(new ByteArrayInputStream(out.toByteArray()), results::add);
            assertEquals(2, results.size());
            assertEquals(List.of("Dept", "Dept"), results.stream().map(r -> r.data().dept).toList());
            assertEquals(List.of("a", "b"), results.stream().map(r -> r.data().name).toList());
        }
    }
}
