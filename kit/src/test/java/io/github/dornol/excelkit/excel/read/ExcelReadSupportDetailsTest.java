package io.github.dornol.excelkit.excel.read;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** How worksheet XML written by other programs is read: dates, and cells and rows without a reference. */
class ExcelReadSupportDetailsTest {

    // ---- Excel's 1900-02-29, a day that never existed

    /** The text of a numeric cell holding {@code serial}, shown with the given number format. */
    private static String shown(double serial, String format) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("v");
            var cell = sheet.createRow(1).createCell(0);
            cell.setCellValue(serial);
            var style = workbook.createCellStyle();
            style.setDataFormat(workbook.createDataFormat().getFormat(format));
            cell.setCellStyle(style);
            workbook.write(out);
            List<String> values = new ArrayList<>();
            ExcelReader.<String>mapping(row -> row.get("v").asString())
                    .read(new ByteArrayInputStream(out.toByteArray()), result -> values.add(result.data()));
            return values.get(0);
        }
    }

    @Test
    void serial60IsTheNonexistentLeapDayWhenTheFormatShowsACalendarDate() throws IOException {
        // Excel counts 1900 as a leap year; serial 60 is its 29 February
        for (String format : new String[]{"yyyy-mm-dd", "yy", "yyyy", "d", "dd", "mmm", "mmmm", "mm", "m", "m/d", "d-mmm-yy",
                "dd/mm/yyyy hh:mm", "yyyy\"-\"mm", "[$-409]yyyy-mm-dd"}) {
            assertEquals("1900-02-29", shown(60, format), format);
        }
    }

    @Test
    void aFractionOfThatDayCarriesTheTimeOfDay() throws IOException {
        assertEquals("1900-02-29T12:00:00", shown(60.5, "yyyy-mm-dd hh:mm"));
        assertEquals("1900-02-29T06:00:00", shown(60.25, "d"));
        // the time is cut at the last millisecond of the day
        assertEquals("1900-02-29T23:59:59.999", shown(60.99999999999, "yyyy-mm-dd"));
    }

    @Test
    void formatsWithoutACalendarFieldShowTheTimeOnly() throws IOException {
        for (String format : new String[]{"h:mm", "hh:mm:ss", "mm:ss", "m:ss", "h:mm AM/PM", "[h]:mm", "[mm]:ss", "[ss]",
                "h:mm \"d\"", "h:mm \"yy\""}) {
            assertNotEquals("1900-02-29", shown(60, format), format);
            assertTrue(!shown(60, format).startsWith("1900"), format + " -> " + shown(60, format));
        }
    }

    @Test
    void onlySerialNumbersFrom60UpToButNotIncluding61AreTheLeapDay() throws IOException {
        assertNotEquals("1900-02-29", shown(59.9999, "yyyy-mm-dd"));
        assertEquals("1900-02-28", shown(59, "yyyy-mm-dd"));
        assertEquals("1900-03-01", shown(61, "yyyy-mm-dd"));
        assertNotEquals("1900-02-29", shown(61, "yyyy-mm-dd"));
        assertEquals("1900-02-29", shown(60, "yyyy-mm-dd"));
    }

    // ---- cells and rows without references

    /** A workbook whose only sheet has the given SpreadsheetML {@code sheetData}. */
    private static byte[] withSheetData(String sheetData) throws IOException {
        byte[] template;
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            workbook.createSheet("S").createRow(0).createCell(0).setCellValue("x");
            workbook.write(out);
            template = out.toByteArray();
        }
        ByteArrayOutputStream result = new ByteArrayOutputStream();
        try (ZipInputStream in = new ZipInputStream(new ByteArrayInputStream(template)); ZipOutputStream zip = new ZipOutputStream(result)) {
            for (ZipEntry entry; (entry = in.getNextEntry()) != null; ) {
                byte[] data = in.readAllBytes();
                if (entry.getName().equals("xl/worksheets/sheet1.xml")) {
                    data = ("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet "
                            + "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                            + sheetData + "</sheetData></worksheet>").getBytes(StandardCharsets.UTF_8);
                }
                zip.putNextEntry(new ZipEntry(entry.getName()));
                zip.write(data);
                zip.closeEntry();
            }
        }
        return result.toByteArray();
    }

    private static String text(String value) {
        return "<c t=\"inlineStr\"><is><t>" + value + "</t></is></c>";
    }

    @Test
    void rowsAndCellsWithoutReferencesAreNumberedInOrder() throws IOException {
        byte[] file = withSheetData("<row>" + text("name") + text("city") + "</row>"
                + "<row>" + text("ann") + text("oslo") + "</row>"
                + "<row>" + text("bob") + text("rome") + "</row>");
        List<String> rows = new ArrayList<>();
        List<Long> numbers = new ArrayList<>();
        ExcelReader.forMap().read(new ByteArrayInputStream(file), result -> {
            rows.add(result.data().toString());
            numbers.add(result.fileRowNum());
        });
        assertEquals(List.of("{name=ann, city=oslo}", "{name=bob, city=rome}"), rows);
        assertEquals(List.of(2L, 3L), numbers);
    }

    @Test
    void aRowReferenceMovesTheCounterOnAndLaterRowsFollowIt() throws IOException {
        byte[] file = withSheetData("<row r=\"1\">" + text("name") + "</row>"
                + "<row r=\"5\">" + text("ann") + "</row>"
                + "<row>" + text("bob") + "</row>");
        List<Long> numbers = new ArrayList<>();
        ExcelReader.forMap().read(new ByteArrayInputStream(file), result -> numbers.add(result.fileRowNum()));
        assertEquals(List.of(5L, 6L), numbers, "an unnumbered row follows the one before it");
    }

    @Test
    void anImplicitCellLandsRightAfterTheOneBefore() throws IOException {
        byte[] file = withSheetData("<row r=\"1\"><c r=\"B1\" t=\"inlineStr\"><is><t>second</t></is></c>"
                + text("third") + "</row>"
                + "<row r=\"2\"><c r=\"B2\" t=\"inlineStr\"><is><t>x</t></is></c>" + text("y") + "</row>");
        List<String> header = new ArrayList<>();
        List<String> values = new ArrayList<>();
        ExcelReader.<List<String>>mapping(row -> row.headerNames()).read(new ByteArrayInputStream(file),
                result -> header.addAll(result.data()));
        ExcelReader.forMap().read(new ByteArrayInputStream(file), result -> values.add(result.data().toString()));
        assertEquals(List.of("", "second", "third"), header);
        assertEquals(List.of("{second=x, third=y}"), values);
    }

    private static String cells(int count, String prefix) {
        StringBuilder cells = new StringBuilder();
        for (int i = 0; i < count; i++) cells.append(text(prefix + i));
        return cells.toString();
    }

    @Test
    void theLastExcelColumnIsAllowedAndOneBeyondItIsNot() throws IOException {
        // 16,384 columns (A..XFD) fit; the 16,385th cell of a row without references does not
        List<String> seen = new ArrayList<>();
        ExcelReader.<String>mapping(row -> String.valueOf(row.size()))
                .read(new ByteArrayInputStream(withSheetData("<row>" + cells(16_384, "h") + "</row><row>"
                        + cells(16_384, "v") + "</row>")), result -> seen.add(result.data()));
        assertEquals(List.of("16384"), seen);

        byte[] tooWide = withSheetData("<row>" + cells(16_384, "h") + text("one more") + "</row><row>" + text("v") + "</row>");
        ExcelReadException error = assertThrows(ExcelReadException.class,
                () -> ExcelReader.forMap().read(new ByteArrayInputStream(tooWide), result -> { }));
        assertTrue(rootMessage(error).contains("maximum column (XFD)"), rootMessage(error));
    }

    private static String rootMessage(Throwable error) {
        StringBuilder messages = new StringBuilder();
        for (Throwable cause = error; cause != null; cause = cause.getCause()) messages.append(cause.getMessage()).append(" / ");
        return messages.toString();
    }
}
