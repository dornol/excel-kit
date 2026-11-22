package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * A column whose header is blank has no name to look it up by, so readers leave it out of the header
 * index: it is not a key of a {@code forMap()} row, and it can neither collide with another blank
 * column nor shadow a named one. The same holds for CSV and Excel.
 */
class BlankHeaderTest {

    private static List<Map<String, String>> csvMap(String text,
            Function<CsvReader<Map<String, String>>, CsvReader<Map<String, String>>> configure) {
        List<Map<String, String>> rows = new ArrayList<>();
        configure.apply(CsvReader.forMap()).read(new ByteArrayInputStream(text.getBytes(StandardCharsets.UTF_8)),
                result -> rows.add(result.data()));
        return rows;
    }

    private static List<Map<String, String>> csvMap(String text) {
        return csvMap(text, reader -> reader);
    }

    /** Cells that are {@code null} are not written, which is how a blank header cell looks in a file. */
    private static byte[] workbook(String[]... rows) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet("Sheet1");
            for (int r = 0; r < rows.length; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < rows[r].length; c++) {
                    if (rows[r][c] != null) row.createCell(c).setCellValue(rows[r][c]);
                }
            }
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static List<Map<String, String>> excelMap(byte[] workbook) {
        List<Map<String, String>> rows = new ArrayList<>();
        ExcelReader.forMap().read(new ByteArrayInputStream(workbook), result -> rows.add(result.data()));
        return rows;
    }

    @Test
    void severalBlankHeadersNoLongerCollideOnAnEmptyKey() throws IOException {
        assertEquals(List.of(Map.of("Name", "A")), csvMap("Name,,\nA,1,2\n"));
        assertEquals(List.of(Map.of("Name", "A")),
                excelMap(workbook(new String[]{"Name", null, null}, new String[]{"A", "1", "2"})));
    }

    @Test
    void aBlankHeaderBetweenNamedOnesKeepsTheOthersAligned() throws IOException {
        assertEquals(List.of(Map.of("Name", "A", "City", "Seoul")), csvMap("Name,,City\nA,1,Seoul\n"));
        assertEquals(List.of(Map.of("Name", "A", "City", "Seoul")),
                excelMap(workbook(new String[]{"Name", null, "City"}, new String[]{"A", "1", "Seoul"})));
    }

    @Test
    void aBlankHeaderIsLeftOutEvenWhenTheLastDuplicateWins() {
        assertEquals(List.of(Map.of("Name", "A")),
                csvMap("Name,,\nA,1,2\n", reader -> reader.duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST)));
    }

    @Test
    void strictHeadersDoesNotMakeABlankHeaderAnError() {
        assertEquals(List.of(Map.of("Name", "A")), csvMap("Name,\nA,1\n", CsvReader::strictHeaders));
    }

    @Test
    void whenEveryHeaderIsBlankTheRowsHaveNoEntries() {
        assertEquals(List.of(Map.of(), Map.of()), csvMap(",\nA,1\nB,2\n"));
    }

    @Test
    void aNormalizerThatNamesBlankHeadersKeepsThoseColumnsLookUpable() {
        List<String> seen = new ArrayList<>();
        CsvReader.<String>mapping(row -> row.get("Name").asString() + "|" + row.get("unnamed").asString())
                .headerNormalizer(header -> header.isBlank() ? "unnamed" : header)
                .read(new ByteArrayInputStream("Name,\nA,1\n".getBytes(StandardCharsets.UTF_8)),
                        result -> seen.add(result.data()));

        assertEquals(List.of("A|1"), seen);
    }

    @Test
    void columnsReadByPositionDoNotDependOnTheirHeaders() {
        List<String> seen = new ArrayList<>();
        CsvReader.<String[]>setter(() -> new String[2])
                .column((row, cell) -> row[0] = cell.asString())
                .column((row, cell) -> row[1] = cell.asString())
                .read(new ByteArrayInputStream(",\nA,1\n".getBytes(StandardCharsets.UTF_8)),
                        result -> seen.add(result.data()[0] + "|" + result.data()[1]));

        assertEquals(List.of("A|1"), seen);
    }

    @Test
    void mappingReadersCanStillReachABlankHeaderedColumnByIndex() {
        List<String> seen = new ArrayList<>();
        CsvReader.<String>mapping(row -> row.get(0).asString() + "|" + row.get(1).asString())
                .read(new ByteArrayInputStream("Name,\nA,1\n".getBytes(StandardCharsets.UTF_8)),
                        result -> seen.add(result.data()));

        assertEquals(List.of("A|1"), seen);
    }

    @Test
    void aNamedColumnIsStillFoundNextToABlankOne() {
        List<String> seen = new ArrayList<>();
        CsvReader.<String[]>setter(() -> new String[1])
                .column("City", (row, cell) -> row[0] = cell.asString())
                .read(new ByteArrayInputStream("Name,,City\nA,1,Seoul\n".getBytes(StandardCharsets.UTF_8)),
                        result -> seen.add(result.data()[0]));

        assertEquals(List.of("Seoul"), seen);
    }
}
